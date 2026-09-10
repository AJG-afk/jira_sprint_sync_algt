#!/usr/bin/env python3
"""
lo_chart_export.py — export every chart on a sheet to PNG via LibreOffice UNO (Linux/sandbox path).
On macOS the production path is Excel+AppleScript (see chart_export.py). Returns {index: png_path}.

Starts a headless soffice listener, runs a UNO script under LibreOffice's bundled Python, exports
each OLE2/Chart shape on the sheet in draw-page order.
"""
import subprocess, time, os, glob, shutil

UNO_TEMPLATE = r'''
import uno, time
from com.sun.star.beans import PropertyValue
def mk(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
ctx0 = uno.getComponentContext()
res = ctx0.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx0)
ctx = None
for _ in range(30):
    try:
        ctx = res.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception:
        time.sleep(1)
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desktop.loadComponentFromURL("file://__XLSX__", "_blank", 0, (mk("Hidden", True),))
sheet = doc.Sheets.getByName("__SHEET__")
exp = smgr.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", ctx)
i = 0
for shape in sheet.DrawPage:
    if any("OLE2Shape" in s or "Chart" in s for s in shape.SupportedServiceNames):
        exp.setSourceDocument(shape)
        exp.filter((mk("URL", "file://__OUT__/__SHEET___chart%d.png" % i), mk("MediaType", "image/png"),
                    mk("FilterData", uno.Any("[]com.sun.star.beans.PropertyValue",
                        (mk("PixelWidth", 900), mk("PixelHeight", 520))))))
        i += 1
print("EXPORTED", i)
doc.close(False)
'''

# Candidate paths for LibreOffice's bundled python (needed for `import uno`).
_LO_PY_CANDIDATES = [
    "/opt/libreoffice25.2/program/python",
    "/opt/libreoffice26.2/program/python",
    "/usr/lib/libreoffice/program/python",
]


def _lo_python():
    for p in _LO_PY_CANDIDATES:
        if os.path.exists(p):
            return p
    return "python3"


def export_sheet_charts(xlsx, sheet, out_dir, port=2002):
    os.makedirs(out_dir, exist_ok=True)
    xabs = os.path.abspath(xlsx)
    oabs = os.path.abspath(out_dir)
    script = (UNO_TEMPLATE.replace("__XLSX__", xabs)
                          .replace("__OUT__", oabs)
                          .replace("__SHEET__", sheet))
    sp = os.path.join(out_dir, f"_uno_{sheet}.py")
    with open(sp, "w") as f:
        f.write(script)

    subprocess.run(["pkill", "-f", "soffice"], capture_output=True)
    time.sleep(2)
    soffice = shutil.which("soffice") or "soffice"
    proc = subprocess.Popen([soffice, "--headless", "--norestore", "--invisible",
                             f"--accept=socket,host=localhost,port={port};urp;"])
    time.sleep(10)
    try:
        subprocess.run([_lo_python(), sp], capture_output=True, text=True, timeout=120)
    finally:
        proc.terminate()
        subprocess.run(["pkill", "-f", "soffice"], capture_output=True)

    out = {}
    for p in sorted(glob.glob(os.path.join(out_dir, f"{sheet}_chart*.png"))):
        idx = int(p.rsplit("chart", 1)[1].split(".")[0])
        out[idx] = p
    return out


def export_many(xlsx, sheets, out_root, port=2002):
    """Export charts for multiple sheets in ONE soffice session (efficient). {sheet:{idx:png}}."""
    os.makedirs(out_root, exist_ok=True)
    xabs = os.path.abspath(xlsx)
    blocks = []
    for sheet in sheets:
        d = os.path.abspath(os.path.join(out_root, sheet))
        os.makedirs(d, exist_ok=True)
        blocks.append((sheet, d))
    runner = os.path.join(out_root, "_uno_many.py")
    body = ['import uno, time, os',
            'from com.sun.star.beans import PropertyValue',
            'def mk(n,v):',
            '    p=PropertyValue(); p.Name=n; p.Value=v; return p',
            'ctx0=uno.getComponentContext()',
            'res=ctx0.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",ctx0)',
            'ctx=None',
            'for _ in range(30):',
            '    try:',
            f'        ctx=res.resolve("uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"); break',
            '    except Exception: time.sleep(1)',
            'smgr=ctx.ServiceManager',
            'desktop=smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)',
            f'doc=desktop.loadComponentFromURL("file://{xabs}","_blank",0,(mk("Hidden",True),))',
            'exp=smgr.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter",ctx)']
    for sheet, d in blocks:
        body += [f'sheet=doc.Sheets.getByName("{sheet}")', 'i=0',
                 'for shape in sheet.DrawPage:',
                 '    if any("OLE2Shape" in s or "Chart" in s for s in shape.SupportedServiceNames):',
                 '        exp.setSourceDocument(shape)',
                 f'        exp.filter((mk("URL","file://{d}/{sheet}_chart%d.png"%i),mk("MediaType","image/png"),'
                 'mk("FilterData",uno.Any("[]com.sun.star.beans.PropertyValue",(mk("PixelWidth",900),mk("PixelHeight",520))))))',
                 '        i+=1',
                 f'print("{sheet}",i)']
    body += ['doc.close(False)']
    with open(runner, "w") as f:
        f.write("\n".join(body))

    subprocess.run(["pkill", "-f", "soffice"], capture_output=True)
    time.sleep(2)
    soffice = shutil.which("soffice") or "soffice"
    proc = subprocess.Popen([soffice, "--headless", "--norestore", "--invisible",
                             f"--accept=socket,host=localhost,port={port};urp;"])
    time.sleep(10)
    try:
        subprocess.run([_lo_python(), runner], capture_output=True, text=True, timeout=240)
    finally:
        proc.terminate()
        subprocess.run(["pkill", "-f", "soffice"], capture_output=True)

    result = {}
    for sheet, d in blocks:
        m = {}
        for p in sorted(glob.glob(os.path.join(d, f"{sheet}_chart*.png"))):
            idx = int(p.rsplit("chart", 1)[1].split(".")[0])
            m[idx] = p
        result[sheet] = m
    return result
