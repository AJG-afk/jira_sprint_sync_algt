#!/usr/bin/env python3
"""
team_columns.py — per-team column layout + fill policy for the IT OPS Sprint Metrics workbook.
Consumed by jira_sprint_sync.py (--mode excel) to drive excel_writer.fill_sprint_row().

  jira          : {logical field -> column letter}  raw values written from Jira
  carry         : [cols]  repeat last row's value  (Number of People, Days Available)
  zero          : [cols]  write 0                  (Days Worked)
  blank         : [cols]  force manual-empty       (People Change, Unplanned OOO)
  formula       : {col: "AE{r}/AF{r}"}  live formula, {r}->row  (Capacity, shows 0 until Days Worked set)
  story_counts  : {logical field -> column letter}  5 new cols at last-used+5

Column letters VERIFIED from workbook headers (2026-09):
  NET  #People=AB PeopleChg=AC Capacity=AD DaysWorked=AE DaysAvail=AF OOO=AG  Kanban=V/W  Spikes=H/I
  STN  #People=Z  PeopleChg=AA Capacity=AB DaysWorked=AC DaysAvail=AD OOO=AE  Kanban=T/U  ExtImpact=F
  WIR  #People=AA PeopleChg=AB Capacity=AC DaysWorked=AD DaysAvail=AE (no OOO) Kanban=T/U
  JST  #People=Z  PeopleChg=AA Capacity=AB DaysWorked=AC DaysAvail=AD OOO=AE  Kanban=T/U
  TEL  inherits standard (STN/JST) layout — verify TEL tab before first commit
  WTE  deferred (different Kanban structure)

Story-count columns at each tab's last-used + 5:  NET/STN/JST -> AN..AR   WIR -> AO..AS
"""

_STD_STORY = {"stories_committed": "AN", "stories_added": "AO", "stories_completed": "AP",
              "stories_notdone": "AQ", "stories_removed": "AR"}

TEAM_COLUMNS = {
    "NET": {
        "jira": {"sprint": "B", "committed_sp": "C", "completed_sp": "D", "carried_over": "E",
                 "kanban_sp": "V", "kanban_count": "W", "spike_count": "H", "spike_sp": "I"},
        "carry": ["AB", "AF"], "zero": ["AE"], "blank": ["AC", "AG"],
        "formula": {"AD": "AE{r}/AF{r}"},
        "story_counts": dict(_STD_STORY),
    },
    "STN": {
        "jira": {"sprint": "B", "committed_sp": "C", "completed_sp": "D", "carried_over": "E",
                 "kanban_sp": "T", "kanban_count": "U", "external_impact": "F"},
        "carry": ["Z", "AD"], "zero": ["AC"], "blank": ["AA", "AE"],
        "formula": {"AB": "AC{r}/AD{r}"},
        "story_counts": dict(_STD_STORY),
    },
    "WIR": {
        "jira": {"sprint": "B", "committed_sp": "C", "completed_sp": "D", "carried_over": "E",
                 "kanban_sp": "T", "kanban_count": "U"},
        "carry": ["AA", "AE"], "zero": ["AD"], "blank": ["AB"],
        "formula": {"AC": "AD{r}/AE{r}"},
        "story_counts": {"stories_committed": "AO", "stories_added": "AP", "stories_completed": "AQ",
                         "stories_notdone": "AR", "stories_removed": "AS"},
    },
    "JST": {
        "jira": {"sprint": "B", "committed_sp": "C", "completed_sp": "D", "carried_over": "E",
                 "kanban_sp": "T", "kanban_count": "U"},
        "carry": ["Z", "AD"], "zero": ["AC"], "blank": ["AA", "AE"],
        "formula": {"AB": "AC{r}/AD{r}"},
        "story_counts": dict(_STD_STORY),
    },
    "TEL": {  # VERIFY TEL tab before first commit
        "jira": {"sprint": "B", "committed_sp": "C", "completed_sp": "D", "carried_over": "E",
                 "kanban_sp": "T", "kanban_count": "U"},
        "carry": ["Z", "AD"], "zero": ["AC"], "blank": ["AA", "AE"],
        "formula": {"AB": "AC{r}/AD{r}"},
        "story_counts": dict(_STD_STORY),
    },
    # WTE deferred.
}
