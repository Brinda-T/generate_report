
epics_column_names = {
    "S.No": {"pos": 0},
    "Custom field (Epic link)": {"pos": 1},
    "Ename": {"pos": 2},
    "Esdate": {"pos": 3},
    "Etdate": {"pos": 4},
    "EASdate": {"pos": 5},
    "EAEdate": {"pos":6}
}

stories_column_names = {
    "S.No": {"pos": 0, "gen_pos": {"column_name": "S.No"}},
    "issue key": {"pos": 1, "gen_pos": {"column_name":"EPIC ID"}},
    "issue id": {"pos": 2, "gen_pos": {"column_name": "SPRINT ID"}},
    "custom field (epic lnk)": {"pos": 3, "gen_pos": {"column_name": "STORY ID"}},
    "ename": {"pos": 4, "gen_pos": {"column_name": "STORY DESCRIPTION"}},
    "assignee": {"pos":5, "gen_pos": {"column_name": "ASSIGNED TO"}},
    "custom field (story points)": {"pos": 6, "gen_pos":{"column_name": "ESTIMATED\nSTORY\nPOINTS"}},
    "teste": {"pos": 7, "gen_pos": {"column_name": "STORY\nPLANNED\nSTART"}},
    "original estimate": {"pos": 8, "gen_pos": {"column_name": "STORY\nPLANNED\nEND"}},
    "time spent": {"pos": 9, "gen_pos": {"column_name": "ESTIMATED\nEFFORT\n(IN HRS)"}},
    "remaining estimate": {"pos":10, "gen_pos": {"column_name": "TIME SPENT\n(IN SECS)"}},
    "sprint": {"pos": 11, "gen_pos": {"column_name": "EFFORT\nCONSUMED\n(IN HRS)"}},
    "sprint 2": {"pos": 12, "gen_pos": {"column_name": "PENDING\nEFFORT\n(IN HRS)"}},
    "sprint 3": {"pos": 13, "gen_pos": {"column_name": "STORY\nACTUAL\nSTART"}},
    "summary": {"pos": 14, "gen_pos": {"column_name": "STORY\nACTUAL\nEND"}},
    "gen progress": {"column_name": "Story\nEffort\ncompletion\n%","pos": 15,},
    "gen scheduled progress": {"column_name": "Story\nSchedule\nprogress%", "pos": 16},
    "gen scheduled overrun": {"column_name": "Story\nSchedule\nOverrun%", "pos": 17},
    "gen comments": {"column_name": "Remarks", "pos": 18}

} 

def get_sno_pos():
    return stories_column_names["S.No"]["pos"]
def get_sno_column_name():
    return stories_column_names["S.No"]["gen_pos"]["column_name"]

def get_issue_key_pos():
    return stories_column_names["issue key"]["pos"]
def get_issue_key_col_name():
    return stories_column_names["issue key"]["gen_pos"]["column_name"]

def get_issue_id_pos():
    return stories_column_names["issue id"]["pos"]
def get_issue_id_col_name():
    return stories_column_names["issue id"]["gen_pos"]["column_name"]

def get_custom_field_epiclink_pos():
    return stories_column_names["custom field (epic lnk)"]["pos"]
def get_custom_field_epiclink_col_name():
    return stories_column_names["custom field (epic lnk)"]["gen_pos"]["column_name"]

def get_ename_pos():
    return stories_column_names["ename"]["pos"]
def get_ename_col_name():
    return stories_column_names["ename"]["gen_pos"]["column_name"]

def get_assignee_pos():
    return stories_column_names["assignee"]["pos"]
def get_assignee_col_name():
    return stories_column_names["assignee"]["gen_pos"]["column_name"]

def get_custom_field_storypoints_pos():
    return stories_column_names["custom field (story points)"]["pos"]
def get_custom_field_storypoints_col_name():
    return stories_column_names["custom field (story points)"]["gen_pos"]["column_name"]

def get_teste_pos():
    return stories_column_names["teste"]["pos"]
def get_teste_col_name():
    return stories_column_names["teste"]["gen_pos"]["column_name"]

def get_original_estimate_pos():
    return stories_column_names["original estimate"]["pos"]
def get_original_estimate_col_name():
    return stories_column_names["original estimate"]["gen_pos"]["column_name"]

def get_time_spent_pos():
    return stories_column_names["time spent"]["pos"]
def get_time_spent_col_name():
    return stories_column_names["time spent"]["gen_pos"]["column_name"]

def get_remaining_estimate_pos():
    return stories_column_names["remaining estimate"]["pos"]
def get_remaining_estimate_col_name():
    return stories_column_names["remaining estimate"]["gen_pos"]["column_name"]

def get_sprint_pos():
    return stories_column_names["sprint"]["pos"]
def get_sprint_col_name():
    return stories_column_names["sprint"]["gen_pos"]["column_name"]

def get_sprint2_pos():
    return stories_column_names["sprint 2"]["pos"]
def get_sprint2_col_name():
    return stories_column_names["sprint 2"]["gen_pos"]["column_name"]

def get_sprint3_pos():
    return stories_column_names["sprint 3"]["pos"]
def get_sprint3_col_name():
    return stories_column_names["sprint 3"]["gen_pos"]["column_name"]

def get_summary_pos():
    return stories_column_names["summary"]["pos"]
def get_summary_col_name():
    return stories_column_names["summary"]["gen_pos"]["column_name"]

def get_epics_custom_field_pos():
    return epics_column_names["Custom field (Epic link)"]["pos"]
def get_epics_Esdate_pos():
    return epics_column_names["Esdate"]["pos"]
def get_epics_Etdate_pos():
    return epics_column_names["Etdate"]["pos"]
def get_epics_EASdate_pos():
    return epics_column_names["EASdate"]["pos"]
def get_epics_EAEdate_pos():
    return epics_column_names["EAEdate"]["pos"]


def get_progress_pos ():
    return stories_column_names["gen progress"]["pos"]
def get_progress_gen_column_name():
    return stories_column_names["gen progress"]["column_name"]
def get_scheduled_progress_pos ():
    return stories_column_names["gen scheduled progress"]["pos"]
def get_scheduled_progress_gen_column_name():
    return stories_column_names["gen scheduled progress"]["column_name"]
def get_scheduled_overrun_pos ():
    return stories_column_names["gen scheduled overrun"]["pos"]
def get_scheduled_overrun_gen_column_name():
    return stories_column_names["gen scheduled overrun"]["column_name"]
def get_remarks_pos ():
    return stories_column_names["gen comments"]["pos"]
def get_remarks_gen_column_name():
    return stories_column_names["gen comments"]["column_name"]

def get_column_names_and_positions():
    return get_sno_pos(),get_issue_key_pos(),get_issue_key_col_name(),get_issue_id_pos(), get_issue_id_col_name(), get_custom_field_epiclink_pos(), get_custom_field_epiclink_col_name(), get_ename_pos(), get_ename_col_name(), get_assignee_pos(), get_assignee_col_name(), get_custom_field_storypoints_pos(), get_custom_field_storypoints_col_name(), get_teste_pos(), get_teste_col_name(), get_original_estimate_pos(), get_original_estimate_col_name(), get_time_spent_pos(), get_time_spent_col_name(), get_remaining_estimate_pos(), get_remaining_estimate_col_name(), get_sprint_pos(), get_sprint_col_name(), get_sprint2_pos(), get_sprint2_col_name(), get_sprint3_pos(), get_sprint3_col_name(), get_summary_pos(), get_summary_col_name(), get_epics_custom_field_pos(), get_epics_Esdate_pos(), get_epics_Etdate_pos(), get_epics_EASdate_pos(), get_epics_EAEdate_pos(), get_progress_pos(), get_progress_gen_column_name(), get_scheduled_progress_pos(), get_scheduled_progress_gen_column_name(), get_scheduled_overrun_pos(), get_scheduled_overrun_gen_column_name(), get_remarks_pos(), get_remarks_gen_column_name()
