import csv
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Fill, colors, PatternFill, Protection, Alignment
from openpyxl.styles import NamedStyle
import columns
import datetime
from sr_log import sr_log_messages, sr_debug

sno_pos = columns.get_sno_pos()
sr_debug("sno_pos", sno_pos)

issue_key_pos = columns.get_issue_key_pos()
sr_debug ("issue_key_pos", issue_key_pos)

issue_key_col_name = columns.get_issue_key_col_name()
sr_debug ("issue_key_col_name", issue_key_col_name)

issue_id_pos = columns.get_issue_id_pos()
sr_debug ("issue_id_pos", issue_id_pos)

issue_id_col_name = columns.get_issue_id_col_name()
sr_debug ("issue_id_col_name", issue_id_col_name)
custom_field_epiclink_pos = columns.get_custom_field_epiclink_pos()
sr_debug ("custom_field_epiclink_pos", custom_field_epiclink_pos)
custom_field_epiclink_col_name = columns.get_custom_field_epiclink_col_name()
sr_debug ("custom_field_epiclink_col_name", custom_field_epiclink_col_name)
ename_pos = columns.get_ename_pos()
sr_debug ("ename_pos", ename_pos)
ename_col_name = columns.get_ename_col_name()
sr_debug ("ename_col_name", ename_col_name)
assignee_pos = columns.get_assignee_pos()
sr_debug ("assignee_pos", assignee_pos)
assignee_col_name = columns.get_assignee_col_name()
sr_debug ("assignee_col_name", assignee_col_name)
custom_field_storypoints_pos = columns.get_custom_field_storypoints_pos()
sr_debug ("custom_field_storypoints_pos", custom_field_storypoints_pos)
custom_field_storypoints_col_name = columns.get_custom_field_storypoints_col_name()
sr_debug ("custom_field_storypoints_col_name", custom_field_storypoints_col_name)
teste_pos = columns.get_teste_pos()
sr_debug ("teste_pos", teste_pos)
teste_col_name = columns.get_teste_col_name()
sr_debug ("teste_col_name", teste_col_name)
original_estimate_pos = columns.get_original_estimate_pos()
sr_debug ("original_estimate_pos", original_estimate_pos)
original_estimate_col_name = columns.get_original_estimate_col_name()
sr_debug ("original_estimate_col_name", original_estimate_col_name)
time_spent_pos = columns.get_time_spent_pos()
sr_debug ("time_spent_pos", time_spent_pos)
time_spent_col_name = columns.get_time_spent_col_name()
sr_debug ("time_spent_col_name", time_spent_col_name)
remaining_estimate_pos = columns.get_remaining_estimate_pos()
sr_debug ("remaining_estimate_pos", remaining_estimate_pos)
remaining_estimate_col_name = columns.get_remaining_estimate_col_name()
sr_debug ("remaining_estimate_col_name", remaining_estimate_col_name)
sprint_pos = columns.get_sprint_pos()
sr_debug ("sprint_pos", sprint_pos)
sprint_col_name = columns.get_sprint_col_name()
sr_debug ("sprint_col_name", sprint_col_name)
sprint2_pos = columns.get_sprint2_pos()
sr_debug ("sprint2_pos", sprint2_pos)
sprint2_col_name = columns.get_sprint2_col_name()
sr_debug ("sprint2_col_name", sprint2_col_name)
sprint3_pos = columns.get_sprint3_pos()
sr_debug ("sprint3_pos", sprint3_pos)
sprint3_col_name = columns.get_sprint3_col_name()
sr_debug ("sprint3_col_name", sprint3_col_name)
summary_pos = columns.get_summary_pos()
sr_debug ("summary_pos", summary_pos)
summary_col_name = columns.get_summary_col_name()
sr_debug ("summary_col_name", summary_col_name)

sr_debug ("------------------")

epics_custom_field_pos = columns.get_epics_custom_field_pos()
sr_debug ("epics_custom_field_pos", epics_custom_field_pos)     
epics_Esdate_pos = columns.get_epics_Esdate_pos()
sr_debug ("epics_Esdate_pos", epics_Esdate_pos)
epics_Etdate_pos = columns.get_epics_Etdate_pos()
sr_debug ("epics_Etdate_pos", epics_Etdate_pos)
epics_EASdate_pos = columns.get_epics_EASdate_pos()
sr_debug ("epics_EASdate_pos", epics_EASdate_pos)
epics_EAEdate_pos = columns.get_epics_EAEdate_pos()
sr_debug ("epics_EAEdate_pos", epics_EAEdate_pos)

sr_debug ("--------------------------")

progress_pos = columns.get_progress_pos()
sr_debug ("progress_pos", progress_pos)
progress_gen_column_name = columns.get_progress_gen_column_name()
sr_debug ("progress_gen_column_name", progress_gen_column_name)

scheduled_progress_pos = columns.get_scheduled_progress_pos()
sr_debug ("scheduled_progress_pos", scheduled_progress_pos)
scheduled_progress_gen_column_name = columns.get_scheduled_progress_gen_column_name()
sr_debug ("scheduled_progress_gen_column_name", scheduled_progress_gen_column_name)

scheduled_overrun_pos = columns.get_scheduled_overrun_pos()
sr_debug ("scheduled_overrun_pos", scheduled_overrun_pos)
scheduled_overrun_gen_column_name = columns.get_scheduled_overrun_gen_column_name()
sr_debug ("scheduled_overrun_gen_column_name", scheduled_overrun_gen_column_name)

remarks_pos = columns.get_remarks_pos()
sr_debug ("remarks_pos", remarks_pos)
remarks_gen_column_name = columns.get_remarks_gen_column_name()
sr_debug ("remarks_gen_column_name", remarks_gen_column_name)


def create_workbook(wbook):
   
    try:
        wb = load_workbook(wbook)
        print("Workbook '%s'exists" %(wb))
    except:
        print("Creating worksheet: '%s'" %(wbook))
        wb = Workbook()

    wb.save(wbook)

def get_read_csv_files(file_name):
    datalist = []
    fd = open(file_name)

    fname = csv.reader(fd)
    for row in fname:
        datalist.append(row)
    return datalist

def create_worksheet(wbname, dst_wname):

    wb = load_workbook(wbname)


    try:
        dwsheet = wb.get_sheet_by_name(dst_wname)
        sr_debug ("worksheet '%s' found"%(dst_wname))
        sr_debug ("removing worksheet:'%s'"%(dst_wname))
        wb.remove_sheet(dwsheet)
    except:
        sr_debug ("worksheet '%s' not found"%(dst_wname))
    finally:
        sr_debug ("creating new worksheet:'%s'"%(dst_wname))
        dwsheet = wb.create_sheet(dst_wname,0)
	
    wb.save(wbname)
def print_row_column1(wbname,stories,epics,dst_wname):
    dates_list = []
    sprint_list = []		
    wb = load_workbook(wbname)
    dwsheet = wb.get_sheet_by_name(dst_wname)

    srcount1 = len(stories)
    
    sr_debug ("stories length :%d)"%(srcount1))

    srcount2 = len(epics)
    
    sr_debug ("epics length :%d)"%(srcount2))
    
    dwsheet.insert_cols(1, 18)
    #drcount =  dwsheet.max_row
    #dccount =  dwsheet.max_column
    #sr_debug("%s: max row:col (%d:%d)" % (dst_wname, drcount, dccount))
    '''
    row = stories[0]       
    for frow in row:
        dwsheet.append(frow)
    '''
    titles_center_aligned_text = Alignment(horizontal = "center", vertical = "center", wrapText=True)
    center_aligned_text = Alignment(horizontal = "center", vertical = "center")

    '''     
    dwsheet['A1'] = "sai\ndheeraj"
    dwsheet['A1'].alignment = Alignment(horizontal = "center", vertical = "center", wrapText=True)
    #dwsheet['A1'].alignment = center_aligned_text
    wb.save("wrap.xlsx")
    '''

    header_row = dwsheet[1]
    dwsheet[1][issue_key_pos].value = issue_key_col_name
    dwsheet[1][issue_key_pos].alignment = titles_center_aligned_text

    dwsheet[1][issue_id_pos].value = issue_id_col_name
    dwsheet[1][issue_id_pos].alignment = titles_center_aligned_text

    dwsheet[1][custom_field_epiclink_pos].value = custom_field_epiclink_col_name
    dwsheet[1][custom_field_epiclink_pos].alignment = titles_center_aligned_text

    dwsheet[1][ename_pos].value = ename_col_name
    dwsheet[1][ename_pos].alignment = titles_center_aligned_text

    dwsheet[1][assignee_pos].value = assignee_col_name
    dwsheet[1][assignee_pos].alignment = titles_center_aligned_text

    dwsheet[1][custom_field_storypoints_pos].value = custom_field_storypoints_col_name
    dwsheet[1][custom_field_storypoints_pos].alignment = titles_center_aligned_text

    dwsheet[1][teste_pos].value = teste_col_name
    dwsheet[1][teste_pos].alignment = titles_center_aligned_text

    dwsheet[1][original_estimate_pos].value = original_estimate_col_name
    dwsheet[1][original_estimate_pos].alignment = titles_center_aligned_text

    dwsheet[1][time_spent_pos].value = time_spent_col_name
    dwsheet[1][time_spent_pos].alignment = titles_center_aligned_text

    dwsheet[1][remaining_estimate_pos].value = remaining_estimate_col_name
    dwsheet[1][remaining_estimate_pos].alignment = titles_center_aligned_text

    dwsheet[1][sprint_pos].value = sprint_col_name
    dwsheet[1][sprint_pos].alignment = titles_center_aligned_text

    dwsheet[1][sprint2_pos].value = sprint2_col_name
    dwsheet[1][sprint2_pos].alignment = titles_center_aligned_text

    dwsheet[1][sprint3_pos].value = sprint3_col_name
    dwsheet[1][sprint3_pos].alignment = titles_center_aligned_text

    dwsheet[1][summary_pos].value = summary_col_name
    dwsheet[1][summary_pos].alignment = titles_center_aligned_text

    dwsheet['p1'] = progress_gen_column_name
    dwsheet['p1'].alignment = titles_center_aligned_text

    dwsheet['q1'] = scheduled_progress_gen_column_name
    dwsheet['q1'].alignment = titles_center_aligned_text

    dwsheet['r1'] = scheduled_overrun_gen_column_name
    dwsheet['r1'].alignment = titles_center_aligned_text

    dwsheet['s1'] = remarks_gen_column_name
    dwsheet['s1'].alignment = titles_center_aligned_text

    for j in range (1, srcount1):
        
        i = j + 1

        sprint_list.append(stories[j][sprint_pos])
        sprint_list.append(stories[j][sprint2_pos])
        sprint_list.append(stories[j][sprint3_pos])
        sprint_list = list(filter(None, sprint_list))
        sprint_list.sort(reverse = True)
        #sr_debug (sprint_list)
        dwsheet[i][issue_id_pos].value = sprint_list[0]
        sprint_list = []

    for i in range (1, srcount1):
        j = i + 1
        dwsheet[j][sno_pos].value = stories[i][sno_pos]
        dwsheet[j][sno_pos].alignment = center_aligned_text

        dwsheet[j][issue_key_pos].value = stories[i][custom_field_epiclink_pos]
        dwsheet[j][issue_key_pos].alignment = center_aligned_text

        dwsheet[j][custom_field_epiclink_pos].value = stories[i][issue_key_pos]
        dwsheet[j][custom_field_epiclink_pos].alignment = center_aligned_text

        dwsheet[j][ename_pos].value = stories[i][summary_pos]
        dwsheet[j][ename_pos].alignment = center_aligned_text

        dwsheet[j][assignee_pos].value = stories[i][assignee_pos]
        dwsheet[j][assignee_pos].alignment = center_aligned_text

        dwsheet[j][custom_field_storypoints_pos].value = stories[i][custom_field_storypoints_pos]
        dwsheet[j][custom_field_storypoints_pos].alignment = center_aligned_text

        estimate_effort_hours = int(stories[i][original_estimate_pos]) / 3600
        dwsheet[j][time_spent_pos].value = estimate_effort_hours
        dwsheet[j][time_spent_pos].alignment = center_aligned_text

        dwsheet[j][remaining_estimate_pos].value = stories[i][time_spent_pos]
        dwsheet[j][remaining_estimate_pos].alignment = center_aligned_text

        effort_consumed_hours = int(stories[i][time_spent_pos]) / 3600
        dwsheet[j][sprint_pos].value = effort_consumed_hours
        dwsheet[j][sprint_pos].alignment = center_aligned_text

        pending_effort_hours = estimate_effort_hours - effort_consumed_hours
        dwsheet[j][sprint2_pos].value = pending_effort_hours
        dwsheet[j][sprint2_pos].alignment = center_aligned_text

        effort_completion = (effort_consumed_hours/estimate_effort_hours) * 100
        dwsheet[j][progress_pos].value = ('{}{}'.format(round(effort_completion),"%"))
        dwsheet[j][progress_pos].alignment = center_aligned_text

        scheduled_progress = 100
        dwsheet[j][scheduled_progress_pos].value = ('{}{}'.format(scheduled_progress,"%"))
        dwsheet[j][scheduled_progress_pos].alignment = center_aligned_text

        scheduled_overrun = 0
        dwsheet[j][scheduled_overrun_pos].value = ('{}{}'.format(scheduled_overrun,"%"))
        dwsheet[j][scheduled_overrun_pos].alignment = center_aligned_text


    for k in range (1, srcount2):
        for l in range (1, srcount1):
            i = l + 1
            dates_list.append(epics[k][epics_Esdate_pos])
            dates_list.append(epics[k][epics_Etdate_pos])
            dates_list.append(epics[k][epics_EASdate_pos])
            dates_list.append(epics[k][epics_EAEdate_pos])
            dates_list.sort()

            if (epics[k][epics_custom_field_pos]):
                stories[l][custom_field_epiclink_pos]

            dwsheet[i][teste_pos].value = (dates_list[0])
            dwsheet[i][teste_pos].alignment = center_aligned_text
            #sr_debug (dates_list[0])
            #sr_debug (type(dates_list[0]))

            dwsheet[i][original_estimate_pos].value = (dates_list[2])
            dwsheet[i][original_estimate_pos].alignment = center_aligned_text

            dwsheet[i][sprint3_pos].value = (dates_list[1])
            dwsheet[i][sprint3_pos].alignment = center_aligned_text

            dwsheet[i][summary_pos].value = (dates_list[3])
            dwsheet[i][summary_pos].alignment = center_aligned_text
            #sr_debug (dwsheet[l][original_estimate_pos].value)



    drcount =  dwsheet.max_row
    dccount =  dwsheet.max_column
    sr_debug("%s: max row:col (%d:%d)" % (dst_wname, drcount, dccount))

    sr_debug("Saving :%s" % (wbname))
    wb.save("employee-details.xlsx")
    wb.save(wbname)


def main():
    filename = "employee-details.xlsx"
    stories = "02-stories.csv"
    epics = "01-epics.csv"
    dst_wname = "generated"
    create_workbook(filename)
    epics_data = get_read_csv_files(epics)
    stories_data = get_read_csv_files(stories)
    create_worksheet(filename, dst_wname)
    #print_all_worksheet_names(filename)
    print_row_column1(filename, stories_data, epics_data, dst_wname)

if (__name__ == '__main__'):
    main()

