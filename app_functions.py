from io import BytesIO
import pandas as pd
from table_classes import bundle_rows, audit_rows
from datetime import datetime
import xlsxwriter
from jira import JIRA
import os
import sys
from dotenv import load_dotenv
sys.path.insert(1,os.getcwd() +"\pages")

load_dotenv()

bundle_rows_incl = []
bundle_rows_off = []
bundle_rows_data = []
bundle_rows_org = []
phire_audit_list = []
fiscal_year_1 = '2022'
fiscal_year_2 = '2023'
test_jira = ''
report_title = ''
'''
bundle_dates = ["11/01/2021", "05/02/2021", "10/03/2021", "04/04/2021", "09/05/2021", "02/07/2021",
 "03/07/2021", "08/08/2021", "01/10/2021", "12/13/2021", "11/14/2021", "10/18/2021", "09/24/2021",
  "07/25/2021", "12/27/2021", "06/27/2021", "11/29/2021", "05/30/2021", "01/10/2022", "01/24/2022",
  "02/07/2022", "02/21/2022", "03/07/2022", "03/21/2022", "04/04/2022", "04/18/2022", "05/02/2022",
  "05/16/2022", "05/31/2022", "06/13/2022", "06/27/2022", "07/11/2022", "07/25/2022", "08/08/2022",
  "08/22/2022", "09/06/2022", "09/23/2022", "10/03/2022", "10/17/2022", "10/31/2022", "11/13/2022",
  "11/28/2022","12/12/2022","12/27/2022", "01/09/2023","01/23/2023","02/06/2023"]
'''

bundle_dates = pd.read_csv("./pages/dates.csv").iloc[:,0].tolist()

jira_obj = JIRA(os.getenv("HOST_URL"),basic_auth=(os.getenv("JIRA_USER_ID"),os.getenv("JIRA_API_KEY")))

bundle_header = ["JIRA #", "Summary", "Team", "Bundle Status", "JIRA Status", "Assignee",
 "Reporter", "Priority", "Sub-Priority", "Category", "Prioritization Date", "HRS/EPM", "HRQA/EPQAS Date",
 "PHIRE CR#", "CR Type", "HRS/EPM Date", "Off Bundle Reason", "Project Association"]

audit_header = ["Phire CR Number", "JIRA Tracking #", "Target DB", "Migrated On", "Migrated By", "CR Type", "Query"]

def classify_category(project_title):
    project_type = ""
    if(project_title == "No"):
        project_type = "Ops"
    else:
        project_type = "Project"
    return project_type

def classify_crtype(phire_cr_type):
    bundle_crtype = ""
    if(phire_cr_type == "HFIX"):
        bundle_crtype = "Data Update/SQL"
    elif(phire_cr_type == "SCRP"):
        bundle_crtype = "Config/.dms"
    elif(phire_cr_type == "SCRT"):
        bundle_crtype = "Security"
    elif(phire_cr_type == "CODE"):
        bundle_crtype = "Code/Object"
    elif(phire_cr_type == "N/A"):
        bundle_crtype = "N/A (Configuration)"
    return bundle_crtype


def process_jira(jira_obj, row, bundle_status):
    try:
        issues = jira_obj.issue(row["Tracking #"])
        #print(issues.raw['fields'])
    except:
        return None

    #processing jira-id of the JIRA
    try:
        jira_code = issues.key if issues.key is not None else "-"
    except:
        jira_code = "-"

    #processing summary (title) of the JIRA 
    try:
        summary = issues.fields.summary if issues.fields.summary is not None else "-"
    except:
        summary = "-"

    #processing team responsible for the JIRA 
    try:
        team = issues.fields.customfield_10085.value if issues.fields.customfield_10085 is not None else "-"
    except:
        team = "-"

    #processing bundle status of JIRA given in JIRA
    try:
        jira_status = issues.fields.status.name if issues.fields.status is not None else "-"
    except:
        jira_status = "-"
    
    try:
        assignee = issues.fields.assignee.displayName if issues.fields.assignee is not None else "-"
    except:
        assignee = "-"

    try:
        reporter = issues.fields.creator.displayName if issues.fields.creator is not None else "-"
    except:
        reporter = "-"
    
    try:
        priority = issues.fields.priority.name if issues.fields.priority is not None else "-"
    except:
        priority = "-"

    try:
        sub_priority = issues.fields.customfield_10332.value if issues.fields.customfield_10332 is not None else "-"
    except:
        sub_priority = "-"

    try:
        category = classify_category(issues.fields.customfield_10482.value) if issues.fields.customfield_10482 is not None else "-"
    except:
        category = "-"

    try:
        priortization_date_string = issues.fields.customfield_13090 if issues.fields.customfield_13090 is not None else "-"
        priortization_date_obj = datetime.strptime(priortization_date_string, "%Y-%m-%d")
        priortization_date = priortization_date_obj.strftime("%m/%d/%Y")
    except:
        priortization_date = '-'
    
    try:
        isHRS = row['Target DB']
    except:
        isHRS = "-"

    try:
        hrqa_epqas_date = "-"
        comments = issues.fields.comment.comments
        for comment in comments:
            if ('Team: HRS Migration' in comment.author.displayName or 'jira_doit' in comment.author.displayName) and ('to HRQA / HRTRN is complete' in comment.body or 'EPQAS is complete' in comment.body):
                hrqa_epqas_date_string = comment.created[:10]
                hrqa_epqas_date_obj = datetime.strptime(hrqa_epqas_date_string, '%Y-%m-%d')
                hrqa_epqas_date = hrqa_epqas_date_obj.strftime('%m/%d/%Y')
        if bundle_status=='Org Dept Update':
            hrqa_epqas_date = 'N/A'
    except:
        hrqa_epqas_date = "-"

    try:
        if isHRS == "HRS" and bundle_status != 'Org Dept Update':
            phire_cr = str("HRS-" + row["CR"])
        elif isHRS == "EPM":
            phire_cr = str("EPM-" + row["CR"]) 
        else:
            phire_cr = row["CR"]
    except:
        phire_cr = "-"

    try:
        off_reason = issues.fields.customfield_11693.value if issues.fields.customfield_11693 is not None else "-"
        if(bundle_status != 'Included in Bundle'):
            off_reason = '-'
    except:
        off_reason = "-"
    
    try:
        project_association = issues.fields.customfield_13390.value if issues.fields.customfield_13390 is not None else "-"
    except:
        project_association = "-"
    
    cr_type = classify_crtype(row["CR Type"])
    hrs_epm_date = row["Migrated On"].split(" ")[0] if bundle_status != "Org Dept Update" else row["Migrated On"].split("T")[0]

    new_bundle_obj = bundle_rows(jira_code, summary, team, bundle_status, jira_status, assignee, reporter, priority,
     sub_priority, category, priortization_date, isHRS, hrqa_epqas_date, phire_cr, cr_type, hrs_epm_date, off_reason,
     project_association)
    
    return new_bundle_obj

def processCSV(data,type):
    count = 0
    #segregating bundle data into lists based on on-bundle, off-bundle, data update
    total_count = data.shape[0]
    if(type == "SQL Data"):
        data = data.rename(columns={"Action Date": "Migrated On", "DB ID": "Target DB", "Migration By":"Migrated By"})
    for index, row in data.iterrows():
        if row["Migrated On"].split(" ")[0] in bundle_dates:
            result = process_jira(jira_obj, row, "Included in Bundle")
            if(result == None):
                print("Problem with ", row["Tracking #"])
                continue
            else:
                bundle_rows_incl.append(result.toList())
        elif row["CR Type"] == "HFIX":
            result = process_jira(jira_obj, row, "Data Update")
            if(result == None):
                print("Problem with ", row["Tracking #"])
                continue
            else:
                bundle_rows_data.append(result.toList())
        else:
            result = process_jira(jira_obj, row, "Off-bundle")
            if(result == None):
                print("Problem with ", row["Tracking #"])
                continue
            else:
                bundle_rows_off.append(result.toList())
        count += 1
        src_query = 'UW_MIGR_HISTORY_AUDIT_LAB' if type == 'Migration Data' else 'UW_SQL_HISTORY_AUDIT_LAB'
        phire_result = audit_rows(row['CR'], row['Tracking #'], row['Target DB'], row['Migrated On'], row['Migrated By'], row['CR Type'], src_query)
        phire_audit_list.append(phire_result.toList())
        print("Processing ",type,": ",count,"/",total_count, "  Included in Bundle: ", len(bundle_rows_incl), 
        " Off-Bundle: ", len(bundle_rows_off), " Data Update: ", len(bundle_rows_data), " Org Dept Update: ", len(bundle_rows_org))

def processOrg(startDate, endDate, type):
    #Creating a search query of organizational update
    org_issues = jira_obj.search_issues('(labels = FY'+fiscal_year_1+'OrgDept OR labels = FY'+fiscal_year_2+'OrgDept) AND (updated >= ' + startDate.strftime("%Y-%m-%d") + ' AND updated <= ' + endDate.strftime("%Y-%m-%d") + ') ORDER BY updated DESC')
    issue_count = 0
    if(len(org_issues) > 0):
        for issue in org_issues:
            row = {'CR':'N/A - No PHIRE CR', 'Tracking #':issue.key, 'Target DB':'HRS', 'Migrated On':jira_obj.issue(issue.key).fields.updated, 'Migrated By':'N/A', 'CR Type':'N/A'}
            ser = pd.Series(data=row, index=['CR', 'Tracking #', 'Target DB', 'Migrated On', 'Migrated By', 'CR Type'])
            result = process_jira(jira_obj, ser, 'Org Dept Update')
            if(result == None):
                print("Problem with ", row["Tracking #"])
                continue
            else:
                bundle_rows_org.append(result.toList())
            issue_count += 1
            row['Migrated On'] = row['Migrated On'].replace("T", " ")[:19]
            phire_result = audit_rows(row['CR'], row['Tracking #'], row['Target DB'], row['Migrated On'], row['Migrated By'], row['CR Type'], 'ORG_DEPT_UPDATE')
            phire_audit_list.append(phire_result.toList())
            print("Processing ",type + ": ",issue_count,"/",len(org_issues), "  Included in Bundle: ", len(bundle_rows_incl), 
            " Off-Bundle: ", len(bundle_rows_off), " Data Update: ", len(bundle_rows_data), " Org Dept Update: ", len(bundle_rows_org))
    else:
        print("NO ORGANIZATION DEPARTMENT UPDATE AT THIS TIME")

def showBundleSummary():
    summary = pd.DataFrame(columns=bundle_header)
    #Adding Included Row
    for issue in bundle_rows_incl:
        summary.loc[len(summary.index)] = issue
    #Adding Off-bundle Row
    for issue in bundle_rows_off:
        summary.loc[len(summary.index)] = issue
    #Adding Org Dept Updates
    for issue in bundle_rows_org:
        summary.loc[len(summary.index)] = issue
    #Adding Data Updates
    for issue in bundle_rows_data:
        summary.loc[len(summary.index)] = issue
    return summary

def export_bundle(out_stream = 'BundleList.xlsx', bundle_rows_incl = [], bundle_rows_off = [], bundle_rows_data = [], bundle_rows_org = [], phire_audit_list=[]):
    workbook = xlsxwriter.Workbook(out_stream)

    #Adding bundle report worksheet
    worksheet_br = workbook.add_worksheet('Bundle Report')

    #adjusting the cell width and format
    cell_format_table_header = workbook.add_format({"font_name":'Arial', "font_size":9, "align":"center",
     "valign":"vcenter", "border":1, 'text_wrap':True, 'bold':True})

    cell_format_included = workbook.add_format({"font_name":'Arial', "font_size":9, "align":"center",
     "valign":"vcenter", "border":1, 'text_wrap':True, 'bg_color':'white'})

    cell_format_off = workbook.add_format({"font_name":'Arial', "font_size":9, "align":"center",
     "valign":"vcenter", "border":1, 'text_wrap':True, 'bg_color':'yellow'})

    cell_format_data = workbook.add_format({"font_name":'Arial', "font_size":9, "align":"center",
     "valign":"vcenter", "border":1, 'text_wrap':True, 'bg_color':'8ea9db'})

    cell_format_org = workbook.add_format({"font_name":'Arial', "font_size":9, "align":"center",
     "valign":"vcenter", "border":1, 'text_wrap':True, 'bg_color':'a7f432'})

    #cell_format_title = workbook.add_format({"font_name":'Arial', "font_size":14})

    worksheet_br.set_column(0, 0, 15)
    worksheet_br.set_column(1, 1, 50)
    worksheet_br.set_column(2, 6, 22)
    worksheet_br.set_column(7, 8, 10)
    worksheet_br.set_column(9,9,10)
    worksheet_br.set_column(10,10,15)
    worksheet_br.set_column(11,11,10)
    worksheet_br.set_column(12,16,20)
    worksheet_br.set_column(17,17,30)

    #writing the header row of the bundle list documentation
    for column_index in range(0, len(bundle_header)):
        worksheet_br.write(0, column_index, bundle_header[column_index], cell_format_table_header)
    
    row_index = 1
    #writing the jira included in bundle and formatting the color background of these cells
    for row in bundle_rows_incl:
        worksheet_br.set_row(row_index,25)
        for column_index in range(0, len(row)):
            worksheet_br.write(row_index, column_index, row[column_index], cell_format_included)
        row_index += 1
    #writing the jira in off-bundle category and formatting the color background of these cells
    for row in bundle_rows_off:
        worksheet_br.set_row(row_index,25)
        for column_index in range(0, len(row)):
            worksheet_br.write(row_index, column_index, row[column_index], cell_format_off)
        row_index += 1
    #writing the jira in org dept update category and formatting the color background of these cells
    for row in bundle_rows_org:
        worksheet_br.set_row(row_index,25)
        for column_index in range(0, len(row)):
            worksheet_br.write(row_index, column_index, row[column_index], cell_format_org)
        row_index += 1
    #writing the jira in data update category and fomatting the color background of these cells
    for row in bundle_rows_data:
        worksheet_br.set_row(row_index,25)
        for column_index in range(0, len(row)):
            worksheet_br.write(row_index, column_index, row[column_index], cell_format_data)
        row_index += 1
    
    #Adding phire audit worksheet
    worksheet_pa = workbook.add_worksheet('Phire Audit')

    cell_format_header = workbook.add_format({"font_name":'Arial', "font_size":11, "align":"center",
     "valign":"vcenter", "border":1, 'text_wrap':True, 'bold':True})

    cell_format_audit = workbook.add_format({"font_name":'Arial', "font_size":9, "align":"center",
     "valign":"vcenter", "border":1, 'text_wrap':True, 'bg_color':'b5e2ff'})

    worksheet_pa.set_column(0,1,20)
    worksheet_pa.set_column(2,2,10)
    worksheet_pa.set_column(3,3,20)
    worksheet_pa.set_column(4,4,15)
    worksheet_pa.set_column(5,5,10)
    worksheet_pa.set_column(6,6,30)

    for column_index in range(0, len(audit_header)):
        worksheet_pa.write(0, column_index, audit_header[column_index], cell_format_header)

    row_index = 1
    for row in phire_audit_list:
        for column_index in range(0, len(row)):
            worksheet_pa.write(row_index, column_index, row[column_index], cell_format_audit)
        row_index += 1

    workbook.close()



def download_excel():
    out_stream = BytesIO()
    export_bundle(out_stream, bundle_rows_incl, bundle_rows_off, bundle_rows_data, bundle_rows_org, phire_audit_list)
    out_stream.seek(0)
    return out_stream.getvalue();