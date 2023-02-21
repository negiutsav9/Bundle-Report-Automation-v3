import streamlit as st
import pandas as pd
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from io import BytesIO
from table_classes import bundle_rows, audit_rows
from datetime import datetime
import xlsxwriter
from jira import JIRA
import os
import sys
from dotenv import load_dotenv
from supabase import create_client, Client

sys.path.insert(1,os.getcwd() +"\pages")
load_dotenv()

st.set_page_config(layout="wide")

bundle_rows_incl = []
bundle_rows_off = []
bundle_rows_data = []
bundle_rows_org = []
phire_audit_list = []
fiscal_year_1 = '2022'
fiscal_year_2 = '2023'
test_jira = ''
report_title = ''

@st.cache_resource
def init_connection():
    url = st.secrets["supabase_url"]
    key = st.secrets["supabase_key"]
    return create_client(url, key)

supabase = init_connection()

bundle_dates_data = supabase.table("dates").select("*").execute().data
bundle_dates = pd.DataFrame(bundle_dates_data)["migration_date"].tolist()

jira_obj = JIRA(os.getenv("HOST_URL"),basic_auth=(os.getenv("JIRA_USER_ID"),os.getenv("JIRA_API_KEY")))

bundle_header = ["JIRA #", "Summary", "Team", "Bundle Status", "JIRA Status", "Assignee",
 "Reporter", "Priority", "Sub-Priority", "Category", "Prioritization Date", "HRS/EPM", "HRQA/EPQAS Date",
 "PHIRE CR#", "CR Type", "HRS/EPM Date", "Off Bundle Reason", "Project Association"]

audit_header = ["Phire CR Number", "JIRA Tracking #", "Target DB", "Migrated On", "Migrated By", "CR Type", "Query"]

summary = pd.DataFrame(columns=bundle_header)

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
    bundle_rows_incl.clear()
    bundle_rows_off.clear()
    bundle_rows_data.clear()
    bundle_rows_org.clear()
    phire_audit_list.clear()
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

    #Adding Included Row
    for issue in bundle_rows_incl:
        summary.loc[len(summary.index)] = issue
    #Adding Off-bundle Row
    for issue in bundle_rows_off:
        summary.loc[len(summary.index)] = issue
    #Adding Data Updates
    for issue in bundle_rows_data:
        summary.loc[len(summary.index)] = issue

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
    #Adding Org Dept Updates
    for issue in bundle_rows_org:
        summary.loc[len(summary.index)] = issue
    else:
        print("NO ORGANIZATION DEPARTMENT UPDATE AT THIS TIME")

def showBundleSummary():
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



st.sidebar.title("Create New Bundle Report")
new_form = st.sidebar.form("new_form")
report_name = new_form.text_input("Name")
report_jira = new_form.text_input("JIRA")
start_date = new_form.date_input("Start Date")
end_date = new_form.date_input("End Date")
checkMigration = new_form.checkbox("Include Migration Query Result")
uploadMigration = new_form.file_uploader("Upload Migration Query Results")
checkSQL = new_form.checkbox("Include SQL Query Result")
uploadSQL = new_form.file_uploader("Upload SQL Query Results")
checkOrgUpdates = new_form.checkbox("Include Organization Department Updates")
submitForm = new_form.form_submit_button("Create a Bundle Report")

if submitForm:
    st.title(report_name)
    jira_col, start_date_col, end_date_col = st.columns([1.5,2,2])
    with jira_col:
        st.metric(label="JIRA Bundle", value=report_jira)
    with start_date_col:
        st.metric(label="Start Date", value=start_date.strftime("%d %B %Y"))
    with end_date_col:
        st.metric(label="End Date", value=end_date.strftime("%d %B %Y"))
    if(checkMigration):
        processCSV(pd.read_csv(uploadMigration), "Migration Data")
    if(checkSQL):
        processCSV(pd.read_csv(uploadSQL), "SQL Data")
    if(checkOrgUpdates):
        processOrg(start_date, end_date, "Org Dept Update")
    st.header("Summary")
    bundleSummary = showBundleSummary()
    bundleSummary.index += 1 
    st.dataframe(bundleSummary, use_container_width=True)
    st.header("Statistics")
    stats_Status = bundleSummary['Bundle Status'].value_counts().to_dict()
    stats_Team = bundleSummary["Team"].value_counts().to_dict()
    stats_Type = bundleSummary["CR Type"].value_counts().to_dict()
    stats_Category = bundleSummary["Category"].value_counts().to_dict()
    status_col1, status_col2, status_col3, status_col4 = st.columns(4)
    with status_col1:
        st.metric("Included in Bundle", stats_Status["Included in Bundle"])
    with status_col2:
        st.metric("Off Bundle", stats_Status["Off-bundle"])
    with status_col3:
        if(checkOrgUpdates):
            st.metric("Organization Department Update", stats_Status["Org Dept Update"])
        else:
            st.metric("Organization Department Update", 0)
    with status_col4:
        st.metric("Data Update", stats_Status["Data Update"])

    fig = make_subplots(rows=1, cols=2,  specs=[[{"type": "pie"}, {"type": "pie"}]], subplot_titles=('Projects VS Ops', 'Phire CR-Type'))
    fig.add_trace(
        go.Pie(
            values=list(stats_Category.values()), 
            labels=list(stats_Category.keys()),
            name="Project vs Ops"
        ),row=1, col=1
    )
    fig.add_trace(
        go.Pie(
            values=list(stats_Type.values()), 
            labels=list(stats_Type.keys()),
            name="CR-Type"
        ),row=1, col=2
    )
    st.plotly_chart(fig,  use_container_width=True)

    fig_team = go.Figure()
    fig_team.add_trace(
        go.Bar(
            y = list(stats_Team.keys()),
            x = list(stats_Team.values()),
            orientation="h",
             marker=dict(
                color='rgba(58, 71, 80, 0.6)',
                line=dict(color='rgba(58, 71, 80, 1.0)', width=3)
            )
        )
    )
    fig_team.update_layout(title="Team Distribution")
    st.plotly_chart(fig_team, use_container_width=True)
    button_col_1, button_col_2 = st.columns(2)
    with button_col_1:
        binary_content = download_excel()
        name = report_name + ".xlsx"
        st.download_button("Download the Bundle Report", data=binary_content, file_name=name)
    with button_col_2:
        st.button("Attach the report to JIRA")
else:
    st.title('Bundle Report Generator')
    st.write("This tool is used to create bundle report which summarizes the changes that are migrated to the University of Wisconsin's Human Resoruce System")
    st.write("\n")
    st.write("\n")
    st.write("\n")
    st.write("Use the form in the sidebar to create a new bundle report")