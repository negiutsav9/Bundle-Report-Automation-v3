import streamlit as st
import pandas as pd
import numpy as np
from app_functions import processCSV, processOrg, showBundleSummary, download_excel
from plotly.subplots import make_subplots
import plotly.graph_objects as go

st.set_page_config(layout="wide")

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
    jira_col, start_date_col, end_date_col = st.columns([1,2,2])
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
    st.write("This tool is used to create bundle report which summarizes the changes that are migrated to the University of Wisconsin's Human Resource System")
    st.write("\n")
    st.write("\n")
    st.write("\n")
    st.write("Use the form in the sidebar to create a new bundle report")
