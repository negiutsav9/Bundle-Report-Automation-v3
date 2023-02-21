import streamlit as st
import pandas as pd
import os
import numpy as np
from supabase import create_client, Client

@st.cache_resource
def init_connection():
    url = st.secrets["supabase_url"]
    key = st.secrets["supabase_key"]
    return create_client(url, key)

supabase:Client = init_connection()


def addDates(migration_date, jira_issue, pay_period):
    values = {"migration_date":migration_date, "jira_issue":jira_issue, "pay_period":pay_period}
    data = supabase.table("dates").insert(values).execute()

def removeDate(remove_date):
    data = supabase.table("dates").delete().eq("migration_date",remove_date).execute()

dates = supabase.table("dates").select("*").execute()

file_path = os.getcwd() + '/pages/dates.csv'

st.title("Add or Remove Migration Dates")
st.write("\n")
st.write("\n")
st.header("Current Migration Dates")

st.sidebar.header("Add a Date")
add_form = st.sidebar.form("add_form")
new_date = add_form.date_input("New Migration Date")
new_jira = add_form.text_input("Associated JIRA Issue")
new_period = add_form.text_input("Pay Period ID")
submitAdd = add_form.form_submit_button("Add the migration date")
if(submitAdd):
    addDates(new_date.strftime('%m/%d/%Y'), new_jira, new_period)
    st.experimental_rerun()

if(len(dates.data) == 0):
    st.write("\n")
    st.write("\n")
    st.write("No migration dates in the bundle report generator")
    st.write("\n")
    st.write("To enter a date, use the form in the sidebar")
else:
    st.sidebar.header("Remove a Date")
    remove_form = st.sidebar.form("remove_form")
    remove_date = remove_form.date_input("Migration Date to be removed")
    submitRemove = remove_form.form_submit_button("Remove the migration date")
    if(submitRemove):
        removeDate(remove_date.strftime('%m/%d/%Y'))
        st.experimental_rerun()
    df = pd.DataFrame(data=dates.data)
    df = df.rename(columns={"migration_date":"Migration Date","jira_issue":"JIRA Issue", "pay_period":"Pay Period"})
    st.dataframe(df, use_container_width=True)
    