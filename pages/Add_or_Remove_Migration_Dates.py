import streamlit as st
import pandas as pd
import os
import numpy as np


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
    newdata = [[new_date.strftime('%m/%d/%Y'),new_jira,new_period]]
    new_df = pd.DataFrame(newdata, columns=["Migration Date","JIRA","Pay Period"])
    new_df.to_csv(os.getcwd() + '/pages/dates.csv', mode='a', index=False, header=False)
    st.experimental_rerun()

if(os.stat(file_path).st_size == 0):
    st.write("\n")
    st.write("\n")
    st.write("No migration dates in the bundle report generator")
    st.write("\n")
    st.write("To enter a date, use the form in the sidebar")
else:
    dates = pd.read_csv(os.getcwd() + '/pages/dates.csv', header=None) 

    st.sidebar.header("Remove a Date")
    remove_form = st.sidebar.form("remove_form")
    remove_date = remove_form.date_input("Migration Date to be removed")
    submitRemove = remove_form.form_submit_button("Remove the migration date")
    if(submitRemove):
        remove_date = remove_date.strftime('%m/%d/%Y')
        dates = dates.loc[dates[0] != remove_date]
        dates.to_csv(os.getcwd() + '/pages/dates.csv', mode='w', index=False, header=False)
        st.experimental_rerun()

    dates.columns = ["Migration Date","JIRA","Pay Period"]
    st.dataframe(dates, use_container_width=True)