import streamlit as st
import pandas as pd
import zipfile

st.set_page_config(
    page_title="iRaste Telangana Reports",
    page_icon="Ihub-Data Logo- Vector.jpg",
    layout="wide",
    initial_sidebar_state="auto",
)

st.write(
    """
# iRaste Telangana Reports

### This app helps with creation of various reports for iRaste Telangana Project.  

- To create reports with 'Scheduled Time' added to the driver roster, choose `Styled Report Maker` option from the sidebar and upload the Monthly file which has one folder 
for each Depot.
- To create Top 5 bottom 5 report, choose the `Report Maker` option, and upload the Styled Reports (If it isn't present on the server)
and the Raw data corresponding to the month (Excel file usually with the name `Report_<startDate>_<endDate>.xlsx`)
- To create faulty list, choose the option `Faulty List Validator` and upload the Data corresponding to time frame
for which the faulty list is to be generated.
Haha
### The generated reports can all be downloaded.
***
"""
)
