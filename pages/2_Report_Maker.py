import streamlit as st
import pandas as pd
import zipfile
import io
import shutil
from scripts import Driver_Mapper, firstMonth, Report_maker
import os

st.write('''
    # iRaste Telangana Reports
    ## Report Maker
    Makes top 5 - bottom 5 report based on the Styled Reports and raw data that contains all data corresponding to all depots per month.
    ***
''')

def unzip_and_read_excel(zip_data, report_file):
    '''
    Unzips the zip file and reads the excel files in it.
    Parameters:
        zip_data: The zip file data
    '''
    NameData = 'pages/MASTER COPY DRIVER DATA.xlsx'
    
    with zipfile.ZipFile(zip_data, 'r') as zip_ref:
        zip_ref.extractall('temp')

        excel_files = [file_name for file_name in zip_ref.namelist() if file_name.endswith('.xlsx') or file_name.endswith('.xls')]

        month = None    
        month = st.selectbox("Select the month", ['Jan', 'Feb', 'March', 'April', 'May'])
        if month == 'Jan':
            lastMonth = None
        elif month == 'Feb':
            lastMonth = 'January'
        elif month == 'March':
            lastMonth = 'February'
        elif month == 'April':
            lastMonth = 'March'
        elif month == 'May':
            lastMonth = 'April'
        try:
            report_src = pd.read_excel(report_file)
        except:
            report_src = pd.read_csv(report_file)

        root = f"./{month}"
        if month is not None:
            if lastMonth is not None:
                lastMonthZip = st.file_uploader(f"Upload the zip file for {lastMonth}", type="zip")
            
            for file_name in excel_files:
                with zip_ref.open(file_name) as file:
                    df = pd.read_excel(file)
                    # st.dataframe(df)
                    if "BHEL" in file_name:
                        depot_name="BHEL"
                        # continue
                    elif "MIYAPUR" in file_name:
                        depot_name="MIYAPUR"
                        # continue
                    elif "HYD" in file_name:
                        depot_name='HYDERABAD-1'
                        # continue
                    elif "NIZAMABAD - 1" in file_name:
                        depot_name="NIZAMABAD-1"
                        # continue
                    elif "NIZAMABAD - 2" in file_name:
                        depot_name="NIZAMABAD-2"
                        # continue
                    save_name=f"{root}/{depot_name}{root[2:]}.xlsx"

                    # Wont do anything if the file already exists, else create the folder
                    if not os.path.exists(save_name):
                        os.makedirs(root, exist_ok=True)
                        st.write(f"Creating report for {depot_name}...")

                    st.subheader(f"Result for {file_name}:")
                    # st.dataframe(df)
                    with st.spinner(f"Creating report for {depot_name}..."):
                        Driver_Mapper.main(depot_src=df, report_src=report_src, depot_name=depot_name, save_name=save_name)
                    st.success(f"Report for {depot_name} created successfully!")
            
            # This part is supposed to work out of the box, but I will have to change it probably
            if lastMonth is not None and lastMonthZip is not None:
                with zipfile.ZipFile(lastMonthZip, 'r') as zip_ref:
                    zip_ref.extractall('PreviousStyledReports')
                st.write(f"Doing first of month analysis for {month}...")
                with st.spinner(f"Doing first of month analysis for {month}..."):
                    firstMonth.main(folderName=root, lastMonth=lastMonth)
                st.success(f"First of month analysis for {month} done successfully!")

            # Run the report maker on the root folder for each file in folder
            for file_name in os.listdir(root):
                if file_name.endswith('.xlsx'):
                    with st.spinner(f"Creating report for {file_name}..."):
                        Report_maker.main(save_name=f'{root}/{file_name}', report_src=f'{root}/{file_name}', nameDf=NameData)
                    st.success(f"Report for {file_name} created successfully!")
            st.balloons()   
            # Zip contents of the root folder
            shutil.make_archive(f'{month}', 'zip', root)

            # Download the zip file
            with open(f'{month}.zip', 'rb') as f:
                bytes = f.read()
                clicked = st.download_button(label=f"Download {month}.zip", data=bytes, file_name=f'{month}.zip', mime='application/zip')

            # Delete the temporary folder if the download button is clicked
            if clicked:
                shutil.rmtree(root)
                shutil.rmtree('temp')
                os.remove(f'{month}.zip')

def main():
    st.title("ZIP File Upload")

    uploaded_file = st.file_uploader("Upload a StyledReports zip file", type="zip")

    report_file = st.file_uploader("Upload a raw data file", type=['xlsx', 'xls', 'csv'])
    # report = pd.ExcelFile(report_file)
    if uploaded_file and report_file is not None:
        zip_data = io.BytesIO(uploaded_file.read())

        unzip_and_read_excel(zip_data, report_file)

if __name__ == '__main__':
    main()