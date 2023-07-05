import streamlit as st
import zipfile
import os
import pandas as pd
from styleframe import StyleFrame
from datetime import datetime, time
from io import BytesIO
import shutil

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

st.write('''
    # iRaste Telangana Reports
    ## Styled Report Maker
    Makes styled reports containing Schedule Departure Time for each bus from the uploaded driver roster and sheet with Scheduled Departure Time.
    ***
''')


uploaded_file = st.file_uploader(
    label="Upload the zip file containing Data corresponding to each depot",
    type=['zip'],
    accept_multiple_files=False,
    label_visibility="visible"
)

month = st.selectbox("Select the month", ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October","November", "December"])

def empty_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")

def compress(file_names):
    print("File Paths:")
    print(file_names)

    path = "./files/"

    # Select the compression mode ZIP_DEFLATED for compression
    # or zipfile.ZIP_STORED to just store the file
    compression = zipfile.ZIP_DEFLATED

    # create the zip file first parameter path/name, second mode
    zf = zipfile.ZipFile("./files/StyledReports.zip", mode="w")
    try:
        for file_name in file_names:
            zf.write(path + file_name, file_name, compress_type=compression)

    except FileNotFoundError:
        print("An error occurred")
    finally:
        # Don't forget to close the file!
        zf.close()

def process(uploaded_file):
    with st.spinner(text="Processing it.."):
        if uploaded_file.type != 'application/x-zip-compressed':
            st.write(uploaded_file.type)
            st.write('Please upload zip file containing the data, not something else.')
        else:
            if os.path.isdir('./files'):
                with zipfile.ZipFile(uploaded_file, 'r') as z:
                    z.extractall('./files')

        parentFolder = [x for x in os.listdir('./files')]

        df1Labels = ['SL. NO', 'CHLN_DATE', 'VEHICLE No.', 'DRIVER', 'SERVICE', 'SER_ TP', 'OPTD_ KMS', 'HSD', 'KMPL']
        df2Labels = ['SRL NO', 'SERV NO', 'ROUTE NO', 'KMS', 'SCHEDULE DEP', 'SCHEDULE ARV' ,'D/', 'TP']
        df2Labels6 = ['SRL NO', 'SERV NO', 'KMS', 'SCHEDULE DEP','D/', 'TP']
        df1 = None
        df2 = None
        # st.write(parentFolder)
        for folder in parentFolder:
            if 'Styled' not in folder:
                subFolders = [x for x in os.listdir(f'./files/{folder}') if 'Styled' not in x]
                for subFolder in subFolders:
                    files = os.listdir(f"./files/{folder}/{subFolder}")
                    # st.write(files)
                    df1 = pd.read_excel(f"./files/{folder}/{subFolder}/{files[0]}")
                    df2 = pd.read_excel(f"./files/{folder}/{subFolder}/{files[1]}")
                    df1.columns = df1Labels
                    df1.insert(5,"START TIME", 0)
                    if len(list(df2)) == 8:
                        df2.columns = df2Labels
                    else:
                        df2.columns = df2Labels6
                    count = 0
                    for row in df1.itertuples():
                        if "SL. NO" in str(row[1]):
                            count = row[0]
                        try:
                            val = list(df2[ df2['SERV NO'] == row[5] ]['SCHEDULE DEP'].index)[0]
                            try:
                                tme = str(df2.loc[val, 'SCHEDULE DEP'].strftime('%H:%M:%S'))
                            except:
                                tme = df2.loc[val, 'SCHEDULE DEP']
                                tme = tme.split(' ')
                                # time = datetime.strptime(time[-1],"%H:%M:%S")
                                tme = str(time.fromisoformat(tme[-1]))
                            df1.loc[row[0],'START TIME'] = datetime.strptime(tme, '%H:%M:%S').time()
                        except Exception as e:
                            continue
                    df1 = df1.iloc[count+1:]
                    # df1.to_excel(f"./{subFolder}-ST-TIME.xlsx")

                    excel_writer = StyleFrame.ExcelWriter(f"./files/{subFolder}-{month}-ST-TIME-Styled.xlsx")
                    df = df1
                    sf = StyleFrame(df)
                    sf.to_excel(
                        excel_writer=excel_writer,
                        best_fit=list(df),
                        columns_and_rows_to_freeze='B2',
                        row_to_add_filters=0,
                        sheet_name=f"Sheet 1"
                    )
                    excel_writer.save()
        files = [x for x in os.listdir('./files') if 'Styled' in x and 'zip' not in x]

        compress(files)

        st.success("Files created successfully, download as ZIP file.")
        clicked = None
        with open("./files/StyledReports.zip", "rb") as f:
            clicked = st.download_button(
                label="Download Styled Reports",
                data = f,
                file_name="StyledReports.zip",
                mime='application/zip'
            )
        if clicked:
            if os.path.exists('./files/StyledReports.zip'):
                os.remove('./files/StyledReports.zip')
            empty_folder('./files')
    
if uploaded_file is not None and month is not None:
    process(
        uploaded_file=uploaded_file
    )
    uploaded_file = None