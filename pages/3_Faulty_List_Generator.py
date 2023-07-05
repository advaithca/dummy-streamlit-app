import streamlit as st
import pandas as pd
from io import BytesIO

st.write('''
    # iRaste Telangana Reports
    ## Faulty List Generator
    Generates Faulty List corresponding to a time frame.
    ***
''')

file = st.file_uploader('Upload the file', type=['xlsx','csv'], accept_multiple_files=False, label_visibility='visible')
file2 = st.file_uploader('Upload previous faulty list', type=['xlsx','csv'], accept_multiple_files=False, label_visibility='visible')

if file is not None and file2 is not None:
    with st.spinner(f'Loading DataFrame from {file.name}'):
        if "xlsx" in file.name:
            df = pd.read_excel(file)
        elif "csv" in file.name:
            df = pd.read_csv(file)
    st.success("Loaded DataFrame successfully")
    df['6_Imp_alerts_sum']  = 0
    print("\n")

    with st.spinner("Summing up Important Alerts"):
        for row in df.iterrows():
            val = row[1]['me_fcw_count']+row[1]['me_hmw_count']+row[1]['me_lldw_count']+row[1]['me_pcw_count']+row[1]['me_pdz_count']+row[1]['me_rldw_count']
            df.loc[row[0],'6_Imp_alerts_sum'] = val
    st.success("Summed up the 6 important alerts")
    # Grouping by vehicle
    groupedByVehicleNum = df.groupby(by="VehicleNum")

    dataForNewDf = []

    with st.spinner("Calculating sum of distance and sum of the 6 important alerts per vehicleNum per date"):
        for row in groupedByVehicleNum:
            vehicleNum = row[0]
            dfVehicle = row[1]
            groupedByDate = dfVehicle.groupby(by='startDate')
            for row2 in groupedByDate:
                date = row2[0]
                dfDate = row2[1]
                dataForNewDf.append([vehicleNum, date,sum(dfDate['distance']),sum(dfDate['6_Imp_alerts_sum']),list(dfDate['groupName'].unique())[0]])
        
        # Creating new dataframe with required details
        newDf = pd.DataFrame(data=dataForNewDf, columns=["VehicleNum", "startDate","distance","6_Imp_alerts_sum","depot_name"])
        # Selecting only those which have zero imp alert sum or near zero distance
        trues = newDf.loc[(newDf['6_Imp_alerts_sum']==0) | (newDf['distance']<1)]

        # Converting datetime objects to strings
        try:
            trues['dateTime']=trues['startDate'].dt.strftime("%Y-%m-%d")
        except:
            trues["dateTime"] = trues['startDate']
        # Grouping the zero imp alert data by vehicles
        gbVeh = trues.groupby(by="VehicleNum")

        uniqueDates = list(trues['dateTime'].unique())
        uniqueDates.sort()
        nDays = len(uniqueDates)
        timeString = f"{uniqueDates[0]} -- {uniqueDates[-1]}"
        nUniqueBusses = len(trues['VehicleNum'].unique())

    nUniqueBusses = len(newDf['VehicleNum'].unique())
    dfStat = pd.DataFrame([timeString, nDays, nUniqueBusses], columns=['Basic Stats'], index=["Time window", "Number of days", "Number of busses"])
    busDf = pd.read_excel('pages/Master Copy Revised.xls')
    nBusses = len(busDf['Veh No'].unique())
    difference = nBusses - nUniqueBusses
    dfStat = dfStat.append(pd.DataFrame([nBusses, difference], columns=['Basic Stats'], index=["Number of busses in Master Copy", "Difference"]))

    st.success("Intermediate Calculations Done!")
    # Final calculations 
    from datetime import datetime, timedelta
    newData = []
    with st.spinner("Calculating sum of distance and sum of the 6 important alerts per vehicleNum across multiple dates"):
        for row in gbVeh:
            vNo = row[0]
            vDf = row[1]
            vDf.sort_values(by=['dateTime'])
            prev = ''
            prev2 = ''
            count2 = 0
            count = 0
            maxC = 0
            maxC2 = 0
            for row2 in vDf.itertuples():
                if row2[4] == 0:
                    if prev == '':
                        count += 1
                        prev = datetime.strptime(row2[-1], "%Y-%m-%d")
                        fromDate = prev
                    elif datetime.strptime(row2[-1], "%Y-%m-%d") - prev == timedelta(days=1):
                        count += 1
                        prev = datetime.strptime(row2[-1], "%Y-%m-%d")
                    else:
                        if count > maxC:
                            maxC = count
                if int(row2[3]) == 0:
                    if prev2 == '':
                        count2 += 1
                        prev2 = datetime.strptime(row2[-1], "%Y-%m-%d")
                        fromDate2 = prev2
                    elif datetime.strptime(row2[-1], "%Y-%m-%d") - prev2 == timedelta(days=1):
                        count2 += 1
                        prev2 = datetime.strptime(row2[-1], "%Y-%m-%d")
                    else:
                        if count2 > maxC2:
                            maxC2 = count2
                toDate = prev
                toDate2 = prev2
            try:
                newData.append([vNo, sum(vDf['distance']),sum(vDf['6_Imp_alerts_sum']),count,f"{fromDate.strftime('%Y-%m-%d')} -- {toDate.strftime('%Y-%m-%d')}", count2, f"{fromDate2} -- {toDate2}",list(vDf['depot_name'].unique())[0]])
            except AttributeError:
                newData.append([vNo, sum(vDf['distance']),sum(vDf['6_Imp_alerts_sum']),count,f"{fromDate} -- {toDate}", count2, f"{fromDate2} -- {toDate2}",list(vDf['depot_name'].unique())[0]])
            except NameError:
                newData.append([vNo, sum(vDf['distance']),sum(vDf['6_Imp_alerts_sum']),count,f"{fromDate} -- {toDate}", count2, f"NA",list(vDf['depot_name'].unique())[0]])
        newNewDf = pd.DataFrame(data=newData, columns = ['VehicleNum', 'distance', '6Imp_alerts_sum',  'numOfDays ZeroCASAlerts continued', "ZeroCAS fromDate-ToDate", "numOfDays ZeroDistance or near ZeroDistance", "ZeroDist from-to","Depot name"])
        master_copy_df = pd.read_excel('pages/Master Copy Revised.xls')
        
        # Take master copy, remove all busses in raw data from it.
        not_found_entries = master_copy_df[~master_copy_df['Veh No'].isin(newDf['VehicleNum'].unique())]
        unfound = newDf[~newDf['VehicleNum'].isin(master_copy_df['Veh No'])]['VehicleNum'].unique()
        len_not_found_entries = len(not_found_entries)
        len_newNewDf = len(newNewDf)
        numColumnsToAppend = len_newNewDf - len_not_found_entries

        dfStat = dfStat.append(pd.DataFrame([len_not_found_entries], columns=['Basic Stats'], index=["Number of busses not found in Raw Data"]))
        empty_df = pd.DataFrame(index=not_found_entries.index)
        for i in range(numColumnsToAppend):
            not_found_entries[f'Column {i}'] = ''
        not_found_entries.columns = ["Depot name","VehicleNum"]
        empty_row = pd.DataFrame(index=[0,1], columns=newNewDf.columns)
        empty_row['VehicleNum'][1] = "List of Busses not found in Raw Data"
        newNewDf = pd.concat([newNewDf, empty_row, not_found_entries], ignore_index=True, axis=0)

        newNewDf['Owner'] = ''
        newNewDf['Previous Week Status'] = ''
        newNewDf['Previous Week Comments'] = ''

        # Create ExcelFile from file2
        if "xlsx" in file2.name:
            prevFaultySheet = pd.ExcelFile(file2)
        elif "csv" in file2.name:
            st.error("Please upload an excel file for previous faulty list")
            file2 = None
        
        # Get the sheet names
        sheetNames = prevFaultySheet.sheet_names

        # Read the last sheet
        prevFaultyList = pd.read_excel(file2, sheet_name=sheetNames[-1])

        # st.write(prevFaultyList)
        for row in newNewDf.iterrows():
            if row[1]['VehicleNum'] in prevFaultyList['BUS NUMBER'].unique():
                newNewDf.at[row[0], 'Previous Week Status'] = prevFaultyList[prevFaultyList['BUS NUMBER']==row[1]['VehicleNum']]['FIXED'].values[0]
                newNewDf.at[row[0], 'Previous Week Comments'] = prevFaultyList[prevFaultyList['BUS NUMBER']==row[1]['VehicleNum']]['COMMENTS'].values[0]


        # DMS thing : Eliminate from entire faulty list. 3rd Pass
        tenDMS = ['TS09Z7676','TS09Z7674','TS08Z0204','TS08Z0244','TS15Z0171','TS09Z7620','TS08Z0251','TS09Z7652','TS09Z7625','TS09Z7677']
        for row in newNewDf.iterrows():
            if row[1]['VehicleNum'] in tenDMS:
                newNewDf.drop(row[0], inplace=True)
        
        Buffer = BytesIO()
        with pd.ExcelWriter(Buffer, engine='xlsxwriter') as writer:
            newNewDf.to_excel(writer, sheet_name="Faulty List")
            dfStat.to_excel(writer, sheet_name="log")

            writer.save()

    st.success("All Calculations Done!")
    st.dataframe(newNewDf)
    download2 = st.download_button(
        label="Download Result",
        data=Buffer,
        file_name=f'{file.name[:-5]}_result.xlsx',
        mime='application/vnd.ms-excel'
    )

    st.stop()