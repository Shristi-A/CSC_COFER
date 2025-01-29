import os
import pandas as pd
import extn_utils as eut
from commonUtils import common as c
from googleapiclient.errors import HttpError
from datetime import date, timedelta
import utils as ut
import pandas_etl as pe

startdate = '10/01/2023'
enddate = '09/30/2024'
#CoferFileList = ['Master COFER - 10-23-23.xlsx', 'Master COFER - 9-23-24.xlsx', 'Master COFER - 9-16-24.xlsx', 'Master COFER - 9-9-24.xlsx', 'Master COFER - 9-3-24.xlsx', 'Master COFER - 8-26-24.xlsx', 'Master COFER - 8-19-24.xlsx', 'Master COFER - 8-12-24.xlsx', 'Master COFER - 8-5-24.xlsx', 'Master COFER - 7-29-24.xlsx', 'Master COFER - 7-22-24.xlsx', 'Master COFER - 7-15-24.xlsx', 'Master COFER - 7-8-24.xlsx', 'Master COFER - 7-1-24.xlsx', 'Master COFER - 6-24-24.xlsx', 'Master COFER - 6-17-24.xlsx', 'Master COFER - 6-10-24.xlsx', 'Master COFER - 6-3-24.xlsx', 'Master COFER - 5-28-24.xlsx', 'Master COFER - 5-20-24.xlsx', 'Master COFER - 5-13-24.xlsx', 'Master COFER - 5-6-24.xlsx', 'Master COFER - 4-29-24.xlsx', 'Master COFER - 4-22-24.xlsx', 'Master COFER - 4-15-24.xlsx', 'Master COFER - 4-8-24.xlsx', 'Master COFER - 4-1-24.xlsx', 'Master COFER - 3-25-24.xlsx', 'Master COFER - 3-18-24.xlsx', 'Master COFER - 3-11-24.xlsx', 'Master COFER - 3-4-24.xlsx', 'Master COFER - 2-26-24.xlsx', 'Master COFER - 2-20-24.xlsx', 'Master COFER - 2-12-24.xlsx', 'Master COFER - 2-5-24.xlsx', 'Master COFER - 1-29-24.xlsx', 'Master COFER - 1-22-24.xlsx', 'Master COFER - 1-16-24.xlsx', 'Master COFER - 1-8-24.xlsx', 'Master COFER - 1-2-24.xlsx', 'Master COFER - 12-26-23.xlsx', 'Master COFER - 12-18-23.xlsx', 'Master COFER - 12-11-23.xlsx', 'Master COFER - 12-04-23.xlsx', 'Master COFER - 11-27-23.xlsx', 'Master COFER - 11-20-23.xlsx', 'Master COFER - 11-13-23.xlsx', 'Master COFER - 11-6-23.xlsx', 'Master COFER - 10-30-23.xlsx', 'Master COFER - 10-16-23.xlsx', 'Master COFER - 10-10-23.xlsx', 'Master COFER - 10-2-23.xlsx']
CoferFileList = ['Master COFER - 1-8-24.xlsx', 'Master COFER - 1-2-24.xlsx','Master COFER - 2-5-24.xlsx']
columnsToCopy = ['Requisition Number', 'Sales Order Date']

def downloadTheFiles():
    fileMimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    cred = 'auth/gsa.gov_api-project-673153394971-d38e0ee051e1.json'
    files = eut.getFilesInGoogleDriveFolder(found_gdf, fileMimeType, cred)
    # check if the next file is in the google drive. If yes, download it.
    for file in files:
        if file.get('name') in CoferFileList:
            eut.downloadFileFromGoogleDrive(file, destinationFolder, cred)
            print(file['name']);
            print('File added in the destination folder.')
        else:
            print('File not found')
    print('Getting the sheet name.')

def combinedAllExcelFiles():
    combined_Data= pd.DataFrame()
    for file in CoferFileList :
        fileDate =file.replace("Master COFER - ","").replace(".xlsx","").replace("-",".")
        sheetname = "Master COFER "+ fileDate
        print(f'Running {sheetname}')
        df = pd.read_excel(file,sheet_name=sheetname,usecols=columnsToCopy);
        df['Fiscal Year'] = 'FY2024'
        df['Last Day on COFER'] = pd.to_datetime(fileDate,format='%m.%d.%y');

        # Ensure 'Sales Order Date' is in datetime format as well
        df['Sales Order Date'] = pd.to_datetime(df['Sales Order Date'], format='%d/%m/%Y')
        # Calculate the difference in days between the two date columns
        df['Days on COFER'] = (df['Last Day on COFER'] - df['Sales Order Date']).dt.days
        combined_Data = pd.concat([combined_Data, df], ignore_index=True)
        print(f'Completed {sheetname}')
    print(combined_Data)

    df_sorted = combined_Data.sort_values(by=['Requisition Number', 'Sales Order Date'], ascending=[True, False])

    # Step 3: Drop duplicates based on 'Requisition Number', keeping the latest 'Sales Order Date'
    df_no_duplicates = df_sorted.drop_duplicates(subset=['Requisition Number'], keep='first')

    # Final DataFrame
    print(df_no_duplicates)
    df_no_duplicates.to_excel("FY24ConferReport.xlsx",index=False);



if __name__ == '__main__':
    destinationFolder = "./"
    found_gdf = '1OMAL8Vfe_WfZF1R2VzlFhNA7gnqMAjHc'
    #downloadTheFiles()
    combinedAllExcelFiles();