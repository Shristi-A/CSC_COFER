import datetime
import os
import shutil
import openpyxl
import Refine as rf
import pandas as pd
import send_emails_smtp as se
import extn_utils as eu
import pandas_etl as pe
import utils
import utils as ut

allVendorDFs = pd.DataFrame()
today = datetime.date.today()
today = datetime.datetime.strptime("2025-1-6", '%Y-%m-%d')
formatted_date = f"{today.year}-{today.month}-{today.day}"
print(formatted_date)

if today.weekday() == 0:
    lastWeekDate = today-datetime.timedelta(days=7)
    lastWeekDate=lastWeekDate.strftime('%m-%d-%Y')
    print(lastWeekDate)
else:
    #test
    lastWeekDate = '12-2-2024'
    masterConferReportFileName = "test_Master COFER Report-2024-12-09.xlsx"
    print("Today is not monday")

def GetExcelFileToDF(dfName,sheetName, fileName):
    os.environ['fileName'] = fileName
    os.environ['sheetName'] = sheetName
    config = ut.load_json("resources/extn/excelToDF.json")
    if config is not None:
        etl = pe.PandasEtl(config)
        df_Name = etl.from_source()
    return df_Name

def downloadTheFiles(MasterConferReportFileName):
    found_gdf = config['googleSheetConfig']['masterConferFileFromCognosSourceFolder']
    fileMimeType = config['googleSheetConfig']['fileMimeType']
    cred = config['googleSheetConfig']['cred']
    files = eu.getFilesInGoogleDriveFolder(found_gdf, fileMimeType, cred)
    newFile = {'name': MasterConferReportFileName}
    # check if the next file is in the google drive. If yes, download it.
    for file in files:
        if newFile.get('name') in file.values():
            eu.downloadFileFromGoogleDrive(file, destinationFolder, cred)
            print(file['name']);
            print('File added in the destination folder.')
        else:
            print('File not found')
    print("file downloaded")

def createFileAndTab(attachment,filtered_df,sheet_name):
    os.environ["fileName"] = attachment
    os.environ["operation"] = sheet_name
    downloadFileConfig = ut.load_json("resources/extn/dfToExcel.json")
    if downloadFileConfig is not None:
        etl = pe.PandasEtl(downloadFileConfig)
        etl.to_destination(filtered_df)
    else:
       print("dfToExcel json file did not load properly.")


def sendemail(notification,filename):
    if notification['process'] is True:
       path= notification.get('emailbody')
       finalBody = eu.getEmailBodyFromHTMLFile(path)
       subject = notification.get('subject')+formatted_date
       finalsubject = f"{subject} {today}"
       emailAddress = notification.get('to')
       allCCEmailAddress = notification.get('cc')
       fromEmail = notification.get('from_replyTo')
       allBCCEmailAddress = ''
       try:
          #extn.setColumnWidthDynamically(attachment)
         if filename != '':
            email_params_list = [se.EmailParams(fromEmail, emailAddress, allCCEmailAddress, allBCCEmailAddress, fromEmail, subject, finalBody, [masterCoferAttachmentPath], filename)]
         else:
             email_params_list = [se.EmailParams(fromEmail, emailAddress, allCCEmailAddress, allBCCEmailAddress, fromEmail, subject,finalBody, [], filename)]
         se.send_email_with_starttls(email_params_list)
       except Exception as e:
           eu.print_colored("An error occurred while sending the email:" + str(e), "red")
    else:
        print("notification is turned off." )

def createCSCMasterCoferAttachment(filename):
    outputfolder = './/output//Master COFER Template.xlsx'
    sourcefile = './/MasterCOFERFromCognos//Master COFER Template.xlsx'
    shutil.copy(sourcefile,outputfolder)
    os.rename(sourcefile,f'.//output//{filename}')

if __name__ == '__main__':
    #download the most current Cofer File
   os.environ['today'] = today.strftime('%Y-%m-%d')
   config = utils.load_json("./resources/extn/Vendors.json")
   if config is not None:
       print(config)
       destinationFolder = config['googleSheetConfig']['destinationFolder']
       MasterConferReportFileName = f"Master COFER Report-{formatted_date}.xlsx"
       #MasterConferReportFileName = "test_Master COFER Report-2024-12-09.xlsx"
       #Download the initial COFER file from the google drive
       #downloadTheFiles(MasterConferReportFileName)
       #load the master Cofer file into df.
       filePath = f'{destinationFolder}/{MasterConferReportFileName}'
       mainDF = GetExcelFileToDF('MasterCOFER', 'page',filePath )
       #print(mainDF)
       vendorFilePath = config['vendorConfig']['VendorFilePath']
   for vendors in config['AllVendors']:
       if vendors['process'] == True:
          #for each vendors- go to the specified folder, match the vendorname, and lookup the item no in the previous file and get the comments
          vendorName = vendors['googleFolder']
          sharedGoogleFolders = vendorFilePath+vendorName
          previousVendorCoferFile = config['vendorConfig']['VendorFilePrefix']+' '+vendors['vendorName']+' '+lastWeekDate+'.xlsx'
          fileName = sharedGoogleFolders+'/'+previousVendorCoferFile
          try:
             vendorDf= GetExcelFileToDF(vendorName, 'Sheet1', fileName)
             vendorDf_cleaned = vendorDf.dropna(axis=1,how='all')
             allVendorDFs = pd.concat([allVendorDFs,vendorDf],  ignore_index=True)
             print(f"{previousVendorCoferFile} was found.")
          except Exception as e:
             eu.print_colored(f"{previousVendorCoferFile} not found:" + str(e), "red")
       else:
           print("vendor process is false.")
   NotifyCSC= config["NotifyCSC"]
   mainDF = mainDF.merge(allVendorDFs[['PO + Part Number','Comments', 'GSA Comments','Response']], on='PO + Part Number', how='left')
   finalmainDF= rf.refineDF(mainDF)
   masterCoferFileName = f"Master Cofer - {formatted_date}.xlsx"
   masterCoferAttachmentPath=f"./Output/{masterCoferFileName}"
   createCSCMasterCoferAttachment(masterCoferFileName)
   createFileAndTab(masterCoferAttachmentPath, finalmainDF, 'Master COFER')
   rf.formatHeader(masterCoferAttachmentPath, 'Master COFER')

   # with pd.ExcelWriter(masterCoferAttachmentPath,mode='a',engine='openpyxl', if_sheet_exists ='overlay') as writer:
   #     finalmainDF.to_excel(writer,sheet_name = 'Master COFER', header=True, index=False)

   sendemail(NotifyCSC,masterCoferFileName);
   for shareToVendors in config['AllVendors']:
        if shareToVendors['process'] == True:
          vendorName = shareToVendors['vendorName']
          vendorGoogleFolder = shareToVendors['googleFolder']
          sharedGoogleFolderfile = vendorFilePath+vendorGoogleFolder
          if vendorName in mainDF['Vendor Name'].values:
              vendorFileName = config['vendorConfig']['VendorFilePrefix'] + ' ' + vendorName + ' ' + today.strftime('%Y-%m-%d') + '.xlsx'
              #vendorFileName = config['vendorConfig']['VendorFilePrefix'] + ' ' + vendorName+ ' ' + '2024-12-09' + '.xlsx'
              attachment = sharedGoogleFolderfile+"//"+vendorFileName
              filteredVendorItemsdf = mainDF[mainDF['Vendor Name'] == vendorName]
             # createFileAndTab(attachment,filteredVendorItemsdf,"Sheet1")
              print(vendorName)
   NotifyVendor = config["NotifyVendor"]
   #sendemail(NotifyVendor, "");