from datetime import datetime
import os
import shutil
import glob
import Refine as rf
import pandas as pd
import extn_utils
import send_emails_smtp as se
import extn_utils as eu
import pandas_etl as pe
import utils
import utils as ut

'''To hold all the prior Vendor information'''


'''Format the date in mm-dd-yyyy format to download COGNOS generated report, and the rest of the files in m-d-yy format. '''
today = datetime.now().date()
#todaydate='2025-02-18'
#today = datetime.datetime.strptime(todaydate,'%Y-%m-%d').date()
print(f'today: {today}')
formattedTDate = f"{today.month}-{today.day}-{today.year % 100}"
print(f'formattedTDate : {formattedTDate}')

#if today.weekday() == 0:
all_file = glob.glob(f'./Output/Master Cofer - *.xlsx')
lastWeekFile = max(all_file)
formattedLWDate = lastWeekFile.split(" ")[-1].rsplit(".")[0]
#formattedLWDate = f"{lastWeekDate.month}-{lastWeekDate.day}-{lastWeekDate.year % 100}"

print(formattedLWDate)
outputfolder = './/output//'
# else:
# print("Today is not monday")

'''Copy Master Cofer template file'''
def copyMasterCOFERFileTemplate():

    extn_utils.deleteFolderContents(outputfolder)
    sourcefile = './/MasterCOFERFromCognos//Master COFER Template.xlsx'
    shutil.copy(sourcefile, outputfolder)

'''Transform excel file to dataframe'''


def GetExcelFileToDF(sheetName, fileName):
    os.environ['fileName'] = fileName
    os.environ['sheetName'] = sheetName
    config = ut.load_json("resources/extn/excelToDF.json")
    if config is not None:
        etl = pe.PandasEtl(config)
        dfVendor = etl.from_source()
    return dfVendor


'''Download COGNOS generated COFER file from the google folder'''


def downloadTheFiles(cognosCOFERFileName):
    found_gdf = config['googleSheetConfig']['cognosCOFERSourceFolder']
    fileMimeType = config['googleSheetConfig']['fileMimeType']
    cred = config['googleSheetConfig']['cred']
    files = eu.getFilesInGoogleDriveFolder(found_gdf, fileMimeType, cred)
    newFile = {'name': cognosCOFERFileName}
    # check if the next file is in the google drive. If yes, download it.
    for file in files:
        if newFile.get('name') in file.values():
            eu.downloadFileFromGoogleDrive(file, destinationFolder, cred)
            print(file['name']);
            print('File added in the destination folder.')
        else:
            print('File not found')
    print("file downloaded")


'''Create excel file from Dataframe'''


def createFileAndTab(attachment, filtered_df, sheetName):
    os.environ["fileName"] = attachment
    os.environ["tab"] = sheetName
    downloadFileConfig = ut.load_json("resources/extn/dfToExcel.json")
    if downloadFileConfig is not None:
        etl = pe.PandasEtl(downloadFileConfig)
        etl.to_destination(filtered_df)
    else:
        print("dfToExcel json file did not load properly.")

'''Upload current Master COFER File'''


def uploadMasterCoferFile(sourceFilePath):
    path = config['UploadMasterCOFERFile']['roGoogleDrivePath']
    outputfolder = path.replace('\x00', '')
    shutil.copy(sourceFilePath, outputfolder)


'''Send email with or without attachment.'''


def sendemail(notification, masterCoferAttachmentPath,filename):
    if notification['process'] is True:
        path = notification.get('emailbody')
        finalBody = eu.getEmailBodyFromHTMLFile(path)
        subject = notification.get('subject') + " " + formattedTDate
        finalsubject = f"{subject} {today}"
        emailAddress = notification.get('to')
        allCCEmailAddress = notification.get('cc')
        fromEmail = notification.get('from_replyTo')
        allBCCEmailAddress = notification.get('bcc')
        try:
            # extn.setColumnWidthDynamically(attachment)
            if filename != '':
                email_params_list = [
                    se.EmailParams(fromEmail, emailAddress, allCCEmailAddress, allBCCEmailAddress, fromEmail, subject,
                                   finalBody, [masterCoferAttachmentPath], filename)]
            else:
                email_params_list = [
                    se.EmailParams(fromEmail, emailAddress, allCCEmailAddress, allBCCEmailAddress, fromEmail, subject,
                                   finalBody, [], filename)]
            se.send_email_with_starttls(email_params_list)
        except Exception as e:
            eu.print_colored("An error occurred while sending the email:" + str(e), "red")
    else:
        print("notification is turned off.")


def archivePriorMonthFile(folderPath, archiveFolder):
   tMonth = today.month
   if tMonth == 1:
        priorMonth = 12
   else:
        priorMonth = tMonth - 1

   # Get all files in the folder
   for f in os.listdir(folderPath):
        filename = f
        if filename.startswith("GSA"):
            fileDateMonth = int(f.split(" ")[-1].rsplit(".")[0].rsplit("-")[0])

            if type(fileDateMonth) is int and fileDateMonth == priorMonth:
                print(f"Moved {filename} to {archiveFolder}")
                shutil.move(f'{folderPath}//{filename}', f'{folderPath}//{archiveFolder}//{filename}')
            else:
                print(f"{filename} is not from prior month.")


def createMasterCOFERFiles(mainDF):
    allVendorDFs = pd.DataFrame()
    for vendors in config['AllVendors']:
        if vendors['process'] == True:
            # for each vendors- go to their shared drive, match the vendorname, and lookup the item no in the previous file and get the comments
            vendorName = f"COFER-{vendors['vendorName']}"
            sharedVGoogleFolders = vendorSharedDirectory + vendorName
            previousVendorCOFERFile = config['vendorConfig']['VendorFilePrefix'] + ' ' + vendors[
                'vendorFileName'] + ' ' + formattedLWDate + '.xlsx'
            vendorFilePath = sharedVGoogleFolders + '//' + previousVendorCOFERFile
            try:
                print(f"previousVendorCOFERFile:{vendorFilePath}")
                vendorDf = GetExcelFileToDF('Sheet1', vendorFilePath)
                allVendorDFs = pd.concat([allVendorDFs, vendorDf], ignore_index=True)
                print(f"{previousVendorCOFERFile} was found.")
            except Exception as e:
                eu.print_colored(f"{previousVendorCOFERFile} not found:" + str(e), "red")
        else:
            print("vendor process is false.")

        # Consolidate previous vendors COFER file to one DF, MainDF
    mainDF = mainDF.merge(allVendorDFs[['PO + Part Number', 'Comments', 'GSA Comments', 'Response']],on='PO + Part Number', how='left')

    # Refine consolidated Vendor DFs to update GSA Comments automatically
    finalMainDF = rf.refineDF(mainDF)

    # Rename Master COFER template file to current master COFER file
    masterCoferFileName = f"Master Cofer - {formattedTDate}.xlsx"
    masterCoferAttachmentPath = f".//Output//{masterCoferFileName}"
    os.rename(f'{outputfolder}//Master COFER Template.xlsx', f'{outputfolder}//{masterCoferFileName}')

    # Create Master Cofer file before distributing to vendors
    createFileAndTab(masterCoferAttachmentPath, finalMainDF, 'Master COFER')
    rf.formatHeader(masterCoferAttachmentPath, 'Master COFER')

    # Upload Master COFER File
    uploadMasterCoferFile(masterCoferAttachmentPath)

    # Notify CSC
    NotifyCSC = config["NotifyCSC"]
    sendemail(NotifyCSC,masterCoferAttachmentPath, masterCoferFileName);
    # with pd.ExcelWriter(masterCoferAttachmentPath,mode='a',engine='openpyxl', if_sheet_exists ='overlay') as writer:
    #     finalmainDF.to_excel(writer,sheetName = 'Master COFER', header=True, index=False)

    return finalMainDF


def splitMasterCOFERToVendor():
    for shareToVendors in config['AllVendors']:
        if shareToVendors['process'] == True:
            vendorFileName = shareToVendors['vendorFileName']
            vendorGoogleFolderName = f"COFER-{shareToVendors['vendorName']}"
            vendorName = shareToVendors['vendorName']
            sharedGoogleFolder = vendorSharedDirectory + vendorGoogleFolderName
            if vendorFileName in finalMainDF['Vendor Name'].values:
            #if vendorName in mainDF['Vendor Name'].values:
                vendorCOFERFileName = config['vendorConfig'][
                                     'VendorFilePrefix'] + ' ' + vendorFileName + ' ' + formattedTDate + '.xlsx'
                attachment = sharedGoogleFolder + "//" + vendorCOFERFileName
                finalMainDF['Response'] = ''
                filteredVendorItemsdf = finalMainDF[finalMainDF['Vendor Name'] == vendorFileName]
                filteredVendorItemsdfSorted = filteredVendorItemsdf.sort_values(by='Days on Report',ascending=False)
                createFileAndTab(attachment, filteredVendorItemsdfSorted, "Sheet1")
                print(vendorName)
            archiveFolder = f"{vendorName} COFER Archived Files"
            archivePriorMonthFile(sharedGoogleFolder, archiveFolder)
            print(f"{vendorName} COFER Archived Files completed")
    NotifyVendor = config["NotifyVendor"]
    sendemail(NotifyVendor, "","");

if __name__ == '__main__':
    copyMasterCOFERFileTemplate()
    #Load Vendor.json file for report configurations
    os.environ['today'] = today.strftime('%Y-%m-%d') #formatted date to match COGNOS date format
    #config = utils.load_json("./resources/extn/TestVendors.json")
    config = utils.load_json("./resources/extn/Vendors.json")
    if config is not None:
        print(config)
        destinationFolder = config['googleSheetConfig']['destinationFolder']
        cognosCOFERFileName = f"Master COFER Report-{today}.xlsx"

        # Download the most current cognos generated Cofer File
        downloadTheFiles(cognosCOFERFileName)

        # load the cognos Cofer file into dataframe.
        filePath = f'{destinationFolder}/{cognosCOFERFileName}'
        mainDF = GetExcelFileToDF('page', filePath)
        print(mainDF)

        #1. Go to vendor shared drive and check if the process is true to all the vendors
        vendorSharedDirectory = config['vendorConfig']['vendorSharedDirectory']
        finalMainDF = createMasterCOFERFiles(mainDF)

        #finalMainDF = pd.read_excel('.//Output//Master Cofer - 2-10-25.xlsx', sheet_name='Master COFER')
    #Distribute current Master COFER files to all vendor shared folder
    splitMasterCOFERToVendor()
