import pandas as pd
import openpyxl
from openpyxl.styles import Font,Alignment
# Sample DataFrame
# mainDF = pd.DataFrame({
#     'Response': ['Yes', '', 'No', 'Ccleared','',''],
#     'Comments': ['', 'Needs clarification', '', 'Approved','something was INVoiced','cancelled'],
#     'Supply or RO':['RODLMS','REFERRAL','WIWO','REFERRAL','WIWO','REFERRAL'],
#     'Days on Report':[10, 100, 130,1000,12,200],
#     'Remaining 856 Quantity':[0,1,0,3,56,67],
#     'Ship Status':['Shipped','Not Shipped','Shipped','Partially Shipped','Not Shipped','Not Shipped'],
#     'GSA Comments':['GSA Note - vendor submitted','IBM Action - Adjust','spoething','something','spoething','something'],
#
# })
def refineDF (mainDF):
    try:
        # Apply your conditions
        mainDF['NewWeekComment'] = mainDF.apply(
        lambda row: row['Response'] if row['Response'] != '' and (pd.isna(row['Comments']) or row['Comments'] == '') else
                    row['Comments'] if (pd.isna(row['Response']) or row['Response'] == '') and row['Comments'] != '' else
                    row['Response'] if row['Response'] != '' and row['Comments'] != '' else '',
        axis=1
        )

        print(mainDF['NewWeekComment'])
        column_to_move = mainDF.pop("NewWeekComment")

        mainDF.insert(44, "NewWeekComment", column_to_move)
        mainDF= mainDF.fillna('')
        df_with_GSAorIBM_notes = mainDF[mainDF['GSA Comments'].str.startswith(('GSA Note','IBM Action'))]
        #print(df_with_GSAorIBM_notes)
        df_without_GSAorIBM_notes = mainDF[~mainDF['GSA Comments'].str.startswith(('GSA Note','IBM Action'))]
        # print(df_without_GSAorIBM_notes)

        #1. 100 days or above and referral orders
        condition1 = (df_without_GSAorIBM_notes['Days on Report'] >= 100) & (df_without_GSAorIBM_notes['Supply or RO'] == 'REFERRAL')

        df_without_GSAorIBM_notes.loc[condition1,'GSA Comments'] = 'Vendor Action - Referral Order over 100 days, please provide updated ESD';

        # print(df_without_GSAorIBM_notes['GSA Comments'])

        #2. remaining 856 QTY Column is 0

        condition2 = df_without_GSAorIBM_notes['Remaining 856 Quantity'] == 0
        df_without_GSAorIBM_notes.loc[condition2,'GSA Comments'] = 'Vendor Action - Need to validate in FEDPAY/Submit Invoice';

        # print(df_without_GSAorIBM_notes['GSA Comments'])

        #3. Shipped status column = “Not Shipped” or “Partially Shipped” and comments column have "inv" keyword

        condition3 = ((df_without_GSAorIBM_notes['Ship Status'] == 'Not Shipped') | (df_without_GSAorIBM_notes['Ship Status'] == 'Partially Shipped')) & (df_without_GSAorIBM_notes['NewWeekComment'].str.contains('inv', case=False,na=False))
        df_without_GSAorIBM_notes.loc[condition3,'GSA Comments'] = 'Vendor Action - If Invoiced and status shows "Not Shipped", then vendor needs to ship';

        print(df_without_GSAorIBM_notes['GSA Comments'])

        #4. comments column have "cancel" keyword


        # df_without_GSAorIBM_notes['NewWeekComment']=df_without_GSAorIBM_notes['NewWeekComment'].fillna('')
        condition4 = df_without_GSAorIBM_notes['NewWeekComment'].str.contains('cancel', case=False,na=False)
        df_without_GSAorIBM_notes.loc[condition4,'GSA Comments'] = 'Vendor Action - Need further elaboration for "Cancelled" for appropriate cells.';

        print(df_without_GSAorIBM_notes['GSA Comments'])

        finalMainDF = pd.concat([df_without_GSAorIBM_notes, df_with_GSAorIBM_notes],ignore_index=True)

        # print(finalMainDF)

        finalMainDF=finalMainDF.drop(['Comments','Response'],axis=1)
        finalMainDF.rename(columns={'NewWeekComment':'Comments'}, inplace=True)
        print(finalMainDF)
        return finalMainDF
    except :
        print("Something went wrong")

def formatHeader(workbook,sheet):
    wb = openpyxl.load_workbook(workbook)
    sheet = wb[sheet]
    headerRow = sheet[1]

    for cell in headerRow:
        cell.font = Font(bold=True,color='FFFFFF')
        cell.alignment = Alignment(horizontal='center',vertical='center', wrap_text=True)
    wb.save(workbook)