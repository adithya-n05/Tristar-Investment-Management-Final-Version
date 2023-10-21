#Package imports
from os import listdir
import pandas as pd
import openpyxl as op
from datetime import datetime


wb_list = []
final_wb_directory = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Scripts/Positions/Version 1/BOS/BOS Upload Template (Blank).xlsx" 
final_wb_directory_save = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Scripts/Positions/Version 1/BOS/BOS Upload Template (Jul).xlsx"
path = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Statements/7. Bank Statement JUL 2023/31 July/Positions/BOS"
path = path + "/" # Adding an extra slash to show that this is a folder

# This section extracts the individual paths of all the files within the folder that the excel files are in
dir_list = listdir(path)
for i in range(len(dir_list)):
    dir_list[i] = path + dir_list[i]

# This section removes DS STORE hidden file if it exists (Useless metadata file found on most computer systems)
ds_store_path = path + ".DS_Store"
if ds_store_path in dir_list:
    dir_list.remove(ds_store_path)

# This section shows you the excel file paths that have been extracted
wb_list = dir_list
print("\nThe workbooks you have added are:")
for i in range(len(wb_list)):
    print(wb_list[i])

# This section defines lists to copy data into
Portfolio_name_list = []
Security_ID_list = []
Security_Name_list = []
Quantity_list = []
Cost_price = []
Market_Value_Orig_Currency_list = []
Market_Value_USD_list = []
Currencylist = []

# This line creates a list of all the lists so that iteration through all of them takes less lines of code
Concatenated_lists = [Portfolio_name_list, Security_ID_list, Security_Name_list, Quantity_list, Cost_price, Market_Value_Orig_Currency_list, Market_Value_USD_list, Currencylist]

# This section defines lists to hold the data that has been processed from the initial pickup from the excel file
# The purpose of this section: When pulling data from BOS sheets, the initial 11 rows are empty. Once these 11 rows have been removed, the contents of the first set of lists are moved into the second. The first set of lists are then wiped.
# Why is this done? Suppose after removing the first 11 rows of data from the first set of lists, the next sheet's data is appended to it. The next sheet's data would also have 11 rows empty, but to find the point in this list that those 11 rows exist in, will not be possible as not every sheet's length is the same. Hence, the data that has been processed is moved away to a second list, defined below, and the first list is reset for use again.
Portfolio_name_list_final = []
Security_ID_list_final = []
New_security_ID_list_final = [] # This list will fill in the gaps from the security ID list for cash accounts
Purchase_date_list_final = []
Security_Name_list_final = []
Quantity_list_final = []
Cost_price_final = []
Cost_price_final2 = []
Market_Value_Orig_Currency_list_final = []
Market_Value_USD_list_final = []
Currencylist_final = []

# This line creates a list of all the lists so that iteration through all of them takes less lines of code
Concatenated_lists_final = [Portfolio_name_list_final, New_security_ID_list_final, Purchase_date_list_final, Security_Name_list_final, Quantity_list_final, Cost_price_final2, Market_Value_Orig_Currency_list_final, Market_Value_Orig_Currency_list_final, Currencylist_final]

# This line reads the first excel file in order to instantiate a new variable that will now hold the latest record date found in the sheet for comparison later on with newer sheets.
# Why is this done? This is performed to ensure that the reco rd date for all entries in the final excel file is the latest one found in all of the sheets, which requires comparing the date values of each sheet
wb_fixinc = pd.read_excel(wb_list[0])
GreatestDate = wb_fixinc.iloc[1,1]
GreatestDate = GreatestDate[:-4] # This is to remove the time from the datetime object picked up
print(GreatestDate)

for i in wb_list:

    # This section reads each excel file one by one using pandas
    wb_fixinc = pd.read_excel(i)
    print("\nSuccessfully opened file {}".format(i))
    print(wb_fixinc) 

    # This line compares the datetime object in this new list to the GreatestDate variable containing the latest date detected so far to check if the date in this sheet is later than the date in the GreatestDate variable, and sets the GreatestDate variable to be the one in this sheet if it is found to be later than that stored in the GreatestDate
    if datetime.strptime(str(wb_fixinc.iloc[1,1][:-4]), "%d-%m-%Y %H:%M:%S") > datetime.strptime(GreatestDate, "%d-%m-%Y %H:%M:%S"):
        GreatestDate = wb_fixinc.iloc[1,1][:-3]

    # This section transfers data from specific columns in the sheet to the first set of lists
    Security_ID_list.extend(wb_fixinc[wb_fixinc.columns[3]].values.tolist())
    Security_Name_list.extend(wb_fixinc[wb_fixinc.columns[2]].values.tolist())
    Quantity_list.extend(wb_fixinc[wb_fixinc.columns[5]].values.tolist())
    Cost_price.extend(wb_fixinc[wb_fixinc.columns[6]].values.tolist())
    Market_Value_Orig_Currency_list.extend(wb_fixinc[wb_fixinc.columns[10]].values.tolist())
    Market_Value_USD_list.extend(wb_fixinc[wb_fixinc.columns[11]].values.tolist())
    Currencylist.extend(wb_fixinc[wb_fixinc.columns[4]].values.tolist())

    # This section removes empty spaces that come from first 11 rows of the files
    del Security_ID_list[0:11]
    del Security_Name_list[0:11]
    del Quantity_list[0:11]
    del Cost_price[0:11]
    del Market_Value_Orig_Currency_list[0:11]
    del Market_Value_USD_list[0:11]
    del Currencylist[0:11]

    # This section pushes data to the final list and resets temporary list that carried data from the excel files to the final file
    Security_ID_list_final.extend(Security_ID_list) 
    Security_Name_list_final.extend(Security_Name_list)
    Quantity_list_final.extend(Quantity_list)
    Cost_price_final.extend(Cost_price)
    Market_Value_Orig_Currency_list_final.extend(Market_Value_Orig_Currency_list)
    Market_Value_USD_list_final.extend(Market_Value_USD_list)
    Currencylist_final.extend(Currencylist)

    # This section resets all the initial list sets
    Portfolio_name_list = []
    Security_ID_list = []
    Security_Name_list = []
    Quantity_list = []
    Cost_price = []
    Market_Value_Orig_Currency_list = []
    Market_Value_USD_list = []
    Currencylist = []

# This sets all of the Portfolio name list to be the portfolio name given once in each sheet
for i in Security_ID_list_final:
    Portfolio_name_list_final.append(wb_fixinc.iloc[3,1])

# This line sets all of the date list to be the GreatestDate value on the sheet
for i in Security_ID_list_final:
    Purchase_date_list_final.append(GreatestDate[:-9])

# This section adds Currency tickers for cash accounts (To follow Bloomberg ticker convention)
for i in range(len(Security_Name_list_final)):
    if "Current Account" in Security_Name_list_final[i] or "External Securities" in Security_Name_list_final[i]:
        Cost_price_final2.append(0)
        if "USD" in Currencylist_final[i]:
            New_security_ID_list_final.append("USD Curncy")
        else:
            New_security_ID_list_final.append(Currencylist_final [i] + " Curncy")
    else:
        New_security_ID_list_final.append(Security_ID_list_final[i])
        Cost_price_final2.append(Cost_price_final[i])

# This line prints all lists for debugging
for i in Concatenated_lists_final:
    print(i)

# This section opens the template workbook and finds a sheet of name "template" within it
final_wb = op.load_workbook(final_wb_directory)
final_sheet = final_wb.get_sheet_by_name("template")

# This section transfers data from the final lists into the final template file and saving using directory given to save to
for i in Concatenated_lists_final:
    for j in range(2, len(i)+2):
        if i == Portfolio_name_list_final:
            final_sheet["Q" + str(j)] = i[j-2]
        if i == New_security_ID_list_final:
            final_sheet["C" + str(j)] = i[j-2]
        if i == Purchase_date_list_final:
            final_sheet["D" + str(j)] = i[j-2]
        if i == Security_Name_list_final:
            final_sheet["A" + str(j)] = i[j-2]
        if i == Quantity_list_final:
            final_sheet["O" + str(j)] = i[j-2]
        if i == Cost_price_final2:
            final_sheet["P" + str(j)] = i[j-2]

# This line saves the output
final_wb.save(final_wb_directory_save)