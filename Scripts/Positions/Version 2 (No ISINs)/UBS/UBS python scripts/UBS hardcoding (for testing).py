# This section lists packages to be imported
from os import listdir
import pandas as pd
import openpyxl as op
from datetime import datetime

# This section asks for an input for the (1) the directory of the template file; (2) the directory of the file to save under; (3) the path with all the excel files that data will be extracted from.
wb_list = []
final_wb_directory = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Scripts/Positions/Version 2 (No ISINs)/UBS/BBG Upload Template (Ready).xlsx"
final_wb_directory_save = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Scripts/Positions/Version 2 (No ISINs)/UBS/BBG Upload Template (Sep).xlsx"
path = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Statements/9. Bank Statement SEP 2023/30 Sep/Positions/UBS"
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

# This section defines lists to copy data into. Lists with a name starting with "New", exist to hold data following some processing step to the original list
Column_1_list = [] # This list is used as a landmark list, to find positions of gaps in rows that follow from the formatting
Portfolio_name_list = []
Security_ID_list = []
New_security_ID_list = [] # This list will fill in the gaps from the security ID list for cash accounts
Security_Name_list_col1 = [] # This list will store part of the security name as 2 columns contain parts of the Security name each
Security_Name_list_col2 = [] # This list will store part of the security name as 2 columns contain parts of the Security name each
Security_Name_list_finalcol = [] # This list will store the full security name as 2 columns contain parts of the Security name each that are merged into this list
Quantity_list = []
Cost_price = []
New_price_list = []
New_price_list2 = []
Currencylist = []
Market_value_list = []
Date_list = []

Portfolio_name_list2 = [] 
New_security_ID_list2 = [] 
Purchase_date_list2 = []
Security_Name_list2 = []
Quantity_list2 = [] 
New_price_list3 = []

# This section creates lists of all the lists so that iteration through all of them takes less lines of code. The second list and third list narrows down the first list of lists into one that contains solely the data to be moved into the final file for storage.
Concatenated_lists_prior = [Column_1_list, Portfolio_name_list, Security_ID_list, Security_Name_list_col1, Security_Name_list_col2, Quantity_list, Cost_price, Currencylist, Market_value_list, Date_list]
Concatenated_lists = [Column_1_list, Portfolio_name_list, Security_ID_list, Security_Name_list_finalcol, Quantity_list, Cost_price, Currencylist, Market_value_list,  Date_list]
Concatenated_lists_final = [Column_1_list, Portfolio_name_list, New_security_ID_list, Security_Name_list_finalcol, Quantity_list, New_price_list2, Date_list]
Concatenated_lists3 = [Portfolio_name_list2, New_security_ID_list2, Purchase_date_list2, Security_Name_list2, Quantity_list2, New_price_list3]

for i in wb_list:

    # This section reads each CSV file one by one using pandas
    print(i)
    wb_fixinc = pd.read_csv(i, sep=';')
    print("\nSuccessfully opened file {}".format(i))
    print(wb_fixinc)

    # This section transfers data from specific columns in the sheet to the lists
    Column_1_list.extend(wb_fixinc[wb_fixinc.columns[0]].values.tolist())
    Portfolio_name_list.extend(wb_fixinc[wb_fixinc.columns[2]].values.tolist())
    Security_ID_list.extend(wb_fixinc[wb_fixinc.columns[8]].values.tolist())
    Security_Name_list_col1.extend(wb_fixinc[wb_fixinc.columns[13]].values.tolist())
    Security_Name_list_col2.extend(wb_fixinc[wb_fixinc.columns[14]].values.tolist())
    Quantity_list.extend(wb_fixinc[wb_fixinc.columns[6]].values.tolist())
    Cost_price.extend(wb_fixinc[wb_fixinc.columns[9]].values.tolist())
    Currencylist.extend(wb_fixinc[wb_fixinc.columns[5]].values.tolist())
    Market_value_list.extend(wb_fixinc[wb_fixinc.columns[23]].values.tolist())
    Date_list.extend(wb_fixinc[wb_fixinc.columns[21]].values.tolist())

    # This line defines the length to iterate over the Column_1_list that, as mentioned above, is used as a landmark list to remove empty rows
    # Why is this done? This is needed, as data will be deleted from each of the lists where the rows are empty, hence to account for a dynamically changing list, a while loop has to be used instead of a for loop, as a for loop does not attempt to re-evaluate the length of the list after each time the length of the list is changed by deletion of a value within that list
    n=len(Column_1_list)-1

    # This section removes empty rows from each of lists based on whether "Detailed positions: Liquidity - Accounts from" is found in the landmark column 1 list, as this is the position from which a set number of empty rows begin.
    while n>0:
        if Column_1_list[n] == "Detailed positions: Liquidity - Accounts from":
            print(Column_1_list[n], n)
            for i in Concatenated_lists_prior:
                del i[n:len(i)]
            break
        n=n-1

# This section merges both of the names in the lists storing part of the security name into one list containing the final security name
for i in range(len(Security_Name_list_col1)):
    Security_Name_list_finalcol.append(Security_Name_list_col1[i] + Security_Name_list_col2[i])

# This section removes the empty header rows using key "Portfolio" as a landmark in the Portfolio name list
m=0
while m < len(Portfolio_name_list):
    if Portfolio_name_list[m] =="Portfolio":
        for k in Concatenated_lists:
            k.pop(m-1)
            k.pop(m-1)
    m=m+1

# This section removes apostrophes for thousands delimiter in UBS
for k in range(len(Quantity_list)):
    Quantity_list[k]=Quantity_list[k].replace("'","")
for k in range(len(Cost_price)):
    replace_value = Cost_price[k].replace("'","")
    Cost_price[k]=replace_value
for k in range(len(Market_value_list)):
    replace_value = Market_value_list[k].replace("'","")
    Market_value_list[k]=replace_value

# This section checks if the price contains a %. If it does, it will be removed
for i in range(len(Cost_price)):
    if "%" in str(Cost_price[i]):
        New_price_list.append(str(Cost_price[i]).replace("%", ""))
    else:
        New_price_list.append(str(Cost_price[i]))

# This section adds Currency tickers for "Current Account" (To follow Bloomberg ticker convention)
for i in range(len(Security_Name_list_finalcol)):
    if "Current Account" in Security_Name_list_finalcol[i]:
        New_price_list2.append(0)
        if "USD" in Currencylist[i]:
            New_security_ID_list.append("USD Curncy")
        else:
            New_security_ID_list.append(Currencylist[i] + " Curncy")
    else:
        New_security_ID_list.append(Security_ID_list[i])
        New_price_list2.append(New_price_list[i])


# This section searches for the first date that is not stored as "nan", not a number, and stores this as the value for the GreatestDate
# Why is this done? This is performed to ensure that the record date for all entries in the final excel file is the latest one found in all of the sheets, which requires comparing the date values of each sheet. We start here with the first date object that is detected that is not a "nan" value
for i in range(len(Date_list)):
    if str(Date_list[i]) != "nan":
        Greatestdate = datetime.strptime(str(Date_list[i]), "%d.%m.%Y")
        break

# This line compares the datetime object in the whole list to the GreatestDate variable containing the latest date detected so far to check if the date in this sheet is later than the date in the GreatestDate variable, and sets the GreatestDate variable to be the one in this sheet if it is found to be later than that stored in the GreatestDate variable
for i in range(len(Date_list)):
    if str(Date_list[i]) != "nan":
        if datetime.strptime(str(Date_list[i]), "%d.%m.%Y") > Greatestdate:
            Greatestdate = datetime.strptime(str(Date_list[i]), "%d.%m.%Y")

# This section sets all the values of the final list that will be used as the record date to be that of the GreatestDate value
for i in range(len(Date_list)):
    Date_list[i] = Greatestdate.strftime("%d.%m.%Y")

for i in range(len(New_security_ID_list)):
    if str(New_security_ID_list[i]) == "nan":
        Portfolio_name_list2.append(Portfolio_name_list[i])
        New_security_ID_list2.append(New_security_ID_list[i])
        Purchase_date_list2.append(Date_list[i])
        Security_Name_list2.append(Security_Name_list_finalcol[i])
        Quantity_list2.append(Quantity_list[i])
        New_price_list3.append(New_price_list2[i])

# This section contains two lists, one containing a list of Security names that we want to add an ISIN to, and the second containing ISINs created internally to add
NameList = ['A&Q Select SPC -A&Q Direct Access P72 SP A']
ISINList = ['HF.P72.AQ']

# This section checks if the Security names that have been extracted match any of the names above and provides them with the corresponding ISIN from the ISIN list
for i in range(len(Security_Name_list2)):
    for j in range(len(NameList)):
        if NameList[j] == Security_Name_list2[i]:
            New_security_ID_list2[i] = ISINList[j]

for i in Concatenated_lists3:
    print(i)

# This section opens the template workbook and finds a sheet of name "template" within it
final_wb = op.load_workbook(final_wb_directory)
final_sheet = final_wb.get_sheet_by_name("template")

# This section transfers data from the final lists into the final template file and saving using directory given to save to
for i in Concatenated_lists3:
    for j in range(2, len(i)+2):
        if i == Portfolio_name_list2:
            final_sheet["Q" + str(j)] = i[j-2]
        if i == New_security_ID_list2:
            final_sheet["C" + str(j)] = i[j-2]
        if i == Security_Name_list2:
            final_sheet["A" + str(j)] = i[j-2]
        if i == Quantity_list2:
            final_sheet["O" + str(j)] = i[j-2]
        if i == New_price_list3:
            final_sheet["P" + str(j)] = i[j-2]
        if i == Purchase_date_list2:
            final_sheet["D" + str(j)] = i[j-2]

# This line saves the output
final_wb.save(final_wb_directory_save)