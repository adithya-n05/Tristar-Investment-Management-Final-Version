# This section lists packages to be imported
from os import listdir
import pandas as pd
import openpyxl as op
from datetime import datetime

# This section asks for an input for the (1) the directory of the template file; (2) the directory of the file to save under; (3) the path with all the excel files that data will be extracted from.
wb_list = []
final_wb_directory = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Scripts/Positions/Version 2 (No ISINs)/JPM/BBG Upload Template (Blank).xlsx"
final_wb_directory_save = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Scripts/Positions/Version 2 (No ISINs)/JPM/BBG Upload Template (Sep).xlsx"
path = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/BBG PORT Enterprise (Assets)/Statements/9. Bank Statement SEP 2023/30 Sep/Positions/JPM"
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
Product_list = []
Market_value_base_currency_list = []
Market_value_position_currency_list = []
New_security_ID_list = [] # This list will fill in the gaps from the security ID list for cash accounts
New_price_list = []
Purchase_date_list = []
Security_Name_list = []
Quantity_list = []
Cost_price = []
Currencylist = []

Portfolio_name_list2 = [] 
New_security_ID_list2 = [] 
Purchase_date_list2 = []
Security_Name_list2 = []
Quantity_list2 = [] 
New_price_list2 = []

# This section creates lists of all the lists so that iteration through all of them takes less lines of code. The second list narrows down the first list of lists into one that contains solely the data to be moved into the final file for storage.
Concatenated_lists = [Portfolio_name_list, Security_ID_list, Product_list, Market_value_base_currency_list, Market_value_position_currency_list, Purchase_date_list, Security_Name_list, Quantity_list, Cost_price, Currencylist]
Concatenated_lists2 = [Portfolio_name_list, New_security_ID_list, Purchase_date_list, Security_Name_list, Quantity_list, New_price_list]
Concatenated_lists3 = [Portfolio_name_list2, New_security_ID_list2, Purchase_date_list2, Security_Name_list2, Quantity_list2, New_price_list2]

# This section defines a function for the lists to transfer data from specific columns in the sheet to the first set of lists 
def dataTransfer(wb_fixinc, Portfolio_name_list, Security_ID_list, Purchase_date_list, Security_Name_list, Quantity_list, Cost_price, Currencylist):
    Portfolio_name_list.extend(wb_fixinc[wb_fixinc.columns[0]].values.tolist())
    Security_ID_list.extend(wb_fixinc[wb_fixinc.columns[3]].values.tolist())
    Product_list.extend(wb_fixinc[wb_fixinc.columns[5]].values.tolist())
    Market_value_base_currency_list.extend(wb_fixinc[wb_fixinc.columns[8]].values.tolist())
    Market_value_position_currency_list.extend(wb_fixinc[wb_fixinc.columns[10]].values.tolist())
    Purchase_date_list.extend(list(pd.to_datetime(wb_fixinc['Price Date'])))
    Security_Name_list.extend(wb_fixinc[wb_fixinc.columns[2]].values.tolist())
    Quantity_list.extend(wb_fixinc[wb_fixinc.columns[22]].values.tolist())
    Cost_price.extend(wb_fixinc[wb_fixinc.columns[23]].values.tolist())
    Currencylist.extend(wb_fixinc[wb_fixinc.columns[11]].values.tolist())

    # This section of the code defines the removal of the final value added to each list.
    # Why is this done? This is done to remove any of the "Total" row's values from the lists as this is not required for the upload
    for k in Concatenated_lists:
        k.pop()

# This section uses the function above to copy data from each excel file into each list
for i in wb_list:

    # This section reads each excel file one by one using pandas
    print(i)
    wb_pd = pd.ExcelFile(i)
    print("Opening workbook {}".format(str(i)))

    # This section checks if the Sheet for "Fixed Income & Cash" exists. If it does, the sheet is opened and the transfer function defined above is run on it. The columns to copy data from are the same, hence the same function is used as is.
    if "Fixed Income & Cash" in wb_pd.sheet_names:
        wb_fixinc = pd.read_excel(i, sheet_name="Fixed Income & Cash", skiprows=1)
        print("\nSuccessfully opened sheet {}".format("Fixed Income & Cash"))
        print(wb_fixinc)
        dataTransfer(wb_fixinc, Portfolio_name_list, Security_ID_list, Purchase_date_list, Security_Name_list, Quantity_list, Cost_price, Currencylist)

    # This section checks if the Sheet for "Equity" exists. If it does, the sheet is opened and the transfer function defined above is run on it. The columns to copy data from are the same, hence the same function is used as is.
    if "Equity" in wb_pd.sheet_names:
        wb_fixinc = pd.read_excel(i, sheet_name="Equity", skiprows=1)
        print("\nSuccessfully opened sheet {}".format("Equity"))
        print(wb_fixinc)
        dataTransfer(wb_fixinc, Portfolio_name_list, Security_ID_list, Purchase_date_list, Security_Name_list, Quantity_list, Cost_price, Currencylist)

    # This section checks if the Sheet for "Alternative Assets" exists. If it does, the sheet is opened and the transfer function defined above is run on it. The columns to copy data from are the same, hence the same function is used as is.
    if "Alternative Assets" in wb_pd.sheet_names:
        wb_fixinc = pd.read_excel(i, sheet_name="Alternative Assets", skiprows=1)
        print("\nSuccessfully opened sheet {}".format("Alternative Assets"))
        print(wb_fixinc)
        dataTransfer(wb_fixinc, Portfolio_name_list, Security_ID_list, Purchase_date_list, Security_Name_list, Quantity_list, Cost_price, Currencylist)

# This section searches for the first date that is not stored as "NaT", not a time, and stores this as the value for the GreatestDate
# Why is this done? This is performed to ensure that the record date for all entries in the final excel file is the latest one found in all of the sheets, which requires comparing the date values of each sheet. We start here with the first date object that is detected that is not a "NaT" value
print(Purchase_date_list[9])
for i in range(len(Purchase_date_list)):
    if str(Purchase_date_list[i]) != "NaT":
        Greatestdate = datetime.strptime(str(Purchase_date_list[i]), "%Y-%m-%d %H:%M:%S")
        break

# This line compares the datetime object in the whole list to the GreatestDate variable containing the latest date detected so far to check if the date in this sheet is later than the date in the GreatestDate variable, and sets the GreatestDate variable to be the one in this sheet if it is found to be later than that stored in the GreatestDate variable
for i in range(len(Purchase_date_list)):
    if str(Purchase_date_list[i]) != "NaT":
        if datetime.strptime(str(Purchase_date_list[i]), "%Y-%m-%d %H:%M:%S") > Greatestdate:
            Greatestdate = datetime.strptime(str(Purchase_date_list[i]), "%Y-%m-%d %H:%M:%S")

# This section sets all the values of the final list that will be used as the record date to be that of the GreatestDate value
for i in range(len(Purchase_date_list)):
    Purchase_date_list[i] = Greatestdate.strftime("%Y-%m-%d")

# This section adds Currency tickers for cash accounts (To follow Bloomberg ticker convention)
for i in range(len(Security_Name_list)):
    if "Cash Account" in Security_Name_list[i]:
        New_price_list.append(0)
        if "USD" in Currencylist[i]:
            New_security_ID_list.append("USD Curncy")
        else:
            New_security_ID_list.append(Currencylist[i] + " Curncy")
    else:
        New_security_ID_list.append(Security_ID_list[i])
        New_price_list.append(Cost_price[i])

# This section sets the value of the quantity list to be that of the Market value in base currency and the price to 1 if the term "Deposit" is found in the product list.
for i in range(len(Security_Name_list)):
    if "Deposit" in Product_list[i]:
        Quantity_list[i] = Market_value_base_currency_list[i]
        New_price_list[i] = 1

for i in range(len(New_security_ID_list)):
    if str(New_security_ID_list[i]) == "nan":
        Portfolio_name_list2.append(Portfolio_name_list[i])
        New_security_ID_list2.append(New_security_ID_list[i])
        Purchase_date_list2.append(Purchase_date_list[i])
        Security_Name_list2.append(Security_Name_list[i])
        Quantity_list2.append(Quantity_list[i])
        New_price_list2.append(New_price_list[i])

# This section contains two lists, one containing a list of Security names that we want to add an ISIN to, and the second containing ISINs created internally to add
NameList = ['ASF IX PI OFF SICAV RAFI SCSP CL A', 'COATUE GROWTH V PI OFF - CL A', 'COATUE GROWTH V PI OFF DIK CL A 22', 'COATUE KONA III OFF - DIK', 'COATUE PVT O/F FEEDER FUND II LP', 'DFJ GROWTH 2016 PI OFF - 2022', 'DFJ GROWTH 2016 PI OFFSHORE - DIK', 'GIF IV PI LP OFF GAVEA CL B $250000', 'LC ASIA OFF LP(US, NON-US TAX EX)CLA', 'MPV 270 GROWTH FUND LUX APAC 4TH CL', 'NEXT GEN FUND LTD UNRES CL A S1', 'SLA II PI OFFSHORE - CLASS A', 'SLP IV PI OFF CL A DIK', 'SLP V PI OFF LP - DIK', 'VP DISLOCATION II OFF CLS A 6TH CL']
ISINList = ['PE.ARDIAN.SF.IX', 'PE.COATUE.GROWTH.V', 'PE.COATUE.GROWTH.V.DIK', 'PE.COATUE.KONA.III.DIK', 'PE.COATUE.KONA.II', 'PE.DFJ.GROWTH', 'PE.DFJ.GROWTH.DIK', 'PE.GIF.IV', 'PE.LC.ASIA', 'PE.MPV.GROWTH', 'HF.NEXT.GEN', 'PE.SLA.|I', 'PE.SLP.IV.DIK', 'PE.SLP.V.DIK', 'PD.VP.DISLOCATION']

# This section checks if the Security names that have been extracted match any of the names above and provides them with the corresponding ISIN from the ISIN list
for i in range(len(Security_Name_list2)):
    for j in range(len(NameList)):
        if NameList[j] == Security_Name_list2[i]:
            New_security_ID_list2[i] = ISINList[j]

# This line prints out the lists for debugging
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
        if i == Purchase_date_list2:
            final_sheet["D" + str(j)] = i[j-2]
        if i == Security_Name_list2:
            final_sheet["A" + str(j)] = i[j-2]
        if i == Quantity_list2:
            final_sheet["O" + str(j)] = i[j-2]
        if i == New_price_list2:
            final_sheet["P" + str(j)] = i[j-2]

# This line saves the output
final_wb.save(final_wb_directory_save)
