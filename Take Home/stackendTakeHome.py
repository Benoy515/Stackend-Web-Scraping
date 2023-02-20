import pandas as pd
from openpyxl import load_workbook
import os

names, dates, sellerNames, sellerStates, sellerCounties, commissions, quantitiesPurchased, bids, lotSubTotals, bidders = [[] for i in range(10)]

for year in ["2021","2022"]:
    for file in os.listdir(f"/Users/abhaybenoy/Downloads/Land Spreadsheets/{year}"):
        if file in [".DS_Store", "12-06-21 Patterson"]:continue
        path = f"/Users/abhaybenoy/Downloads/Land Spreadsheets/{year}/{file}/"
        for sheet in os.listdir(f"/Users/abhaybenoy/Downloads/Land Spreadsheets/{year}/{file}"):
            if sheet[-4:] == "xlsx" and sheet[-6].isdigit():
                path+=sheet
                break
        else:
            print(path)
            continue
        # print(path)
        dates.append(file[:file.find(" ")])
        names.append(file[file.find(" ")+1:])
        wb = load_workbook(filename=path, data_only=True)

        contract = wb["Contract"]
        sellerNames.append(contract["B2"].value)
        sellerStates.append(contract["B5"].value)
        sellerCounties.append(contract["B6"].value)
        for i in range(17, 20):
            if contract[f"B{i}"].value == "Real Estate":
                commissions.append(contract[f"D{i}"].value)
                break
        else:commissions.append(None)

        clerking = wb["Clerking - RE"]
        for i, value in enumerate(clerking['A:A']):
            if value.value == "Tract 1":
                num = i + 1

        quantitiesPurchased.append(clerking[f"B{num}"].value)
        bids.append(clerking[f"D{num}"].value)
        lotSubTotals.append(clerking[f"F{num}"].value)
        bidders.append(clerking[f"H{num}"].value[:4] if clerking[f"H{num}"].value and clerking[f"H{num}"].value[:4].isdigit() else None)

        # print("-----------------------")


d = {"Name": names, "Date": dates, "sellerName": sellerNames, "sellerState": sellerStates, "sellerCounty": sellerCounties,
        "Comm%": commissions, "QuantityPurchased": quantitiesPurchased, "Bid": bids,
        "LotSubTotal": lotSubTotals, "Bidder (paddle#)": bidders}

for i in d:
    print(i, len(d[i]))

df = pd.DataFrame(data = d)
    
df.to_csv("auctionData.csv")
print(df)