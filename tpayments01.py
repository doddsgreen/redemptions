import xlrd
import pandas as pd
from datetime import datetime, timedelta

# Create dfs
gbp_redeem_pd = pd.DataFrame(columns=['Investor code','Investor Name','Cell','Obligor','Product Code', 'Redemption amount','Subscription date','Period','Interest earned'])
usd_redeem_pd = pd.DataFrame(columns=['Investor code','Investor Name','Cell','Obligor','Product Code', 'Redemption amount','Subscription date','Period','Interest earned'])
eur_redeem_pd = pd.DataFrame(columns=['Investor code','Investor Name','Cell','Obligor','Product Code', 'Redemption amount','Subscription date','Period','Interest earned'])
zar_redeem_pd = pd.DataFrame(columns=['Investor code','Investor Name','Cell','Obligor','Product Code', 'Redemption amount','Subscription date','Period','Interest earned'])
cad_redeem_pd = pd.DataFrame(columns=['Investor code','Investor Name','Cell','Obligor','Product Code', 'Redemption amount','Subscription date','Period','Interest earned'])
switch_pd = pd.DataFrame(columns=['Investor code','Investor Name','Cell',"CCY",'Obligor','Product Code', 'Redemption amount','Subscription date','Period','Interest earned',"Next product"])

# Maturity date set
maturity_date = input("Enter maturity date (yyyy-mm-dd): ")
maturity_date_format = datetime.strptime(maturity_date, '%Y-%m-%d')

# Load redemption file
redemption_file = input("Enter path to holdings: ")
redemption_file_read = pd.DataFrame(pd.read_excel(redemption_file))

# Parse
redemption_row = 0
gbp_row = 0
usd_row = 0
eur_row = 0
cad_row = 0
zar_row = 0
switch_row = 0
while redemption_row < len(redemption_file_read):
    red_maturity_date = redemption_file_read.iloc[redemption_row, 10]
    red_maturity_date_format = datetime.strptime(red_maturity_date, '%Y-%m-%d')
    next_action = redemption_file_read.iloc[redemption_row, 20]
    if red_maturity_date_format == maturity_date_format:
        status = red_investor_code = redemption_file_read.iloc[redemption_row, 19]
        if status == "Redeemed":
            pass
        if status == "Cancelled":
            pass
        else:
            if next_action == "Withdraw":
                red_investor_code = redemption_file_read.iloc[redemption_row, 2]
                red_investor_name = redemption_file_read.iloc[redemption_row, 3]
                red_cell = redemption_file_read.iloc[redemption_row, 4]
                obligor = redemption_file_read.iloc[redemption_row, 6]
                product_code = redemption_file_read.iloc[redemption_row, 7]
                redemption_amount = redemption_file_read.iloc[redemption_row, 23]
                sub_date = redemption_file_read.iloc[redemption_row, 8]
                sub_date_format = datetime.strptime(sub_date, '%Y-%m-%d')
                sub_date_format_date = sub_date_format.date()
                period = maturity_date_format - sub_date_format
                ccy = redemption_file_read.iloc[redemption_row, 11]
                cost = redemption_file_read.iloc[redemption_row, 22]
                interest = redemption_amount - cost
                if ccy == "GBP":
                    gbp_redeem_pd.loc[gbp_row] = [red_investor_code, red_investor_name, red_cell, obligor, product_code, redemption_amount, sub_date_format_date, period, interest]
                    gbp_row = gbp_row+1
                elif ccy == "USD":
                    usd_redeem_pd.loc[usd_row] = [red_investor_code, red_investor_name, red_cell, obligor, product_code, redemption_amount, sub_date_format_date, period, interest]
                    usd_row = usd_row+1
                elif ccy == "EUR":
                    eur_redeem_pd.loc[eur_row] = [red_investor_code, red_investor_name, red_cell, obligor, product_code, redemption_amount, sub_date_format_date, period, interest]
                    eur_row = eur_row+1
                elif ccy == "ZAR":
                    zar_redeem_pd.loc[zar_row] = [red_investor_code, red_investor_name, red_cell, obligor, product_code, redemption_amount, sub_date_format_date, period, interest]
                    zar_row = zar_row+1
                elif ccy == "CAD":
                    cad_redeem_pd.loc[cad_row] = [red_investor_code, red_investor_name, red_cell, obligor, product_code, redemption_amount, sub_date_format_date, period, interest]
                    cad_row = cad_row+1
                else:
                    pass
        if next_action == "Switch":
            red_investor_code = redemption_file_read.iloc[redemption_row, 2]
            red_investor_name = redemption_file_read.iloc[redemption_row, 3]
            red_cell = redemption_file_read.iloc[redemption_row, 4]
            obligor = redemption_file_read.iloc[redemption_row, 6]
            product_code = redemption_file_read.iloc[redemption_row, 7]
            redemption_amount = redemption_file_read.iloc[redemption_row, 23]
            sub_date = redemption_file_read.iloc[redemption_row, 8]
            sub_date_format = datetime.strptime(sub_date, '%Y-%m-%d')
            sub_date_format_date = sub_date_format.date()
            period = maturity_date_format - sub_date_format
            ccy = redemption_file_read.iloc[redemption_row, 11]
            cost = redemption_file_read.iloc[redemption_row, 22]
            interest = redemption_amount - cost
            next_product = redemption_file_read.iloc[redemption_row, 20]
            switch_pd.loc[switch_row] = [red_investor_code, red_investor_name, red_cell, ccy, obligor, product_code, redemption_amount, sub_date_format_date, period, interest, next_product]
        else:
            pass
    redemption_row = redemption_row+1
    switch_row = switch_row+1

# Print
with pd.ExcelWriter("Redemptions.xlsx") as writer:
    gbp_redeem_pd.to_excel(writer, sheet_name = "GBP")
    usd_redeem_pd.to_excel(writer, sheet_name = "USD")
    eur_redeem_pd.to_excel(writer, sheet_name = "EUR")
    zar_redeem_pd.to_excel(writer, sheet_name = "ZAR")
    cad_redeem_pd.to_excel(writer, sheet_name = "CAD")
    switch_pd.to_excel(writer, sheet_name = "Switch")

