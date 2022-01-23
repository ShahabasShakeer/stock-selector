from nsetools import Nse
import xlrd
import colorama

read_path = ("sample.xls")
wb = xlrd.open_workbook(read_path)
sheet = wb.sheet_by_index(0)

NSE = Nse()
print("\n")
print(colorama.Fore.LIGHTWHITE_EX, "Stocks for OHL Strategy")
print(colorama.Fore.LIGHTWHITE_EX, "-"*23)

for i in range(sheet.nrows):
    user_quote = str(sheet.cell_value(i, 0))
    quote = NSE.get_quote(user_quote)

    try:
        last_price = quote["lastPrice"]
        open_price = quote["open"]
        day_high = quote["dayHigh"]
        day_low = quote["dayLow"]
    except TypeError:
        print("PLEASE CHECK: ", user_quote)
        break

    if open_price == day_low:
        print(colorama.Fore.LIGHTWHITE_EX, user_quote, end=" ")
        print(colorama.Fore.LIGHTGREEN_EX, "BUY")
    elif open_price == day_high:
        print(colorama.Fore.LIGHTWHITE_EX, user_quote, end=" ")
        print(colorama.Fore.LIGHTRED_EX, "SELL")

print(colorama.Fore.LIGHTWHITE_EX, "")
x = input("Stock analysis is complete. Press any key to exit.")
