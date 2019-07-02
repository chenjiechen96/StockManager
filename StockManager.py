import tkinter as tk
from tkinter import messagebox
import time
from bs4 import BeautifulSoup
import requests
import xlrd
import xlutils.copy

file = 'info.xlsx'
url = 'https://au.finance.yahoo.com/quote/'
header = {'User-Agent': "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 "
                        "(KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1"}


def get_price(address, head):
    html = requests.get(address, headers=head).text
    soup = BeautifulSoup(html, 'html.parser')
    print(soup.title.string)
    money = soup.find_all('span', {'class': 'Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)'})
    money = money[0].contents
    return money[0]


def read_data(name, row, col):
    datain = xlrd.open_workbook(name)
    sheet = datain.sheet_by_name('Info')
    return sheet.cell(row, col).value


def write_data(name, row, col, val):
    datain = xlrd.open_workbook(name)
    dataout = xlutils.copy.copy(datain)
    sheet = dataout.get_sheet(0)
    sheet.write(row, col, val)
    dataout.save(name)


def find_cell(name, val):
    datain = xlrd.open_workbook(name)
    sheet = datain.sheet_by_name('Info')
    for i in range(sheet.ncols-1):
        try:
            if val == sheet.cell_value(i, 2):
                return i
        except:
            return 0
    else:
        return 0


def money_check():
    money.set(read_data(file, 1, 0))


def money_add():
    money.set(str(float(money.get())+float(money_modify.get())))
    write_data(file, 1, 0, money.get())


def money_pop():
    balance = float(money.get()) - float(money_modify.get())
    if balance < 0:
        messagebox.showerror('Error', 'Balance should not be negative!!!')
    else:
        money.set(str(balance))
        write_data(file, 1, 0, money.get())


def stock_check():
    price = get_price(url+stock.get(), header)
    stock_info.set(price)
    info = stock.get() + ' : ' + price + ' : ' + time.strftime('%Y-%m-%d', time.localtime())
    stock_infos.set(info)


def share_buy():
    try:
        balance = float(money.get())
    except:
        messagebox.showwarning('Warning', 'Please check balance first.')
    try:
        price = float(stock_info.get().replace(',', ''))
    except:
        messagebox.showwarning('Warning', 'Please check price first.')
    if balance < (float(shares.get())*price):
        messagebox.showerror('Error', 'Your money is not enough!')
    else:
        row = find_cell(file, stock.get())
        if row:
            temp = int(read_data(file, row, 1))
            temp += int(shares.get())
            write_data(file, row, 1, str(temp))
            money.set(str(balance-(float(shares.get())*price)))
            write_data(file, 1, 0, money.get())
        else:
            datain = xlrd.open_workbook(file)
            sheet = datain.sheet_by_name('Info')
            rows = sheet.nrows
            write_data(file, rows, 1, shares.get())
            write_data(file, rows, 2, stock.get())
            write_data(file, rows, 3, stock_info.get())
            write_data(file, rows, 4, time.strftime('%Y-%m-%d', time.localtime()))
            money.set(str(balance - (float(shares.get()) * price)))
            write_data(file, 1, 0, money.get())
    share_check()


def share_sell():
    try:
        balance = float(money.get())
    except:
        messagebox.showwarning('Warning', 'Please check balance first.')
    try:
        price = float(stock_info.get().replace(',', ''))
    except:
        messagebox.showwarning('Warning', 'Please check price first.')
    row = find_cell(file, stock.get())
    if row:
        temp = int(read_data(file, row, 1))
        if temp < int(shares.get()):
            messagebox.showerror('Error', 'You don\'t have enough stocks to sell.')
        elif temp == int(shares.get()):
            write_data(file, row, 1, '')
            write_data(file, row, 2, '')
            write_data(file, row, 3, '')
            write_data(file, row, 4, '')
            money.set(str(balance + (float(shares.get()) * price)))
            write_data(file, 1, 0, money.get())
        else:
            temp -= int(shares.get())
            write_data(file, row, 1, str(temp))
            money.set(str(balance-(float(shares.get())*price)))
            write_data(file, 1, 0, money.get())
    else:
        messagebox.showerror('Error', 'You don\'t have that stock.')
    share_check()


def share_check():
    datain = xlrd.open_workbook(file)
    sheet = datain.sheet_by_name('Info')
    infos = []
    for i in range(sheet.nrows):
        infos += (':'.join(sheet.row_values(i)[1:]) + '\n')
    share_bought.set(infos)


master = tk.Tk()
master.title('Stock Check')
master.geometry('700x400')

money = tk.StringVar()
money.set('Waiting for check')
money_modify = tk.StringVar()
money_x = 20
money_y = 20
MoneyBalance = tk.Label(master, width=20, height=2, bg='light gray', text='Balance')
MoneyBalance.place(x=money_x, y=money_y)
MoneyNumber = tk.Label(master, width=20, height=2, bg='light green', textvariable=money)
MoneyNumber.place(x=money_x, y=money_y + 30)
MoneyCheck = tk.Button(master, text='Check', width=8, height=1, command=money_check)
MoneyCheck.place(x=money_x + 150, y=money_y + 30)
MoneyModify = tk.Entry(master, width=20, textvariable=money_modify)
MoneyModify.place(x=money_x, y=money_y + 80)
MoneyModify.insert(tk.END, '0')
MoneyAdd = tk.Button(master, text='Add', width=5, height=1, command=money_add)
MoneyAdd.place(x=money_x + 150, y=money_y + 80)
MoneyPop = tk.Button(master, text='Pop', width=5, height=1, command=money_pop)
MoneyPop.place(x=money_x + 200, y=money_y + 80)

stock = tk.StringVar()
stock_info = tk.StringVar()
stock_info.set('Please click <Check Price>')
stock_infos = tk.StringVar()
shares = tk.StringVar()
stock_x = 20
stock_y = 200
StockFind = tk.Button(master, text='Check Price', width=10, height=1, command=stock_check)
StockFind.place(x=stock_x, y=stock_y)
ShareBuy = tk.Button(master, text='Buy Shares', width=10, height=1, command=share_buy)
ShareBuy.place(x=stock_x+140, y=stock_y+37)
ShareSell = tk.Button(master, text='Sell Shares', width=10, height=1, command=share_sell)
ShareSell.place(x=stock_x+140, y=stock_y)
ShareNum = tk.Entry(master, width=5, textvariable=shares)
ShareNum.place(x=stock_x+90, y=stock_y+40)
ShareNum.insert(tk.END, '1')
StockName = tk.Entry(master, width=10, textvariable=stock)
StockName.place(x=stock_x, y=stock_y + 40)
StockName.insert(tk.END, 'GOOG')
StockInfo = tk.Label(master, width=30, height=2, bg='light blue', textvariable=stock_info)
StockInfo.place(x=stock_x, y=stock_y + 80)
StockInfos = tk.Label(master, width=30, height=2, bg='light pink', textvariable=stock_infos)
StockInfos.place(x=stock_x, y=stock_y + 120)

share_x = 320
share_y = 20
share_bought = tk.StringVar()
ShareCheck = tk.Button(master, text='Check Owned', width=20, height=1, command=share_check)
ShareCheck.place(x=share_x, y=share_y)
ShareInfo = tk.Label(master, width=50, height=18, bg='Wheat', textvariable=share_bought)
ShareInfo.place(x=share_x, y=share_y + 40)

master.mainloop()
