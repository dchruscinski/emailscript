#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Reading an Excel file using Python
import pymysql
import xlrd

def emailQuery(email):
    db = pymysql.connect(host="localhost",  # your host, usually localhost
                         port=3306, # port
                         user="root",  # your username
                         passwd="admin",  # your password
                         db="web_customer_tracker")  # name of the data base

    emailToGet = email

    # you must create a Cursor object. It will let you execute all the queries you need
    cur = db.cursor()

    # Use all the SQL you like
    cur.execute("SELECT last_name FROM web_customer_tracker.customer WHERE email = %s", (emailToGet))

    # print all the first cell of all the rows
    for result in cur.fetchall():
        sendingEmail(emailToGet, result[0])

    db.close()

def sendingEmail(emailAddress, lastName):
    # import smtplib
    from O365 import Account

    credentials = ('dc5c5f85-b8ff-43fd-b964-6c145fd1cae0', 'yEjGet8pglE/hY:gS/OpFL2oeg4=v81=')
    emailAccount = Account(credentials)
    m = emailAccount.new_message()
    m.to.add(emailAddress)
    m.subject = 'Sending test e-mail'
    m.body = "Zgodnie z poleceniem miałem przesłać Twoje nazwisko: " + lastName + "." + "\n\nZ poważaniem,\nAndrzej Kiełbasa"
    m.send()

# Main program

# Give the location of the file 
loc = ("C:\\Users\\User\\Documents\\Programowanie\\2019-11-17 - Python - skrypt e-mail\\emailscript\\emaillist.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
sheet.cell_value(0, 0)

# Get every e-mail from e-mail column; start from second row
for i in range(1, sheet.nrows):
    excelEmail = sheet.cell_value(i, 2)
    emailQuery(excelEmail)

    # Extracting number of rows
        # print(sheet.nrows)

    # Extracting number of columns
        # print(sheet.ncols)

    # Exctracting all columns name
        # for i in range(sheet.ncols):
        #   print(sheet.cell_value(0, i))

    # Exctracting first column
        # for i in range(sheet.nrows):
        #   print(sheet.cell_value(i, 0))

    # Exctracting particular row value
        # print(sheet.row_values(1))