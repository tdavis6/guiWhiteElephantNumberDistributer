# Import packages
from os.path import exists
import os
import sys
import sqlite3
import smtplib
import ssl
import random
import json
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

version = "v0.1.0"

dbfile = './guiWhiteElephantDistributor/data.db'
configfile = './guiWhiteElephantDistributor/config.json'

if exists("./guiWhiteElephantDistributor/"):
    print("Directory found")
else:
    os.mkdir("./guiWhiteElephantDistributor/")

if exists(dbfile):
    print("Database file found")
else:
    print("Database file not found")
    print("Creating new database file")

    open(dbfile, "x")

    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    data = cur.execute("""CREATE TABLE IF NOT EXISTS whiteElephantData (
    name TEXT,
    primaryEmail TEXT,
    number INTEGER
    );""")
    con.commit()
    con.close()

if exists(configfile):
    print("Configuration file found")
else:
    print("Configuration file not found")
    print("Creating new configuration file")

    with open(configfile, 'w') as config:
        config.writelines('{"smtpServer": "","smtpPort": "465","smtpPassword": "","fromAddress": ""}')
    print("Please completely fill config.json before running again.")
    sys.exit()

config = json.load(open('guiWhiteElephantDistributor/config.json'))


if config['smtpServer'] == "" or config['smtpPort'] == "" or config['smtpPassword'] == "" or config['fromAddress'] == "":
    print("Please completely fill config.json before running again.")
    sys.exit()

smtpServer = config['smtpServer']
smtpPort = config['smtpPort']
fromAddress = config['fromAddress']
smtpPassword = config['smtpPassword']

# main window definitions
rootWindow = tk.Tk()

window_width = 400
window_height = 300

# get the screen dimension
screen_width = rootWindow.winfo_screenwidth()
screen_height = rootWindow.winfo_screenheight()

# find the center point
center_x = int(screen_width/2 - window_width / 2)
center_y = int(screen_height/2 - window_height / 2)

rootWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

name = tk.StringVar()
primaryEmail = tk.StringVar()


#functions
def clearDB():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    cur.execute("DELETE FROM whiteElephantData")
    con.commit()
    con.close()
    print("Database cleared")
def onClosing():
    print("Main window closed")
    rootWindow.destroy()

def submitValues():
    name = nameBox.get()
    primaryEmail = primaryEmailBox.get()

    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    sqliteCommand = "INSERT INTO whiteElephantData (name, primaryEmail) VALUES (?, ?);"
    cur.execute(sqliteCommand, (name, primaryEmail))
    con.commit()
    con.close()

    nameBox.delete(0, tk.END)
    primaryEmailBox.delete(0, tk.END)
    nameBox.focus()

def assignNumbers():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
    name, primaryEmail, number = zip(*data)

    assignedNumbers = random.sample(range(len(name)), len(name))
    assignedNumbers = [x + 1 for x in assignedNumbers]

    for i, name in enumerate(name):
        cur.execute("UPDATE whiteElephantData SET number = ? WHERE name = ?;", (assignedNumbers[i], name))

    con.commit()
    con.close()

    print("Numbers assigned")

def emailNumbers():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
    name, primaryEmail, number = zip(*data)
    fullNumberList = [(name[i], number[i]) for i in range(len(name))]
    con.close()

    message = """From: noreply@tylerdavis.net\nTo: {email}\nSubject: White Elephant Number\n\n{name}, your number for White Elephant is {number}!"""
    messageFullNumberList= """From: noreply@tylerdavis.net\nTo: {email}\nSubject: Full list of White Elephant Numbers\n\nSee below for the full list of numbers:\n{fullNumberList}"""

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtpServer, smtpPort, context=context) as server:
        server.login(fromAddress, smtpPassword)
        for i in range(len(name)):
            server.sendmail(
                fromAddress,
                primaryEmail[i],
                message.format(name=name[i], email=primaryEmail[i], number=number[i])
            )
            print("Email sent to " + name[i] + " at " + primaryEmail[i])
        server.sendmail(
            fromAddress,
            primaryEmail[0],
            messageFullNumberList.format(email=primaryEmail[0],fullNumberList=fullNumberList)
        )
        print("Full Number List sent to " + name[0] + " at " + primaryEmail[0])

def listCurrentPlayers():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
    name, primaryEmail, number = zip(*data)
    con.close()

    for i in range(len(name)):
        print(name[i])


rootWindow.title("White Elephant Number Distributor")

label = ttk.Label(rootWindow, text="Welcome to the White Elephant Game!")
label.pack()

nameLabel = ttk.Label(rootWindow, text="Please enter your name below")
nameLabel.pack(padx=10, pady=10, fill='x', expand=True)
nameBox = ttk.Entry(rootWindow, textvariable=name)
nameBox.pack(padx=10, pady=10, fill='x', expand=True)
nameBox.focus()

primaryEmailLabel = ttk.Label(rootWindow, text="Please enter your email below")
primaryEmailLabel.pack(padx=10, pady=10, fill='x', expand=True)
primaryEmailBox = ttk.Entry(rootWindow, textvariable=primaryEmail)
primaryEmailBox.pack(padx=10, pady=10, fill='x', expand=True)

submitButton = ttk.Button(
   rootWindow,
   text="Submit",
   command=submitValues
)
submitButton.pack(padx=10, pady=10, fill='x', expand=True)

def openWindow():
    # remove blur and open main window
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        print("Not a Windows machine, skipping ctypes")
    finally:
        rootWindow.protocol("WM_DELETE_WINDOW", onClosing)
        rootWindow.mainloop()

menu = {}
menu['1']="Gather new players"
menu['2']="Assign and email numbers"
menu['3']="List current players"
menu['4']="Clear DB"
menu['5']="Exit"

while True:
    options=menu.keys()
    print()
    for entry in options:
        print(entry, menu[entry])

    print()
    selection=input("Please select an option:")
    if selection =='1':
        print()
        openWindow()
    elif selection =='2':
        print()
        assignNumbers()
        emailNumbers()
    elif selection =='3':
        print()
        listCurrentPlayers()
    elif selection =='4':
        print()
        clearDB()
    elif selection =='5':
        break
    else:
        print("Invalid selection")