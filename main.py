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
import time

version = "v0.1.0"

fileDirectory = './whiteElephantNumberDistributor/'
dbfile = './whiteElephantNumberDistributor/data.db'
configfile = './whiteElephantNumberDistributor/config.json'

if exists(fileDirectory):
    print("Directory found")
else:
    os.mkdir(fileDirectory)

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
    print("\nPlease completely fill config.json before running again. Program will exit in 10 seconds.")
    time.sleep(10)
    sys.exit()

config = json.load(open(configfile))


if config['smtpServer'] == "" or config['smtpPort'] == "" or config['smtpPassword'] == "" or config['fromAddress'] == "":
    print("\nPlease completely fill config.json before running again. Program will exit in 10 seconds.")
    time.sleep(10)
    sys.exit()

smtpServer = config['smtpServer']
smtpPort = config['smtpPort']
fromAddress = config['fromAddress']
smtpPassword = config['smtpPassword']

#functions
def clearDB():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    cur.execute("DELETE FROM whiteElephantData")
    con.commit()
    con.close()
    print("Database cleared")

def assignNumbers():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    try:
        data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
        name, primaryEmail, number = zip(*data)
    except ValueError:
        print("No participants found")
    else:
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
    try:
        data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
        name, primaryEmail, number = zip(*data)
    except ValueError:
        print("No participants found")
    else:
        fullNumberList = [(name[i], number[i]) for i in range(len(name))]
        con.close()

        message = """From: noreply@tylerdavis.net\nTo: {email}\nSubject: White Elephant Number\n\n{name}, your number for White Elephant is {number}!"""
        messageFullNumberList= """From: noreply@tylerdavis.net\nTo: {email}\nSubject: Full list of White Elephant Numbers\n\nSee below for the full list of numbers:\n{fullNumberList}"""

        context = ssl.create_default_context()

        with smtplib.SMTP_SSL(smtpServer, smtpPort, context=context) as server:
            server.login(fromAddress, smtpPassword)
            server.sendmail(
                fromAddress,
                primaryEmail[0],
                messageFullNumberList.format(email=primaryEmail[0],fullNumberList=fullNumberList)
            )
            print("Full Number List sent to " + name[0] + " at " + primaryEmail[0])
            for i in range(len(name)):
                try:
                    server.sendmail(
                        fromAddress,
                        primaryEmail[i],
                        message.format(name=name[i], email=primaryEmail[i], number=number[i])
                    )
                    print("Email sent to " + name[i] + " at " + primaryEmail[i])
                except smtplib.SMTPRecipientsRefused:
                    print("Email failed to send to " + name[i] + " at " + primaryEmail[i])

def listCurrentParticipants():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    try:
        data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
        name, primaryEmail, number = zip(*data)
    except ValueError:
        print("No participants found")
    else:
        con.close()
        print("Name and email of all participants")
        for i in range(len(name)):
            print(name[i] + " | " + primaryEmail[i])

def openWindow():
    # remove blur and open main window
    if rootWindow.state() == "withdrawn":
        rootWindow.deiconify()
        print("Main window opened")
    else:
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            print("Not a Windows machine, skipping ctypes")
        finally:
            rootWindow.mainloop()
            print("Main window opened")

def onClosing():
    rootWindow.withdraw()
    print("Main window closed")

def closeWindow(event):
    rootWindow.withdraw()
    print("Main window closed")

def submitValesReturn(event):
    submitValues()

def submitValues():
    name = nameBox.get()
    primaryEmail = primaryEmailBox.get()

    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    sqliteCommand = "INSERT INTO whiteElephantData (name, primaryEmail) VALUES (?, ?);"
    cur.execute(sqliteCommand, (name, primaryEmail))
    con.commit()
    con.close()

    print("Added " + name + " with email " + primaryEmail + " to the database")

    nameBox.delete(0, tk.END)
    primaryEmailBox.delete(0, tk.END)
    nameBox.focus()

def deleteParticipant():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    while True:
        try:
            data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
            name, primaryEmail, number = zip(*data)
        except ValueError:
            print("No participants found")
            break
        else:
            print("Name and email of all participants")
            for i in range(len(name)):
                print(name[i] + " | " + primaryEmail[i])
        deleteName = input("Enter the name of who you want to remove as it appears, or press enter to exit:")
        if deleteName == "":
            break
        else:
            sqliteCommand = "DELETE FROM whiteElephantData WHERE name=?;"
            cur.execute(sqliteCommand, (deleteName,))
            print("\nDeleted " + deleteName + " from the database\n")
    con.commit()
    con.close()

def pruneParticipants():
    con = sqlite3.connect(dbfile)
    cur = con.cursor()
    try:
        data = cur.execute("SELECT * FROM whiteElephantData").fetchall()
        name, primaryEmail, number = zip(*data)
    except ValueError:
        print("No participants found")
    else:
        cur.execute("DELETE FROM whiteElephantData WHERE name=''")
        cur.execute("DELETE FROM whiteElephantData WHERE name IS NULL")
        print("Participants pruned")
    con.commit()
    con.close()

# main window definition
rootWindow = tk.Tk()

rootWindow.attributes('-fullscreen', True)
rootWindow.protocol("WM_DELETE_WINDOW", onClosing)
rootWindow.bind('<Escape>', closeWindow)
rootWindow.bind('<Return>', submitValesReturn)

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

rootWindow.withdraw()

def printMenu():
    menu = {}
    menu['0'] = "Print Menu"
    menu['1'] = "Gather new participants"
    menu['2'] = "List current participants"
    menu['3'] = "Assign numbers"
    menu['4'] = "Email numbers"
    menu['5'] = "Delete participant"
    menu['6'] = "Prune empty entries"
    menu['7'] = "Clear DB"
    menu['8'] = "Exit"

    options=menu.keys()
    for entry in options:
        print(entry, menu[entry])

print()
printMenu()

while True:
    if rootWindow.state() == "normal":
        input("Please press enter when the data collection window is closed and you are ready to continue\n")
    else:
        selection=input("Please select an option:")
        if selection =='0':
            print()
            printMenu()
        elif selection =='1':
            print()
            openWindow()
            print()
        elif selection =='2':
            print()
            listCurrentParticipants()
            print()
        elif selection =='3':
            print()
            assignNumbers()
            print()
        elif selection =='4':
            print()
            emailNumbers()
            print()
        elif selection == '5':
            print()
            deleteParticipant()
            print()
        elif selection =='6':
            print()
            pruneParticipants()
            print()
        elif selection =='7':
            print()
            clearDB()
            print()
        elif selection =='8':
            break
        else:
            print("\nInvalid selection\n")