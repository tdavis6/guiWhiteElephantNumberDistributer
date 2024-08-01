# White Elephant Number Distributor

Python program to help distribute White Elephant numbers. Collects participants, assigns numbers, and notifies everyone
of their number via an email. An additional email will be sent to the first participant registered containing the
number for everyone in the database.

## Details

When running the program, there are 6 options in the main menu. Option 1 opens a GUI for collecting names and emails,
option 2 lists all the names for the participants currently in data.db, option 3 assigns everyone (new) numbers and
emails it to them, option 4 clears the entire database, option 5 reprints the selections to the console, and
option 6 exits the program. 

When option 1 is selected, the GUI will open in a full screen window. Please press escape to exit the window.

config.json and data.db will be generated on the first run.

## Set up
Set up is minimal. Clone this repository and run run.bat. Next, fill out all values in config.json. The program will 
close after 10 seconds if any value in config.json is blank. For sending via Gmail, set smtpServer = smtp.google.com,
smtpPort = 465, fromAddress to your email, and smtpPassword to an app password for the sending account.

## Requirements
- Python 3
### Package Requirements
- os 
- sys 
- sqlite3 
- smtplib 
- ssl 
- random 
- json
- tkinter
- time