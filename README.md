# White Elephant Number Distributor

Python program to help distribute White Elephant numbers. Collects participants, assigns numbers, and notifies everyone
of their number via an email. An additional email will be sent to the first participant registered containing the
number for everyone in the database.

## Details

When option 1 is selected, the GUI will open in a full screen window. Please press escape to exit the window. When back 
in the terminal, press enter to be able to select a new option.

config.json and data.db will be generated on the first run.

### Menu Options
| Number | Function                                                                   |
|--------|----------------------------------------------------------------------------|
| 0      | Display the menu options                                                   |
| 1      | Opens a GUI to collect names and emails                                    |
| 2      | Lists the names and emails of all participants                             |
| 3      | Assigns numbers (new numbers will be assigned if there are already values) |
| 4      | Distributes all numbers via email                                          |
| 5      | Delete a participant from the database                                     |
| 6      | Prune empty rows from the database                                         |
| 7      | Clear the database                                                         |
| 8      | Exit the program                                                           |

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
- thread