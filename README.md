Read Me: MJ Brand 2024 INTERNSHIP 
-
*all code must be saved in a MACRO-ENABLED Excel workbook or it will not run*

INTERACTING WITH THE "REBOOT NOTIF" VBA CODE:
incomplete. Code is relatively pre-structured.

INTERACTING WITH THE "U&L" VBA CODE AND THE USERS & LICENSES (“U&L”) EXCEL FILE:
The file must be saved as an “Excel Macro-Enabled Workbook”, or the code will not run.

After saving the workbook as an “Excel Macro-Enabled Workbook”. 
If it is not already, the “Developer” tab will need to be added to the ribbon displayed at the top of the Application. 
To do this, click; “File” > “Options” > “Customize Ribbon”. 
Within the “Main Tabs” pop-up window, “Developer” will appear as an option next to a checkbox. 
Select the box and hit “OK”.

The name of each worksheet is equivalent to its function; adding users, replacing users, and removing users all from the datatable on the “Users & Licenses” worksheet. 

The “User & Licenses” worksheet within the workbook contains all information. 
The table in the worksheet displays the names of the users that have been entered, as well as an “x”, or multiple, across the row, to indicate which license(s) that user is tied to. 
At the far right of the table, the number of licenses per user shows. Below the table, the total number of each license, how many are in-use, and the availability on-hand can be found. 
This worksheet should always be protected to ensure no unnecessary changes are made to the datatable. 		
	
The “Users & Licenses” Worksheet contains the following information:
Names of Licenses,
Names of Users (as they are added),
What Licenses are tied to each User,
Number of Licenses per User,
Total number of Active Users,
Number of Active Users per License,
Total number of Licenses,
Number of Licenses per type,
Total number of licenses available, 
Number of Licenses available per type

In the “Users & Licenses” worksheet, in cells B4:G4, enter the number of licenses on-site, whether in-use or simply on-hand, per each type of license. 

In order for the workbook to maintain its authenticity, it is necessary to protect each worksheet individually. 
This allows for changes to be made only to pre-selected cells within each sheet. 
It locks the value of cells that either don’t need to be changed, or will automatically change regardless. 
Protecting each worksheet ensures that the information the workbook contains is true and accurate. 

In order to ensure integrity, click the “Review” tab at the top of the ribbon. 
If the worksheet is protected, “Unprotect Worksheet” will show below the icon, and no more action is required. 
If the icon reads “Protect Sheet”,  it is necessary to lock it. 
Up to this point, the “Users & Licenses” worksheet is the only sheet that should be unprotected. Lock it using the password “sheetlock”. 
If this exact password is not used to lock the sheet, the code will not run correctly. 
Do not change any other settings when setting a password.

Cell A4 on the “ADD USER”, “REPLACE USER”, and “REMOVE USER” worksheets, and cell A6 on the “REPLACE USER” worksheet are the only cells whose value can and need to be changed in order to run the VBA code. 
These cells are highlighted in yellow.
The “Users & Licenses” worksheet will adjust automatically when the code is run using the macro assigned to each “SUBMIT” button within the worksheets.

- USERS & LICENSES: ONLY FOR CHANGING LICENSE NAMES WITHIN THE TAABLE HEADINGS !!!!!
- ADD USER: Cell A4 = Name of New User, Checkboxes, “SUBMIT” Button
- REPLACE USER: Cell A4 = Name of Old User, Cell A6 = Name of New User , “SUBMIT” Button
- REMOVE USER: Cell A4 = Name of User to-be-removed, “SUBMIT” Button

What is typed in the yellow cells will not change the table in the “Users & Licenses” worksheet until the “SUBMIT” button of the active worksheet is pressed. 
The “SUBMIT” button will prompt the macro to run, and if the operation is successful, a message box will pop up.

Close the message box, and the changes should now be reflected in the table. 
Save the changes made to the workbook before closing the Excel application.

NOTE: The only part of the workbook as a whole that must be adjusted on a user-to-user basis is the number of total licenses per type of license (Step 4). 
This must be done before protecting and saving the workbook as macro-enabled (Step 5). 
This section of the workbook can be found in the “Users & Licenses” worksheet in cells B4:G4. 
