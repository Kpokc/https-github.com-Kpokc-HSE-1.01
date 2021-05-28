
# Automated HSE Monthly invoices

 Key project milestones:

 - Using C# language, develop an app to automate monthly invoices (6 departments at the first stage).
 - Bring all recent and further departments to a standard invoicing schema.
 - Reduce necessary time consumption for invoicing from 40+ hours to roughly 1 minute.


## FAQ

### Geting tarted
#### Step 1.

Dropdown from SAP - ZGRR, ZGIR reports for the previouse month. Use 
"Export button" ** -> "Spreadsheet" -> from drop down menu select "Excel 
(In Existing XXL Format) -> "Ok" -> select radio button "Table" - > "Ok" -> 
File -> Save as -> your prefered folder. Note: To turn Off - Total button**

** - check Screenshots: 1a & 1b
#### Step 2

Dropdown or copy stock backup file. (Ask HSE admin for correct transactions)

#### Step 3

Open an app and select all three dropped excel files, press OK.
 - App will notify a user if there is something wrong with the files.
 - App will notify a user if a reversal was found.
 - App will let you know when calculations were finished.

 ** - check Screenshots: 1c

  
### Screenshots

![App Screenshot](https://via.placeholder.com/468x300?text=App+Screenshot+Here)

  
## Features

### Stock (Backup)
### Receipts
 - App will remove SAP standard colors and apply "all-around border" to all used cella.
 - Receipt report is going to be split by weeks of the month (year). The schema in two colors represented below. ** - check Screenshots: 2a 
 - Every next Receipt (example 5000121212) is separated by the grey row.
