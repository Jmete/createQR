# createQR
Creates QR codes and a small excel sheet per employee from a master excel sheet to be used in business cards.

### Inspiration
Created due to our HR department's need to have business cards with unique QR codes per employee. This helps automate the process.

# Input
The script is intended to read an Excel sheet called "Business Cards Contact Information.xlsx"
Each line of this Excel sheet corresponds to one employee and their relevant information.
I have included a blank template, you can remove the " - Blank" part from the name. It is also expecting the columns to be the same.
However if you have custom columns, you can add them to the sheet and then modify the createQR function in the script.

# Output
The script will generate a folder called "Employees" in which will be a multitude of folders.
Each sub-folder represents one employee and contains 3 files:
- address.svg : A small QR code for the google maps location
- qr.svg : A larger QR code that contains business card information to add a contact to a mobile phone
- card_info.xlsx : This is a small Excel file that contains key text information to be used in the business card such as their names and titles in multiple languages.

The idea is that these qr codes and card_info file can be given to design companies alongside a business card design file template (such as in Adobe InDesign) to help make it easy to mass-produce business cards with custom QR codes per employee.

# Requirements
I primarily use 3 main libraries for this:
- pandas : This helps me read the Excel file easily.
- segno : This is a great library that generates the business card QR codes.
- pathlib : To help with paths.
