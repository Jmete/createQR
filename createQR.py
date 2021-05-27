import pandas as pd
from pathlib import Path
import segno
from segno import helpers

def load_employee_info(filepath):
    infofile=filepath
    dtypes = {
        "Mobile Number": str,
        "Phone Number": str,
        "Phone Number EXT": str,
        "Fax": str,
        "ZIP": str,
    }
    # Load excel file
    df = pd.read_excel(infofile,engine='openpyxl',dtype=dtypes)

    # Drop rows that have null values
    df.dropna(how="all",inplace=True)

    # Strips white space from names to make it cleaner
    df["First Name"] = df["First Name"].str.strip()
    df["Last Name"] = df["Last Name"].str.strip()

    return df

def createQR(row):
    
    """
    Takes in a dict of employee information, creates:
    1) A QR code with vcard data for the employee information on the back of the business card
    2) A QR code with Google Maps Link for the address on the front of the business card
    3) An excel file with the employee information text that will go on the front of the business card
    
    It will generate a folder with the employees full name and place the 3 files above in the folder.
    """
    
    # Checks for NaN and replaces them with None
    for k in row.keys():
        if row[k] != row[k]:
            row[k] = None
    
    # Phone Extension needs a bit more of a careful check to make sure the string is correctly formatted with extension
    if row["Phone Number EXT"] is None:
        work_num = row["Phone Number"]
    else:
        work_num = row["Phone Number"]+";ext="+row["Phone Number EXT"]
    
    # Create variable for full name. We will also use this for the folder creation.
    fn = row["First Name"].strip() +" "+ row["Last Name"].strip()
    
    # Create vcard data for employee
    vcard = helpers.make_vcard_data(
        name = row["Last Name"].strip() + ";" + row["First Name"].strip(),
        displayname = fn,
        email = row["Email"],
        phone = (row["Mobile Number"], work_num),
        fax = row["Fax"],
        url = row["Web"],
#         street = row["Street"],
#         city = row["City"],
#         region = row["State"],
#         zipcode = row["ZIP"],
#         country = row["Country"],
        org = row["Company"],
        title = row["Title"]
    )
    
    # Create folder path if it does not exist
    Path(f"Employees/{fn}").mkdir(exist_ok=True)
    
    # Make the main QR code
    qr = segno.make(vcard,error='H')
    
    # Export the QR code
    qr.save(f"Employees/{fn}/qr.svg",scale=4)
    
    # Create and export the address QR code
    address_qr = segno.make(row["MapsLink"])
    address_qr.save(f"Employees/{fn}/address.svg",scale=3)

    # Create DF for employee information
    df = pd.DataFrame(row,index=[0])
    # These are the columns we need for the front business card text
    cols_info = [
        "English Name (on Card)",
        "English Title (on Card)",
        "Arabic Name (on Card)",
        "Arabic Title (on Card)",
        "Phone Number",
        "Phone Number EXT",
        "Fax",
        "Mobile Number",
        "Email"
    ]
    # Filter DF
    df = df[cols_info]
    # Export Employee information for business card to excel file
    df.to_excel(f"Employees/{fn}/card_info.xlsx")

def main():
    """
    Load in excel file to dataframe, process, and export folders with business card relevant files such as QR codes and Excel sheet with text info.
    """

    # Path to excel
    infofile = "Business Cards Contact Information.xlsx"

    # Load employee info df
    df = load_employee_info(infofile)

    # Create Employees folder if it doesn't exist
    Path(f"Employees").mkdir(exist_ok=True)

    # Loop through rows in df to generate qr code folders and excel per employee
    keys = df.columns
    for i, r in enumerate(df.values):
        row_dict = dict(zip(keys,r))
        createQR(row_dict)

if __name__ == '__main__':
    main()
    print("Completed")