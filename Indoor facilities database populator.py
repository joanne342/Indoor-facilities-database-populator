import xlwings as xw
import pandas as pd
import os
import shutil
import glob

# # Hard-coded for debugging
# local_authority_code = "E07000130"

# Prompt user to enter Local Authority Code
local_authority_code = input("Enter the Local Authority Code: ")

# Desired columns to keep
desired_columns = [
    "Accessibility Type (Text)",
    "Facility ID",
    "Facility Subtype",
    "Has changing rooms?",
    "Is Refurbished?",
    "Last Updated Date",
    "Operational Status",
    "Site ID",
    "Site Name",
    "Year Built",
    "Area",
    "Length",
    "Width",
    "Dimensions Estimate",
    "Diving Boards",
    "Lanes",
    "Maximum Depth",
    "Minimum Depth",
    "Movable Floor",
    "Courts",
    "Bike Stations",
    "Badminton Courts",
    "Clearance exists - Ball / shuttlecock",
    "Rinks",
    "Stations",
    "Oval Track Lanes",
    "Disability: Parking",
    "Disability: Finding and reaching the entrance",
    "Disability: Reception Area",
    "Disability: Doorways",
    "Disability: Changing Facilities",
    "Disability: Activity Area",
    "Disability: Toilets",
    "Disability: Social Areas",
    "Disability: Spectator Areas",
    "Disability: Emergency Exits",
    "Closure Reason",
    "Pitches",
    "Artificial Sports Lighting"
]

# File names
files = [
    "sportshalls.csv",
    "squashcourts.csv",
    "studios.csv",
    "swimmingpools.csv",
    "athletics.csv",
    "healthandfitnessgym.csv",
    "indoorbowls.csv",
    "indoortenniscentre.csv",
    "artificialgrasspitches.csv"
]

# Load and process each CSV
dataframes = []
for filename in files:
    df = pd.read_csv(filename, low_memory=False)
    # Add source file column
    df["Source File"] = filename
    # Add any missing columns with NaN values to standardize columns
    for col in desired_columns:
        if col not in df.columns:
            df[col] = pd.NA
    # Filter to keep only desired columns (in order)
    df = df[desired_columns + ["Source File"]]  # include "Source File" at the end
    df = df.dropna(axis=1, how='all')  # remove columns that are all NA
    dataframes.append(df)

# Concatenate all dataframes
combined_df = pd.concat(dataframes, ignore_index=True)

# Mapping dictionary for Facility Subtype
subtype_mapping = {
    1004: "Standard Oval Outdoor",
    1005: "Mini Outdoor",
    1006: "Compact Outdoor",
    1007: "Standalone Field",
    1008: "Standalone Oval Indoor",
    1009: "Indoor Training",
    3002: "Indoor Bowls",
    4001: "Airhall",
    4002: "Airhall (seasonal)",
    4003: "Framed Fabric",
    4004: "Traditional",
    6001: "Main",
    6002: "Activity Hall",
    6003: "Barns",
    13001: "Glass-backed",
    13002: "Normal",
    12001: "Fitness Studio",
    12002: "Cycle Studio",
    7001: "Main/General",
    7002: "Leisure Pool",
    7003: "Learner/Teaching/Training",
    7004: "Diving",
    7005: "Lido",
    8001: "Sand Filled",
    8002: "Water Based",
    8003: "Long Pile Carpet",
    8004: "Sand Dressed"
}

# Convert to numeric if not already
combined_df["Facility Subtype"] = pd.to_numeric(combined_df["Facility Subtype"], errors='coerce')

# Translate codes to descriptive text
combined_df["Facility Subtype"] = combined_df["Facility Subtype"].map(subtype_mapping).fillna(combined_df["Facility Subtype"])

# Mapping dictionary for Operational Status
status_mapping = {
    1: "Planned",
    2: "Under Construction",
    3: "Operational",
    4: "Temporarily Closed",
    5: "Closed",
    7: "Indoor Bowls",
    8: "Airhall"
}

# Convert to numeric if not already
combined_df["Operational Status"] = pd.to_numeric(combined_df["Operational Status"], errors='coerce')

# Translate codes to descriptive text
combined_df["Operational Status"] = combined_df["Operational Status"].map(status_mapping).fillna(combined_df["Operational Status"])

# List of columns to map 0/1 to "No"/"Yes"
binary_columns = [
    "Has changing rooms?",
    "Is Refurbished?",
    "Dimensions Estimate",
    "Clearance exists - Ball / shuttlecock",
    "Diving Boards",
    "Movable Floor",
    "Artificial Sports Lighting"
]

# Mapping dictionary
binary_mapping = {0: "No", 1: "Yes"}

# Apply the mapping to each specified column
for column in binary_columns:
    combined_df[column] = combined_df[column].map(binary_mapping).fillna(combined_df[column])

# List of disability columns in desired order with their simplified names
disability_cols = [
    ("Disability: Parking", "Parking"),
    ("Disability: Finding and reaching the entrance", "Finding and reaching the entrance"),
    ("Disability: Reception Area", "Reception Area"),
    ("Disability: Doorways", "Doorways"),
    ("Disability: Changing Facilities", "Changing Facilities"),
    ("Disability: Activity Area", "Activity Area"),
    ("Disability: Toilets", "Toilets"),
    ("Disability: Social Areas", "Social Areas"),
    ("Disability: Spectator Areas", "Spectator Areas"),
    ("Disability: Emergency Exits", "Emergency Exits")
]

def concat_disability_standards(row):
    parts = []
    for col, name in disability_cols:
        # Check if column exists and value is 1
        if col in row and row[col] == 1:
            parts.append(name)
    return ", ".join(parts)

# Apply the function row-wise
combined_df["Disability Standard"] = combined_df.apply(concat_disability_standards, axis=1)

# Add the simple Yes/No Disability Standard column based on "is disability standard not blank?"
combined_df["Disability Access"] = combined_df["Disability Standard"].apply(lambda x: "Yes" if pd.notna(x) and str(x).strip() != "" else "No")

# Load sites.csv
sites_df = pd.read_csv("sites.csv")

# Columns to keep from sites.csv
site_columns = [
    "Site ID",
    "Site Name",
    "Site Alias",
    "Thoroughfare Name",
    "Town",
    "Postcode",
    "Telephone Number",
    "Ownership Type (Text)",
    "Management Type (Text)",
    "Easting",
    "Northing",
    "Ward Name",
    "Local Authority Name",
    "Local Authority Code",
    "Closure Reason"
]

# Keep only the desired columns
sites_df = sites_df[site_columns]

# Merge with combined_df on 'Site ID'
merged_df = pd.merge(combined_df, sites_df, on="Site ID", how="left")

# Remove rows where either Closure Reason_x or Closure Reason_y is not null
merged_df = merged_df[merged_df["Closure Reason_x"].isna() & merged_df["Closure Reason_y"].isna()]

# Now filter by local authority code
merged_df = merged_df[merged_df["Local Authority Code"] == local_authority_code]

# Get Local Authority Name for the file name
local_authority_name = merged_df["Local Authority Name"].iloc[0].replace("/", "-").strip()
output_filename = f"{local_authority_name} DEV Indoor Facilities Database v16.0.xlsm"
output_file = output_filename

# Make a copy of the original Excel file
original_file = "DEV Indoor Facilities Database v16.0.xlsm"
shutil.copyfile(original_file, output_file)

# Load CSV and sort alphabetically by Site Name
merged_df = merged_df.sort_values(by="Site Name", ascending=True)

# Open the copied Excel workbook
wb = xw.Book(output_file)

# Populate "AP_Db" sheet
try:
    all_site_id = merged_df['Site ID'].tolist()
    all_site_name = merged_df['Site Name'].tolist()
    all_site_alias = merged_df['Site Alias'].tolist()
    all_thoroughfare_name = merged_df['Thoroughfare Name'].tolist()
    all_town = merged_df['Town'].tolist()
    all_postcode = merged_df['Postcode'].tolist()
    all_telephone_number = merged_df['Telephone Number'].tolist()
    all_ownership_type = merged_df['Ownership Type (Text)'].tolist()
    all_management_type = merged_df['Management Type (Text)'].tolist()
    all_last_updated = merged_df['Last Updated Date'].tolist()
    all_easting = merged_df['Easting'].tolist()
    all_northing = merged_df['Northing'].tolist()
    all_ward_name = merged_df['Ward Name'].tolist()
    all_local_authority_name = merged_df['Local Authority Name'].tolist()
    ap_sheet = wb.sheets["AP_Db"]
    ap_sheet.api.Unprotect(Password="mancity")
    ap_sheet.range("A3").options(transpose=True).value = all_site_id
    ap_sheet.range("B3").options(transpose=True).value = all_site_name
    ap_sheet.range("C3").options(transpose=True).value = all_site_alias
    ap_sheet.range("D3").options(transpose=True).value = all_thoroughfare_name
    ap_sheet.range("E3").options(transpose=True).value = all_town
    ap_sheet.range("F3").options(transpose=True).value = all_postcode
    ap_sheet.range("G3").options(transpose=True).value = all_telephone_number
    ap_sheet.range("H3").options(transpose=True).value = all_ownership_type
    ap_sheet.range("I3").options(transpose=True).value = all_management_type
    ap_sheet.range("J3").options(transpose=True).value = all_last_updated
    ap_sheet.range("K3").options(transpose=True).value = all_easting
    ap_sheet.range("L3").options(transpose=True).value = all_northing
    ap_sheet.range("M3").options(transpose=True).value = all_ward_name
    ap_sheet.range("N3").options(transpose=True).value = all_local_authority_name
    ap_sheet.api.Protect(Password="mancity")
except Exception as e:
    print(f"Error processing sheet AP_Db: {e}")

# Mapping from source file name to sheet tab
source_file_to_sheet = {
    'sportshalls.csv': 'SH',
    'swimmingpools.csv': 'SP',
    'studios.csv': 'ST',
    'healthandfitnessgym.csv': 'HF',
    'indoorbowls.csv': 'IB',
    'squashcourts.csv': 'SQ',
    'athletics.csv': 'AT',
    'indoortenniscentre.csv': 'IT',
    'artificialgrasspitches.csv': 'AGP'
}

# Populate other sheets with columns they all have in common
for source_file, sheet_name in source_file_to_sheet.items():
    filtered_df = merged_df[merged_df['Source File'] == source_file]

    if not filtered_df.empty:
        site_ids = filtered_df['Site ID'].tolist()
        #site_names = filtered_df['Site Name'].tolist()
        facility_ID = filtered_df['Facility ID'].tolist()
        facility_subtype = filtered_df['Facility Subtype'].tolist()
        accessibility_type = filtered_df['Accessibility Type (Text)'].tolist()
        disability_access = filtered_df['Disability Access'].tolist()
        disability_standard = filtered_df['Disability Standard'].tolist()
        changing_rooms = filtered_df['Has changing rooms?'].tolist()
        year_built = filtered_df['Year Built'].tolist()
        refurbished = filtered_df['Is Refurbished?'].tolist()
        last_updated = filtered_df['Last Updated Date'].tolist()
        operational_status = filtered_df['Operational Status'].tolist()

        try:
            sheet = wb.sheets[sheet_name]
            sheet.api.Unprotect(Password="mancity")
            sheet.range("A3").options(transpose=True).value = site_ids
            #sheet.range("B3").options(transpose=True).value = site_names
            sheet.range("C3").options(transpose=True).value = facility_ID
            sheet.range("D3").options(transpose=True).value = facility_subtype
            sheet.range("E3").options(transpose=True).value = accessibility_type
            sheet.range("F3").options(transpose=True).value = disability_access
            sheet.range("G3").options(transpose=True).value = disability_standard
            sheet.range("H3").options(transpose=True).value = changing_rooms
            sheet.range("I3").options(transpose=True).value = year_built
            sheet.range("J3").options(transpose=True).value = refurbished
            sheet.range("K3").options(transpose=True).value = last_updated
            sheet.range("L3").options(transpose=True).value = operational_status

            # Add courts to IT and SQ tabs
            if sheet_name in ['IT', 'SQ']:
                courts = filtered_df['Courts'].tolist()
                sheet.range("M3").options(transpose=True).value = courts  # Assuming 'Courts' goes in column M

            # Add Width, Length, and Area to the AGP, SH, SP ad IB tabs
            if sheet_name in ['SH', 'SP', 'IB', 'AGP']:
                sheet.range("N3").options(transpose=True).value = filtered_df['Width'].tolist()
                sheet.range("O3").options(transpose=True).value = filtered_df['Length'].tolist()
                sheet.range("P3").options(transpose=True).value = filtered_df['Area'].tolist()                

            # Add Width, Length, Area, Dimensions Estimate and Bike Stations to the ST tab
            if sheet_name == 'ST':
                sheet.range("M3").options(transpose=True).value = filtered_df['Width'].tolist()
                sheet.range("N3").options(transpose=True).value = filtered_df['Length'].tolist()
                sheet.range("O3").options(transpose=True).value = filtered_df['Area'].tolist()
                sheet.range("P3").options(transpose=True).value = filtered_df['Dimensions Estimate'].tolist()
                sheet.range("Q3").options(transpose=True).value = filtered_df['Bike Stations'].tolist()

            # Add Lanes, Min Depth, Max Depth, Diving Boards and Movable Floor to the SP tab
            if sheet_name == 'SP':
                sheet.range("M3").options(transpose=True).value = filtered_df['Lanes'].tolist()
                sheet.range("Q3").options(transpose=True).value = filtered_df['Minimum Depth'].tolist()
                sheet.range("R3").options(transpose=True).value = filtered_df['Maximum Depth'].tolist()
                sheet.range("S3").options(transpose=True).value = filtered_df['Diving Boards'].tolist()
                sheet.range("T3").options(transpose=True).value = filtered_df['Movable Floor'].tolist()

            # Add Lanes, Min Depth, Max Depth, Diving Boards and Movable Floor to the SH tab
            if sheet_name == 'SH':
                sheet.range("M3").options(transpose=True).value = filtered_df['Badminton Courts'].tolist()
                sheet.range("Q3").options(transpose=True).value = filtered_df['Dimensions Estimate'].tolist()
                sheet.range("R3").options(transpose=True).value = filtered_df['Clearance exists - Ball / shuttlecock'].tolist()

            # Add Oval Track Lanes to the AT tab
            if sheet_name == 'AT':
                sheet.range("M3").options(transpose=True).value = filtered_df['Oval Track Lanes'].tolist()

            # Add Stations to the HF tab
            if sheet_name == 'HF':
                sheet.range("M3").options(transpose=True).value = filtered_df['Stations'].tolist()  

            # Add Rinks to the IB tab
            if sheet_name == 'IB':
                sheet.range("M3").options(transpose=True).value = filtered_df['Rinks'].tolist() 

            # Add Pitches and Artificial Sports Lighting to the AGP tab
            if sheet_name == 'AGP':
                sheet.range("M3").options(transpose=True).value = filtered_df['Pitches'].tolist() 
                sheet.range("Q3").options(transpose=True).value = filtered_df["Artificial Sports Lighting"].tolist() 

            sheet.api.Protect(Password="mancity")
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {e}")

# #===============
# # SCHOOLS SECTION
# #===============

# # Load the edubasealldata file dynamically
# schools_files = glob.glob('edubasealldata*.csv')
# if not schools_files:
#     raise FileNotFoundError("No edubasealldata CSV files found.")

# # Read the first CSV file into a DataFrame with ANSI encoding and low_memory set to False
# schools_df = pd.read_csv(schools_files[0], encoding='windows-1252', low_memory=False)

# # Filter schools_df based on the 'DistrictAdministrative (name)' column
# schools_df = schools_df[schools_df['DistrictAdministrative (name)'] == local_authority_name]

# # Further filter to keep only rows where 'CloseDate' is blank (NaN)
# schools_df = schools_df[schools_df['CloseDate'].isna()]

# # Select specific columns
# schools_df = schools_df[[ 
#     'URN',
#     'EstablishmentNumber',
#     'EstablishmentName',
#     'TypeOfEstablishment (name)',
#     'OpenDate',
#     'PhaseOfEducation (name)',
#     'StatutoryLowAge',
#     'StatutoryHighAge',
#     'Gender (name)',
#     'NumberOfPupils',
#     'NumberOfBoys',
#     'NumberOfGirls',
#     'Street',
#     'Locality',
#     'Address3',
#     'Town',
#     'Postcode',
#     'SchoolWebsite',
#     'TelephoneNum',
#     'HeadTitle (name)',
#     'HeadFirstName',
#     'HeadLastName',
#     'HeadPreferredJobTitle',
#     'AdministrativeWard (name)',
#     'Easting',
#     'Northing'
# ]]

# # Create or get the 'Schools' sheet
# if 'Schools' not in [sheet.name for sheet in wb.sheets]:
#     # Add the Schools sheet at the end (far right)
#     wb.sheets.add('Schools', after=wb.sheets[-1])
# else:
#     # If the sheet already exists, move it to the far right
#     schools_sheet = wb.sheets['Schools']
#     # Move it after the last sheet
#     schools_sheet.api.Move(After=wb.sheets[-1].api)

# schools_sheet = wb.sheets['Schools']

# # Define the desired header names for the Schools tab
# schools_headers = [
#     'URN', 'EstablishmentNumber', 'Name', 'Type Of Establishment', 'Open Date', 'Phase Of Education',
#     'Low Age', 'Upper Age', 'Gender', 'Pupils', 'Boys', 'Girls',
#     'Street', 'Locality', 'Address3', 'Town', 'Postcode', 'SchoolWebsite',
#     'TelephoneNum', 'HeadTitle__name_', 'HeadFirstName', 'HeadLastName',
#     'HeadPreferredJobTitle', 'AdministrativeWard__name_', 'Easting', 'Northing'
# ]

# # Set the headers manually in row 1
# schools_sheet.range('A1').value = schools_headers

# # Write the data to the 'Schools' sheet
# schools_sheet.range('A2').options(index=False, header=False).value = schools_df

# Save and close the workbook
wb.save()
wb.close()
