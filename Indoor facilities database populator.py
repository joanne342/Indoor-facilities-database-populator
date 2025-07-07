import pandas as pd
import xlwings as xw
import math
import glob

def time_to_minutes(time_str):
    """Convert hh:mm string to total minutes. Return 0 for NaN or invalid entries."""
    if isinstance(time_str, str):
        try:
            hours, minutes = map(int, time_str.split(':'))
            return hours * 60 + minutes
        except ValueError:
            return 0  # Handle invalid hh:mm format
    return 0  # Handle NaN or non-string values

def total_time_to_hhmm(minutes):
    """Convert total minutes to HH:MM format."""
    hours = minutes // 60
    minutes = minutes % 60
    return f"{int(hours)}:{int(minutes):02d}"

# Get local authority code from user input
local_authority_code = input("Please enter the local authority code: ").strip()

#hard coded for debugging
#local_authority_code = "E06000065"

# Columns to keep for sports files
individual_sports_columns_to_keep = [
    'Accessibility Type (Text)', 'Area', 'Artificial Sports Lighting', 'Automatic Start Gate',
    'Badminton Courts', 'Bike Stations', 'Bike Wash', 'Black Trails', 'Blue Trails',
    'Clearance exists - Ball / shuttlecock', 'Closure Reason', 'Courts', 'Degree of Banking at Middle of Bends',
    'Degree of Banking at Middle of Straight', 'Dimensions Estimate', 'Disability: Activity Area',
    'Disability: Changing Facilities', 'Disability: Doorways', 'Disability: Emergency Exits',
    'Disability: Finding and reaching the entrance', 'Disability: Parking', 'Disability: Reception Area',
    'Disability: Social Areas', 'Disability: Spectator Areas', 'Disability: Toilets', 'Diving Boards',
    'Doubles', 'Easting', 'Extreme Trails', 'Facility ID', 'Facility Type', 'Facility Subtype', 'Finish Straight Length',
    'Finish Straight Width', 'Green Trails', 'Has changing rooms?', 'Holes', 'Lanes', 'Last Updated Date',
    'Length', 'Length of Black Trails', 'Length of Blue Trails', 'Length of Extreme Trails',
    'Length of Green Trails', 'Length of Red Trails', 'Length of Straights', 'Local Authority Name',
    'Maximum Depth', 'Minimum Depth', 'Movable Floor', 'Movable Wall', 'No of Persons Start Gate',
    'Northing', 'Number of Straights', 'Number of Turns', 'Operational Status', 'Oval Track Lanes',
    'Overall Width', 'Overmarked', 'Pitches', 'Radius of Turns/Bends', 'Red Trails', 'Rinks', 'Site ID',
    'Site Name', 'Skiable Length', 'Skiable Width', 'Slope Type', 'Start Gate', 'Start Hill Elevation',
    'Start Hill Width', 'Start Straight Length', 'Start Straight Width', 'Stations', 'Surface',
    'Surface Type', 'Timing System', 'Total Length', 'Tow', 'Width', 'Width of Turns/Bends', 'Year Built',
    'Year Refurbished'
]

# Columns to keep for sites.csv
sites_columns_to_keep = [
    'Site ID', 'Site Name', 'Local Authority Name', 'Local Authority Code', 'Ward Name', 
    'Car Park Capacity', 'Disability Notes', 'Email', 'Ownership Type (Text)', 'Closure Reason', 
    'Postcode','Site Alias', 'Telephone Number', 'Thoroughfare Name', 'Town', 'Website'
]

#load individual sports files
df_artificialgrasspitches = pd.read_csv('activeplacescsvs/artificialgrasspitches.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_athletics = pd.read_csv('activeplacescsvs/athletics.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_cycling = pd.read_csv('activeplacescsvs/cycling.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_golf = pd.read_csv('activeplacescsvs/golf.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_grasspitches = pd.read_csv('activeplacescsvs/grasspitches.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_healthandfitnessgym = pd.read_csv('activeplacescsvs/healthandfitnessgym.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_icerinks = pd.read_csv('activeplacescsvs/icerinks.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_indoorbowls = pd.read_csv('activeplacescsvs/indoorbowls.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_indoortenniscentre = pd.read_csv('activeplacescsvs/indoortenniscentre.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_outdoortenniscourts = pd.read_csv('activeplacescsvs/outdoortenniscourts.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_skislopes = pd.read_csv('activeplacescsvs/skislopes.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_sportshalls = pd.read_csv('activeplacescsvs/sportshalls.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_squashcourts = pd.read_csv('activeplacescsvs/squashcourts.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_studios = pd.read_csv('activeplacescsvs/studios.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)
df_swimmingpools = pd.read_csv('activeplacescsvs/swimmingpools.csv', usecols=lambda col: col in individual_sports_columns_to_keep, low_memory=False)

df_list = []

#append them to datafiles list
df_list.append(df_artificialgrasspitches)
df_list.append(df_athletics)
df_list.append(df_cycling)
df_list.append(df_golf)
df_list.append(df_grasspitches)
df_list.append(df_healthandfitnessgym)
df_list.append(df_icerinks)
df_list.append(df_indoorbowls)
df_list.append(df_indoortenniscentre)
df_list.append(df_outdoortenniscourts)
df_list.append(df_skislopes)
df_list.append(df_sportshalls)
df_list.append(df_squashcourts)
df_list.append(df_studios)
df_list.append(df_swimmingpools)

# Concatenate all individual sports DataFrames
individual_sports_df = pd.concat(df_list, ignore_index=True)

# Load the sites.csv file
sites_df = pd.read_csv('activeplacescsvs/sites.csv', usecols=sites_columns_to_keep, low_memory=False)

# Check if the local authority code exists in the sites_df
if local_authority_code not in sites_df['Local Authority Code'].values:
    raise ValueError(f"The local authority code '{local_authority_code}' does not exist in the sites data.")

# Filter sites_df to only keep rows where the 'Local Authority Code' matches the desired value
filtered_sites_df = sites_df[sites_df['Local Authority Code'] == local_authority_code]

# Check for Site IDs
site_ids = filtered_sites_df['Site ID'].unique()
if not site_ids.size:
    raise ValueError(f"No Site IDs found for local authority code '{local_authority_code}'.")

# Merge the filtered sites DataFrame with the sports DataFrame on 'Site ID', keeping filtered_sites_df on the left
merge_sites_and_individual_sports_df = pd.merge(filtered_sites_df, individual_sports_df, on='Site ID', how='left')

# Now, load the facilitytimings.csv file
facilitytimings_df = pd.read_csv('activeplacescsvs/facilitytimings.csv')

# Perform an inner merge between the facility timings and the previously merged DataFrame based on 'Facility ID'
final_df = pd.merge(merge_sites_and_individual_sports_df, facilitytimings_df, on='Facility ID', how='left')

# Filter out closed facilities
final_df = final_df[
    (final_df["Closure Reason_x"].isna()) &
    (final_df["Closure Reason_y"].isna())
]

# Load the edubasealldata file dynamically
schools_files = glob.glob('activeplacescsvs/edubasealldata*.csv')
if not schools_files:
    raise FileNotFoundError("No edubasealldata CSV files found.")

# Read the first CSV file into a DataFrame with ANSI encoding and low_memory set to False
schools_df = pd.read_csv(schools_files[0], encoding='windows-1252', low_memory=False)

# Get the first value of the 'Local Authority Name' column from final_df
local_authority_name = final_df['Local Authority Name'].iloc[0]

# Filter schools_df based on the 'DistrictAdministrative (name)' column
schools_df = schools_df[schools_df['DistrictAdministrative (name)'] == local_authority_name]

# Further filter to keep only rows where 'CloseDate' is blank (NaN)
schools_df = schools_df[schools_df['CloseDate'].isna()]

#add blank columns for AP DB to prevent later error
column_names = [
    "When form was printed / sent to WorkMobile",
    "Urban/Rural (MAPINFO)",
    "To be visited?",
    "If 'Yes' which device?",
    "Who's Going?",
    "Re-Assigned /Current Device"
]

# Adding blank columns to the DataFrame
for column in column_names:
    final_df[column] = ""

#====================
#replacements section
#====================

# Define the columns to replace 0/1 with No/Yes
binary_columns = [
    'Has changing rooms?', 'Artificial Sports Lighting', 'Clearance exists - Ball / shuttlecock',
    'Dimensions Estimate', 'Automatic Start Gate', 'Bike Wash', 'Black Trails', 'Blue Trails',
    'Diving Boards', 'Doubles', 'Extreme Trails', 'Green Trails', 'Movable Floor', 'Movable Wall',
    'Red Trails', 'Start Gate', 'Surface', 'Timing System'
]

# Replace 1 with "Yes" and 0 with "No" in the binary columns
final_df[binary_columns] = final_df[binary_columns].replace({1: "Yes", 0: "No"})

# Define the disability-related columns to modify
disability_columns = [
    'Disability: Parking',
    'Disability: Finding and reaching the entrance',
    'Disability: Reception Area',
    'Disability: Doorways',
    'Disability: Changing Facilities',
    'Disability: Activity Area',
    'Disability: Toilets',
    'Disability: Social Areas',
    'Disability: Spectator Areas',
    'Disability: Emergency Exits'
]

# Replace 1 with "Yes" and 0 with "" in the disability columns
final_df[disability_columns] = final_df[disability_columns].replace({1: "Yes", 0: ""})

# Mappings
week_day_mapping = {
    0: 'Sunday', 1: 'Monday', 2: 'Tuesday', 3: 'Wednesday', 4: 'Thursday', 
    5: 'Friday', 6: 'Saturday', 10: 'Weekend', 11: 'Monday-Friday', 12: 'Every day'
}

operational_status_mapping = {
    1: 'Planned', 2: 'Under Construction', 3: 'Operational', 4: 'Temporarily Closed', 
    5: 'Closed', 7: 'No Grass Pitches Currently Marked Out', 8: 'Not Known'
}

surface_type_mapping = {
    1: 'Acrylic', 2: 'Polymeric', 3: 'Macadam', 4: 'Artificial grass', 5: 'Textile', 
    6: 'Plastic tile', 7: 'Clay', 8: 'Acrylic/clay', 9: 'Concrete', 10: 'Other', 
    11: 'Grass', 12: 'Shale', 13: 'Nordic or Siberian Pine', 14: 'Timber', 15: 'Not Known', 
    16: 'Dolomite Stone Dust', 17: 'Various', 18: 'Synthetic', 19: 'Cinder', 20: 'Carpet Mat', 
    21: 'Diamond Mat', 22: 'Snow'
}

facility_subtype_mapping = {
    1004: 'Standard Oval Outdoor', 1005: 'Mini Outdoor', 1006: 'Compact Outdoor', 
    1007: 'Standalone Field', 1008: 'Standalone Oval Indoor', 1009: 'Indoor Training', 
    2001: 'Health and Fitness Gym', 3002: 'Indoor Bowls', 4001: 'Airhall', 
    4002: 'Airhall (seasonal)', 4003: 'Framed Fabric', 4004: 'Traditional', 
    5001: 'Adult Football', 5002: 'Junior Football 11v11', 5003: 'Cricket', 
    5004: 'Senior Rugby League', 5005: 'Junior Rugby League', 5006: 'Senior Rugby Union', 
    5007: 'Junior Rugby Union', 5008: 'Australian Rules Football', 5009: 'American Football', 
    5010: 'Hockey', 5011: 'Lacrosse', 5012: 'Rounders', 5013: 'Baseball', 
    5014: 'Softball', 5015: 'Gaelic Football', 5016: 'Shinty', 5018: 'Polo', 
    5019: 'Cycling Polo', 5020: 'Mini Soccer 7v7', 5021: 'Mini Rugby Union', 
    5022: 'Junior Football 9v9', 5023: 'Mini Soccer 5v5', 5024: 'Mini Rugby League', 
    6001: 'Main', 6002: 'Activity Hall', 6003: 'Barns', 7001: 'Main/General', 
    7002: 'Leisure Pool', 7003: 'Learner/Teaching/Training', 7004: 'Diving', 
    7005: 'Lido', 8001: 'Sand Filled', 8002: 'Water Based', 8003: 'Long Pile Carpet', 
    8004: 'Sand Dressed', 9001: 'Standard', 9002: 'Par 3', 9003: 'Driving Range', 
    10001: 'Ice Rinks', 11001: 'Outdoor Artificial', 11002: 'Outdoor Natural', 
    11003: 'Indoor', 11004: 'Indoor Endless', 12001: 'Fitness Studio', 
    12002: 'Cycle Studio', 13001: 'Glass-backed', 13002: 'Normal', 
    17001: 'Tennis Courts', 20001: 'Track - Indoor Velodrome', 20002: 'Track - Outdoor Velodrome', 
    20003: 'BMX - Race Track', 20004: 'BMX - Pump Track', 20005: 'Mountain Bike - Trails', 
    20006: 'Cycle Speedway - Track', 20007: 'Road - Closed Road Cycling Circuit', 
    33001: 'Gymnastics Hall'
}

facility_type_mapping = {
    1: 'Athletics',
    2: 'Health and Fitness Gym',
    3: 'Indoor Bowls',
    4: 'Indoor Tennis Centre',
    5: 'Grass Pitches',
    6: 'Sports Hall',
    7: 'Swimming Pool',
    8: 'Artificial Grass Pitch',
    9: 'Golf',
    10: 'Ice Rinks',
    11: 'Ski Slopes',
    12: 'Studio',
    13: 'Squash Courts',
    17: 'Outdoor Tennis Courts',
    20: 'Cycling',
    33: 'Gymnastics'
}

slope_type_mapping = {
    1: 'Nursery', 2: 'Intermediate', 3: 'Advanced', 10: 'Other'
}

# Replace mappings in the DataFrame
final_df['Facility Type'] = final_df['Facility Type'].replace(facility_type_mapping)
final_df['Facility Subtype'] = final_df['Facility Subtype'].replace(facility_subtype_mapping)
final_df['Operational Status'] = final_df['Operational Status'].replace(operational_status_mapping)
final_df['Surface Type'] = final_df['Surface Type'].replace(surface_type_mapping)
final_df['Week Day'] = final_df['Week Day'].replace(week_day_mapping)
final_df['Slope Type'] = final_df['Slope Type'].replace(slope_type_mapping)

#===========================
#end of replacements section
#===========================

# make a copy of the dataframe to stop later fragmentation warnings
final_df = final_df.copy()

# Define a function to determine the Disability Standard
def create_disability_standard_for_ap_db(row):
    return ','.join('Y' if row[col] == 'Yes' else 'N' for col in disability_columns)

# Apply the function to create the new column
final_df['Disability Standard for AP DB'] = final_df.apply(create_disability_standard_for_ap_db, axis=1)

# Define a function to determine the Disability Standard for individual sports
def create_disability_standard_for_individual_sports(row):
    # List of disability columns and their descriptive parts to concatenate
    disability_areas = {
        'Disability: Parking': 'Parking',
        'Disability: Finding and reaching the entrance': 'Finding and reaching the entrance',
        'Disability: Reception Area': 'Reception Area',
        'Disability: Doorways': 'Doorways',
        'Disability: Changing Facilities': 'Changing Facilities',
        'Disability: Activity Area': 'Activity Area',
        'Disability: Toilets': 'Toilets',
        'Disability: Social Areas': 'Social Areas',
        'Disability: Spectator Areas': 'Spectator Areas',
        'Disability: Emergency Exits': 'Emergency Exits'
    }

    # Initialize an empty list to hold the accessible areas
    accessible_areas = []

    # Iterate over the disability columns and append the area description if the value is "Yes"
    for column, description in disability_areas.items():
        if row[column] == 'Yes':
            accessible_areas.append(description)

    # Concatenate the accessible areas with a comma separator
    return ', '.join(accessible_areas)

# Apply the function to create the new column 'Disability Standard for individual sports'
final_df['Disability Standard for individual sports'] = final_df.apply(create_disability_standard_for_individual_sports, axis=1)

# Create the 'Disability Access' column by checking if 'Disability Standard for individual sports' is not blank
final_df['Disability Access'] = final_df['Disability Standard for individual sports'].apply(lambda x: 'Yes' if pd.notna(x) and x.strip() != '' else 'No')

# List of facility names and corresponding CSV files
facilities = {
    'Athletic Tracks': 'activeplacescsvs/athletics.csv',
    'Health and Fitness Suite': 'activeplacescsvs/healthandfitnessgym.csv',
    'Indoor Bowls': 'activeplacescsvs/indoorbowls.csv',
    'Indoor Tennis Centre': 'activeplacescsvs/indoortenniscentre.csv',
    'Grass Pitches': 'activeplacescsvs/grasspitches.csv',
    'Sports Hall': 'activeplacescsvs/sportshalls.csv',
    'Swimming Pool': 'activeplacescsvs/swimmingpools.csv',
    'Synthetic Turf Pitch': 'activeplacescsvs/artificialgrasspitches.csv',
    'Golf': 'activeplacescsvs/golf.csv',
    'Ice Rinks': 'activeplacescsvs/icerinks.csv',
    'Ski Slopes': 'activeplacescsvs/skislopes.csv',
    'Studios': 'activeplacescsvs/studios.csv',
    'Squash Courts': 'activeplacescsvs/squashcourts.csv',
    'Tennis': 'activeplacescsvs/outdoortenniscourts.csv',
    'Cycling': 'activeplacescsvs/cycling.csv'
}

# Dictionary to hold sets of Site IDs
site_ids_dict = {}

# Load Site IDs for each facility into the dictionary
for facility, file in facilities.items():
    df = pd.read_csv(file, usecols=['Site ID'], low_memory=False)
    site_ids_dict[facility] = set(df['Site ID'])

# Add columns to the main DataFrame
for facility, site_ids in site_ids_dict.items():
    final_df[facility] = final_df['Site ID'].apply(lambda x: 'Yes' if x in site_ids else '')

# Define the columns for activity areas
activity_area_columns = [
    'Athletic Tracks', 'Health and Fitness Suite', 'Indoor Bowls', 'Indoor Tennis Centre', 
    'Grass Pitches', 'Sports Hall', 'Swimming Pool', 'Synthetic Turf Pitch', 
    'Golf', 'Ice Rinks', 'Ski Slopes', 'Studios', 'Squash Courts', 'Tennis', 'Cycling'
]

# Create the 'No of Activity Areas' column by counting the number of 'Yes' responses
final_df['No of Activity Areas'] = final_df[activity_area_columns].apply(lambda row: row.value_counts().get('Yes', 0), axis=1)

# concetenate address together
final_df['Address'] = final_df['Thoroughfare Name'] + ", " + final_df['Town'] + "-" + final_df['Postcode']

# Create the "Open-Close Timings" column
final_df['Open-Close Timings'] = final_df['Start Time'] + ' - ' + final_df['End Time']

# append " spaces" on Car Park Capacity
final_df['Car Park Capacity'] = final_df['Car Park Capacity'].astype(int).astype(str) + " spaces"

# Remove the time part from the "Last checked" column and keep only the date
final_df['Last Updated Date'] = pd.to_datetime(final_df['Last Updated Date']).dt.strftime('%d/%m/%Y')

#-------------------------------

# Convert 'Start Time' and 'End Time' to timedelta and create new columns
final_df['start time in timedelta'] = pd.to_timedelta(final_df['Start Time'] + ':00')
final_df['end time in timedelta'] = pd.to_timedelta(final_df['End Time'] + ':00')

# Calculate the duration
final_df['Duration'] = final_df['end time in timedelta'] - final_df['start time in timedelta']

# Check for NaN values before formatting
final_df['Duration'] = final_df['Duration'].where(final_df['Duration'].notna(), pd.Timedelta(0))

# Convert the duration to HH:MM format
final_df['Duration'] = final_df['Duration'].dt.components.apply(
    lambda x: f"{int(x['hours']):d}:{int(x['minutes']):02d}", axis=1
)

# List of acceptable week day values
acceptable_days = {
    'Mon': ['Monday', 'Monday-Friday', 'Every day'],
    'Tue': ['Tuesday', 'Monday-Friday', 'Every day'],
    'Wed': ['Wednesday', 'Monday-Friday', 'Every day'],
    'Thu': ['Thursday', 'Monday-Friday', 'Every day'],
    'Fri': ['Friday', 'Monday-Friday', 'Every day'],
    'Sat': ['Saturday', 'Weekend', 'Every day'],
    'Sun': ['Sunday', 'Weekend', 'Every day'],
}

# List of columns to be added
columns_to_add = [
    'Mon 17:00-22:00', 'Tues 17:00-22:00', 'Wed 17:00-22:00', 'Thurs 17:00-22:00', 
    'Fri 17:00-22:00', 'Sat 09:30-17:00', 'Sun 09:00-14:30', 'Sun 17:00-19:30',
    
    'Mon 12:00-13:30', 'Mon 16:00-22:00', 'Tues 12:00-13:30', 'Tues 16:00-22:00',
    'Wed 12:00-13:30', 'Wed 16:00-22:00', 'Thurs 12:00-13:30', 'Thurs 16:00-22:00',
    'Fri 12:00-13:30', 'Fri 16:00-22:00', 'Sat 09:00-16:00', 'Sun 09:00-16:30'
]

# Adding the columns to the DataFrame with NaN values (or any default value)
for col in columns_to_add:
    final_df[col] = pd.NA  # Initialize with NaN (or any default value)

# Function to extract the day from the column name
def extract_day(column_name):
    return column_name.split(' ')[0][:3]

# Function to extract the start time from the column name
def extract_start_time(column_name):
    time_range = column_name.split(' ')[1]
    return time_range.split('-')[0]

# Function to extract the end time from the column name
def extract_end_time(column_name):
    time_range = column_name.split(' ')[1]
    return time_range.split('-')[1]

# Creating a dictionary to store start and end times for each column
start_times = {col: extract_start_time(col) for col in columns_to_add}
end_times = {col: extract_end_time(col) for col in columns_to_add}

# Function to calculate the overlap in timedelta
def calculate_overlap(row, col):
    # Get the extracted day
    day = extract_day(col)

    # Check if the day is acceptable
    if row['Week Day'] not in acceptable_days.get(day, []):
        return pd.NA  # Leave the cell blank if not an acceptable value

    # Convert the start and end times from the column name into timedelta
    start_time = pd.to_timedelta(start_times[col] + ':00')
    end_time = pd.to_timedelta(end_times[col] + ':00')

    # Latest start time between the "start time in timedelta" column and the start_time from column name
    latest_start = max(row['start time in timedelta'], start_time)
    
    # Earliest end time between the "end time in timedelta" column and the end_time from column name
    earliest_end = min(row['end time in timedelta'], end_time)
    
    # Calculate the overlap duration by subtracting latest_start from earliest_end
    overlap = earliest_end - latest_start
    
    # Ensure that if there's no overlap, return a timedelta of 0
    overlap_duration = overlap if overlap > pd.Timedelta(0) else pd.Timedelta(0)
    
    # Convert the overlap duration to hh:mm format
    hours, remainder = divmod(overlap_duration.total_seconds(), 3600)
    minutes = remainder // 60
    return f"{int(hours):02}:{int(minutes):02}"

# Apply the overlap calculation to each row and each relevant column
for col in columns_to_add:
    final_df[col] = final_df.apply(lambda row: calculate_overlap(row, col), axis=1)

#=================================================

# Specify the relevant columns for summing (for SH Peak Periods)
sh_columns_to_sum = [
    'Mon 17:00-22:00', 'Tues 17:00-22:00', 'Wed 17:00-22:00',
    'Thurs 17:00-22:00', 'Fri 17:00-22:00', 'Sat 09:30-17:00',
    'Sun 09:00-14:30', 'Sun 17:00-19:30'
]

# Specify the relevant columns for summing
sp_columns_to_sum = ['Mon 12:00-13:30', 'Mon 16:00-22:00', 'Tues 12:00-13:30', 
                     'Tues 16:00-22:00', 'Wed 12:00-13:30', 'Wed 16:00-22:00', 
                     'Thurs 12:00-13:30', 'Thurs 16:00-22:00', 'Fri 12:00-13:30', 
                     'Fri 16:00-22:00', 'Sat 09:00-16:00', 'Sun 09:00-16:30']

# Calculate total minutes for each SH row
final_df['SH Peak Period Total'] = final_df[sh_columns_to_sum].apply(
    lambda row: sum(time_to_minutes(row[col]) for col in sh_columns_to_sum),
    axis=1
)

# Calculate total minutes for each SP row
final_df['SP Peak Period Total'] = final_df[sp_columns_to_sum].apply(
    lambda row: sum(time_to_minutes(row[col]) for col in sp_columns_to_sum),
    axis=1
)

# List of columns to modify
columns_to_modify = [
    "Mon 17:00-22:00", "Tues 17:00-22:00", "Wed 17:00-22:00", 
    "Thurs 17:00-22:00", "Fri 17:00-22:00", "Sat 09:30-17:00", 
    "Sun 09:00-14:30", "Sun 17:00-19:30", "Mon 12:00-13:30", 
    "Mon 16:00-22:00", "Tues 12:00-13:30", "Tues 16:00-22:00", 
    "Wed 12:00-13:30", "Wed 16:00-22:00", "Thurs 12:00-13:30", 
    "Thurs 16:00-22:00", "Fri 12:00-13:30", "Fri 16:00-22:00", 
    "Sat 09:00-16:00", "Sun 09:00-16:30"
]

# Replace "00:00" with an empty string in the specified columns
final_df[columns_to_modify] = final_df[columns_to_modify].replace("00:00", "")

# Convert total minutes back to hh:mm format
final_df['SH Peak Period Total'] = final_df['SH Peak Period Total'].apply(total_time_to_hhmm)
final_df['SP Peak Period Total'] = final_df['SP Peak Period Total'].apply(total_time_to_hhmm)


#-----------------------------------

# Calculate 'PP Total for Comm Use' - Identical to 'Peak Period Total' if Accessibility is 'Pay and Play' or 'Sports Club / Community Association', otherwise blank

final_df['SH PP Total for Comm Use'] = final_df.apply(
    lambda row: row['SH Peak Period Total'] if row['Accessibility Type (Text)'] in ['Pay and Play', 'Sports Club / Community Association'] else '', 
    axis=1
)

final_df['SP PP Total for Comm Use'] = final_df.apply(
    lambda row: row['SP Peak Period Total'] if row['Accessibility Type (Text)'] in ['Pay and Play', 'Sports Club / Community Association'] else '', 
    axis=1
)

# populate 'Person minutes in pool' - 68 if leisure pool, otherwise 64
final_df['Person minutes in pool'] = final_df['Facility Subtype'].apply(
    lambda x: 68 if 'Leisure Pool' in x else 64
)

# populate 'PP Pool capacity (visits)': ('Area' / 6) * ('Peak Period Total' * 60) / 'Person minutes in pool'
final_df['PP Pool capacity (visits)'] = round(
    (final_df['Area'] / 6) * 
    (final_df['SP Peak Period Total'].apply(time_to_minutes)) / 
    final_df['Person minutes in pool']
)

# populate "PP Pool capacity (visits) CommUse": equal to 'PP Pool capacity (visits)' if accessibility type is 'Pay and Play' or 'Sports Club / Community Association', otherwise blank
final_df['PP Pool capacity (visits) CommUse'] = final_df.apply(
    lambda row: row['PP Pool capacity (visits)'] if row['Accessibility Type (Text)'] in ['Pay and Play', 'Sports Club / Community Association'] else '', 
    axis=1
)

# Populate "Large Pool (over 100m2)"
final_df['Large Pool (over 100m2)'] = final_df['Area'].apply(lambda x: "Yes" if x > 100 else "No")

#--------------------------

# Populate "Calculated courts from Area"
def calculate_courts_from_area(area, sh_calc_areas):
    if pd.isna(area) or area == '':
        return 0
    for i in range(len(sh_calc_areas) - 1, -1, -1):
        if area >= sh_calc_areas[i][0]:
            return sh_calc_areas[i][1]
    return 0

# Define the SH_calc_areas table
sh_calc_areas = [
    (0, 0), (180, 1), (324, 2), (486, 3), (594, 4), (810, 5), (918, 6),
    (1134, 7), (1221, 8), (1377, 9), (1530, 10), (1683, 11), (1782, 12),
    (2079, 14), (2376, 16), (2970, 20), (3564, 24)
]

# Apply the function to calculate the new column
final_df['Calculated courts from Area'] = final_df['Area'].apply(lambda x: calculate_courts_from_area(x, sh_calc_areas))

#----------------------------------------

# Populate "Main or Ancillary Hall?"
def determine_main_or_ancillary(row):
    if row['Clearance exists - Ball / shuttlecock'] == 'no':
        return 'Ancillary'
    if pd.isna(row['Badminton Courts']) or row['Badminton Courts'] == '':
        return 'Ancillary'
    min_courts = min(
        row['Badminton Courts'] if not pd.isna(row['Badminton Courts']) else float('inf'), 
        row['Calculated courts from Area']
    )
    if min_courts < 3:
        return 'Ancillary'
    return 'Main'

# Apply the function to create the "Main or Ancillary Hall?" column
final_df['Main or Ancillary Hall?'] = final_df.apply(determine_main_or_ancillary, axis=1)

#--------------------------------------

# Populate "Estimated courts"
def calculate_estimated_courts(row):
    if row['Main or Ancillary Hall?'] == 'Main':
        return row['Calculated courts from Area']
    else:
        if not pd.isna(row['Area']) and row['Area'] != 0:
            return math.floor(row['Area'] * 1.6 / 144)
        elif not pd.isna(row['Badminton Courts']) and row['Badminton Courts'] != 0:
            return math.floor(row['Badminton Courts'] * 1.6)
        else:
            return 0

# Apply the function to create the "Estimated courts" column
final_df['Estimated courts'] = final_df.apply(calculate_estimated_courts, axis=1)

# populate "Equivalent courts"
def calculate_equivalent_courts(row):
    if row['Main or Ancillary Hall?'] == 'Main':
        if row['Calculated courts from Area'] == 0:
            return row['Badminton Courts']
        else:
            if pd.isna(row['Badminton Courts']) or row['Badminton Courts'] == '':
                return row['Calculated courts from Area']
            else:
                return min(row['Badminton Courts'], row['Calculated courts from Area'])
    else:
        return row['Estimated courts']

# Apply the function to create the "Equivalent courts" column
final_df['Equivalent courts'] = final_df.apply(calculate_equivalent_courts, axis=1)

#--------------------------------------

# populate "Equivalent court PP hours" by multiplying 'Peak Period Total' with 'Equivalent courts'
def calculate_equivalent_court_pp_hours(row):
    peak_period_total_minutes = time_to_minutes(row['SH Peak Period Total'])
    equivalent_courts = row['Equivalent courts']
    total_minutes = peak_period_total_minutes * equivalent_courts
    return total_time_to_hhmm(total_minutes)

# Apply the function to create the "Equivalent court PP hours" column
final_df['Equivalent court PP hours'] = final_df.apply(calculate_equivalent_court_pp_hours, axis=1)

#--------------------------------------

# Function to create "Equivalent courts for Comm Use" column
def create_equivalent_courts_for_comm_use(row):
    if row['Accessibility Type (Text)'] in ['Pay and Play', 'Sports Club / Community Association']:
        return math.floor(row['Equivalent courts'])
    else:
        return ''

# Apply the function to create the new column
final_df['Equivalent courts for Comm Use'] = final_df.apply(create_equivalent_courts_for_comm_use, axis=1)

#--------------------------------------

# Function to create "Equivalent court for Comm Use PP hours" column
def create_equivalent_court_for_comm_use_pp_hours(row):
    pp_total_for_comm_use_minutes = time_to_minutes(row['SH PP Total for Comm Use'])
    equivalent_courts_for_comm_use = row['Equivalent courts for Comm Use']
    
    if pp_total_for_comm_use_minutes == 0 or equivalent_courts_for_comm_use == '':
        return ''
    
    total_minutes = pp_total_for_comm_use_minutes * equivalent_courts_for_comm_use
    return total_time_to_hhmm(total_minutes)

# Apply the function to create the "Equivalent court for Comm Use PP hours" column
final_df['Equivalent court for Comm Use PP hours'] = final_df.apply(create_equivalent_court_for_comm_use_pp_hours, axis=1)

#------------------------------------

# Populate the "on Site? columns
final_df['100m2 pool on Site? working out'] = (final_df['Site ID'].astype(str) + final_df['Large Pool (over 100m2)'])
final_df['Main Hall on Site? working out'] = final_df['Site ID'].astype(str) + final_df['Main or Ancillary Hall?']

final_df['100m2 pool on Site?'] = final_df.apply(lambda row: 'Yes' if f"{row['Site ID']}Yes" in final_df['100m2 pool on Site? working out'].values else 'No', axis=1)
final_df['Main Hall on Site?'] = final_df.apply(lambda row: 'Yes' if f"{row['Site ID']}Main" in final_df['Main Hall on Site? working out'].values else 'No', axis=1)

#--------------------------------------

#populate 'Weekly Duration' by multiplying 'Duration' by 2 for 'Weekend', 5 for 'Monday-Friday', 7 for 'Every day', otherwise by 1
def calculate_weekly_duration(row):
    duration_minutes = time_to_minutes(row["Duration"])
    if row["Week Day"] == "Weekend":
        return total_time_to_hhmm(duration_minutes * 2)
    elif row["Week Day"] == "Monday-Friday":
        return total_time_to_hhmm(duration_minutes * 5)
    elif row["Week Day"] == "Every day":
        return total_time_to_hhmm(duration_minutes * 7)
    else:
        return row["Duration"]

final_df["Weekly Duration (hh:mm)"] = final_df.apply(calculate_weekly_duration, axis=1)
final_df["Weekly Duration (hh:mm)"] = final_df.apply(calculate_weekly_duration, axis=1)

#--------------------------------------

# Define the custom order for 'Week Day'
week_day_order = [
    "Saturday",
    "Sunday",
    "Weekend",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday", 
    "Monday-Friday",
    "Every day"
]

# Convert 'Week Day' to a categorical type with the specified order
final_df['Week Day'] = pd.Categorical(final_df['Week Day'], categories=week_day_order, ordered=True)

# Sort by 'Site Name', 'Facility ID', and 'Week Day'
final_df.sort_values(by=["Site Name", "Facility ID", "Week Day"], inplace=True)

#---------------------------------------

#add blank Urban/Rural column
final_df['Urban/Rural'] = ''

#---------------------------------------

# Step 1: Drop duplicates based on Site ID and Site Name
unique_sites = final_df[['Site ID', 'Site Name', 'Easting', 'Northing']].drop_duplicates()

# Step 2: Sort by Northing (descending) and Easting (descending)
unique_sites = unique_sites.sort_values(by=['Northing', 'Easting'], ascending=[False, False])

# Step 3: Assign sequential numbers
unique_sites['KKP Ref'] = range(1, len(unique_sites) + 1)

# Step 4: Merge back to the original DataFrame
final_df = final_df.merge(unique_sites[['Site ID', 'Site Name', 'KKP Ref']], on=['Site ID', 'Site Name'], how='left')

#------------------------------------------

# Filter final_df for rows where "Facility Type" matches
ap_db = final_df.copy()
sp_hours_df = final_df[final_df["Facility Type"] == "Swimming Pool"].copy()
sh_hours_df = final_df[final_df["Facility Type"] == "Sports Hall"].copy()
asp_times_df = final_df[final_df["Facility Type"] == "Swimming Pool"].copy()
ash_times_df = final_df[final_df["Facility Type"] == "Sports Hall"].copy()
artificialgrasspitches_df = final_df[final_df["Facility Type"] == "Artificial Grass Pitch"].copy()
athletics_df = final_df[final_df["Facility Type"] == "Athletics"].copy()
cycling_df = final_df[final_df["Facility Type"] == "Cycling"].copy()
golf_df = final_df[final_df["Facility Type"] == "Golf"].copy()
grasspitches_df = final_df[final_df["Facility Type"] == "Grass Pitches"].copy()
healthandfitnessgym_df = final_df[final_df["Facility Type"] == "Health and Fitness Gym"].copy()
icerinks_df = final_df[final_df["Facility Type"] == "Ice Rinks"].copy()
indoorbowls_df = final_df[final_df["Facility Type"] == "Indoor Bowls"].copy()
indoortenniscentre_df = final_df[final_df["Facility Type"] == "Indoor Tennis Centre"].copy()
outdoortenniscourts_df = final_df[final_df["Facility Type"] == "Outdoor Tennis Courts"].copy()
skislopes_df = final_df[final_df["Facility Type"] == "Ski Slopes"].copy()
sportshalls_df = final_df[final_df["Facility Type"] == "Sports Hall"].copy()
squashcourts_df = final_df[final_df["Facility Type"] == "Squash Courts"].copy()
studios_df = final_df[final_df["Facility Type"] == "Studio"].copy()
swimmingpools_df = final_df[final_df["Facility Type"] == "Swimming Pool"].copy()

ap_db = ap_db[[
    "Site ID",
    "Site Name",
    "Site Alias",
    "Telephone Number",
    "Town",
    "Postcode",
    "Address",
    "Email",
    "Ownership Type (Text)",
    "Website",
    "Car Park Capacity",
    "Disability Notes",
    "Disability Standard for AP DB",
    "Last Updated Date",
    "Easting",
    "Northing",
    "Ward Name",
    "Local Authority Name",
    "Athletic Tracks",
    "Health and Fitness Suite",
    "Indoor Bowls",
    "Indoor Tennis Centre",
    "Grass Pitches",
    "Sports Hall",
    "Swimming Pool",
    "Synthetic Turf Pitch",
    "Golf",
    "Ice Rinks",
    "Ski Slopes",
    "Studios",
    "Squash Courts",
    "Tennis",
    "Cycling",
    "No of Activity Areas",
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
    "When form was printed / sent to WorkMobile",
    "Urban/Rural (MAPINFO)",
    "To be visited?",
    "If 'Yes' which device?",
    "Who's Going?",
    "Re-Assigned /Current Device",
    "KKP Ref"

    
]]

sp_hours_df = sp_hours_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Week Day",
    "Open-Close Timings",
    "Accessibility Type (Text)",
    "Start Time",
    "End Time",
    "Duration",
    "Ward Name",
    "Local Authority Name",
    "Mon 12:00-13:30",
    "Mon 16:00-22:00",
    "Tues 12:00-13:30",
    "Tues 16:00-22:00",
    "Wed 12:00-13:30",
    "Wed 16:00-22:00",
    "Thurs 12:00-13:30",
    "Thurs 16:00-22:00",
    "Fri 12:00-13:30",
    "Fri 16:00-22:00",
    "Sat 09:00-16:00",
    "Sun 09:00-16:30",
    "SP Peak Period Total",
    "SP PP Total for Comm Use",
    "Facility Subtype",
    "Area",
    "Person minutes in pool",
    "PP Pool capacity (visits)",
    "PP Pool capacity (visits) CommUse",
    "Large Pool (over 100m2)",
    "100m2 pool on Site? working out",
    "100m2 pool on Site?"
]]

sh_hours_df = sh_hours_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Week Day",
    "Open-Close Timings",
    "Accessibility Type (Text)",
    "Start Time",
    "End Time",
    "Duration",
    "Ward Name",
    "Local Authority Name",
    "Mon 17:00-22:00",
    "Tues 17:00-22:00",
    "Wed 17:00-22:00",
    "Thurs 17:00-22:00",
    "Fri 17:00-22:00",
    "Sat 09:30-17:00",
    "Sun 09:00-14:30",
    "Sun 17:00-19:30",
    "SH Peak Period Total",
    "SH PP Total for Comm Use",
    "Clearance exists - Ball / shuttlecock",
    "Area",
    "Badminton Courts",
    "Main or Ancillary Hall?",
    "Calculated courts from Area",
    "Estimated courts",
    "Equivalent courts",
    "Equivalent court PP hours",
    "Equivalent courts for Comm Use",
    "Equivalent court for Comm Use PP hours",
    "Main Hall on Site? working out",
    "Main Hall on Site?"
]]

asp_times_df = asp_times_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Week Day",
    "Accessibility Type (Text)",
    "Start Time",
    "End Time",
    "Duration",
    "Weekly Duration (hh:mm)",
    "Mon 12:00-13:30",
    "Mon 16:00-22:00",
    "Tues 12:00-13:30",
    "Tues 16:00-22:00",
    "Wed 12:00-13:30",
    "Wed 16:00-22:00",
    "Thurs 12:00-13:30",
    "Thurs 16:00-22:00",
    "Fri 12:00-13:30",
    "Fri 16:00-22:00",
    "Sat 09:00-16:00",
    "Sun 09:00-16:30",
    "SP Peak Period Total",
    "SP PP Total for Comm Use"
]]

ash_times_df = ash_times_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Week Day",
    "Accessibility Type (Text)",
    "Start Time",
    "End Time",
    "Duration",
    "Weekly Duration (hh:mm)",
    "Mon 17:00-22:00",
    "Tues 17:00-22:00",
    "Wed 17:00-22:00",
    "Thurs 17:00-22:00",
    "Fri 17:00-22:00",
    "Sat 09:30-17:00",
    "Sun 09:00-14:30",
    "Sun 17:00-19:30",
    "SH Peak Period Total",
    "SH PP Total for Comm Use"
]]

artificialgrasspitches_df = artificialgrasspitches_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Pitches",
    "Width",
    "Length",
    "Area",
    "Artificial Sports Lighting",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

athletics_df = athletics_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Oval Track Lanes",
    "Artificial Sports Lighting",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

cycling_df = cycling_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Automatic Start Gate",
    "Bike Wash",
    "Black Trails",
    "Blue Trails",
    "Degree of Banking at Middle of Bends",
    "Degree of Banking at Middle of Straight",
    "Extreme Trails",
    "Finish Straight Length",
    "Finish Straight Width",
    "Artificial Sports Lighting",
    "Green Trails",
    "Length of Black Trails",
    "Length of Blue Trails",
    "Length of Extreme Trails",
    "Length of Green Trails",
    "Length of Red Trails",
    "Length of Straights",
    "No of Persons Start Gate",
    "Number of Straights",
    "Number of Turns",
    "Overall Width",
    "Radius of Turns/Bends",
    "Red Trails",
    "Start Gate",
    "Start Hill Elevation",
    "Start Hill Width",
    "Start Straight Length",
    "Start Straight Width",
    "Surface",
    "Timing System",
    "Total Length",
    "Width of Turns/Bends",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

golf_df = golf_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Holes",
    "Length",
    "Artificial Sports Lighting",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

grasspitches_df = grasspitches_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Pitches",
    "Artificial Sports Lighting",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

healthandfitnessgym_df = healthandfitnessgym_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Stations",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

icerinks_df = icerinks_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Width",
    "Length",
    "Area",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

indoorbowls_df = indoorbowls_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Rinks",
    "Width",
    "Length",
    "Area",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

indoortenniscentre_df = indoortenniscentre_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Courts",
    "Surface Type",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

outdoortenniscourts_df = outdoortenniscourts_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Courts",
    "Artificial Sports Lighting",
    "Surface Type",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name",
]]

skislopes_df = skislopes_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Skiable Width",
    "Skiable Length",
    "Artificial Sports Lighting",
    "Tow",
    "Slope Type",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

sportshalls_df = sportshalls_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Badminton Courts",
    "Width",
    "Length",
    "Area",
    "Dimensions Estimate",
    "Clearance exists - Ball / shuttlecock",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

squashcourts_df = squashcourts_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Courts",
    "Width",
    "Length",
    "Area",
    "Surface Type",
    "Movable Wall",
    "Doubles",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

studios_df = studios_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Width",
    "Length",
    "Area",
    "Dimensions Estimate",
    "Bike Stations",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

swimmingpools_df = swimmingpools_df[[
    "Site ID",
    "Site Name",
    "Facility ID",
    "Facility Subtype",
    "Accessibility Type (Text)",
    "Disability Access",
    "Disability Standard for individual sports",
    "Has changing rooms?",
    "Year Built",
    "Year Refurbished",
    "Last Updated Date",
    "Operational Status",
    "Lanes",
    "Width",
    "Length",
    "Area",
    "Minimum Depth",
    "Maximum Depth",
    "Diving Boards",
    "Movable Floor",
    "Easting",
    "Northing",
    "Urban/Rural",
    "Local Authority Name"
]]

schools_df = schools_df[[
    'URN',
    'EstablishmentNumber',
    'EstablishmentName',
    'TypeOfEstablishment (name)',
    'OpenDate',
    'PhaseOfEducation (name)',
    'StatutoryLowAge',
    'StatutoryHighAge',
    'Gender (name)',
    'NumberOfPupils',
    'NumberOfBoys',
    'NumberOfGirls',
    'Street',
    'Locality',
    'Address3',
    'Town',
    'Postcode',
    'SchoolWebsite',
    'TelephoneNum',
    'HeadTitle (name)',
    'HeadFirstName',
    'HeadLastName',
    'HeadPreferredJobTitle',
    'AdministrativeWard (name)',
    'Easting',
    'Northing'
]]

#----------------------------------
ap_db.drop_duplicates(subset=['Site ID', 'Site Name'], inplace=True)
artificialgrasspitches_df.drop_duplicates(subset=artificialgrasspitches_df.columns.difference(['Last Updated Date']), inplace=True)
athletics_df.drop_duplicates(subset=athletics_df.columns.difference(['Last Updated Date']), inplace=True)
cycling_df.drop_duplicates(subset=cycling_df.columns.difference(['Last Updated Date']), inplace=True)
golf_df.drop_duplicates(subset=golf_df.columns.difference(['Last Updated Date']), inplace=True)
grasspitches_df.drop_duplicates(subset=grasspitches_df.columns.difference(['Last Updated Date']), inplace=True)
healthandfitnessgym_df.drop_duplicates(subset=healthandfitnessgym_df.columns.difference(['Last Updated Date']), inplace=True)
icerinks_df.drop_duplicates(subset=icerinks_df.columns.difference(['Last Updated Date']), inplace=True)
indoorbowls_df.drop_duplicates(subset=indoorbowls_df.columns.difference(['Last Updated Date']), inplace=True)
indoortenniscentre_df.drop_duplicates(subset=indoortenniscentre_df.columns.difference(['Last Updated Date']), inplace=True)
outdoortenniscourts_df.drop_duplicates(subset=outdoortenniscourts_df.columns.difference(['Last Updated Date']), inplace=True)
skislopes_df.drop_duplicates(subset=skislopes_df.columns.difference(['Last Updated Date']), inplace=True)
sportshalls_df.drop_duplicates(subset=sportshalls_df.columns.difference(['Last Updated Date']), inplace=True)
squashcourts_df.drop_duplicates(subset=squashcourts_df.columns.difference(['Last Updated Date']), inplace=True)
studios_df.drop_duplicates(subset=studios_df.columns.difference(['Last Updated Date']), inplace=True)
swimmingpools_df.drop_duplicates(subset=swimmingpools_df.columns.difference(['Last Updated Date']), inplace=True)

#----------------------------------

# Load the existing Excel file using xlwings
excel_path = "activeplacescsvs/Blank Indoor Facilities database v15.9.xls"
app = xw.App(visible=False)  # Create a hidden Excel application
book = xw.Book(excel_path)   # Open the Excel file

# Select the sheets to write to
apdb_sheet = book.sheets['AP_Db']
swimmingpools_hours_sheet = book.sheets['SP-hours']
sportshalls_hours_sheet = book.sheets['SH-hours']
ash_times_sheet = book.sheets['A_SHtimes']
asp_times_sheet = book.sheets['A_SPtimes']
artificialgrasspitches_sheet = book.sheets['AGP']
athletics_sheet = book.sheets['AT']
cycling_sheet = book.sheets['CY']
golf_sheet = book.sheets['Golf']
grasspitches_sheet = book.sheets['GP']
healthandfitnessgym_sheet = book.sheets['H&F']
icerinks_sheet = book.sheets['Ice']
indoorbowlssheet = book.sheets['IB']
indoortenniscentre_sheet = book.sheets['IT']
outdoortenniscourts_sheet = book.sheets['TE']
skislopes_sheet = book.sheets['Ski']
sportshalls_sheet = book.sheets['SH']
squashcourts_sheet = book.sheets['SQ']
studios_sheet = book.sheets['St']
swimmingpools_sheet = book.sheets['SP']
schools_sheet = book.sheets['Schools']

# Write the APDB DataFrame to the Excel sheet starting at A3
apdb_sheet.range('A3').options(index=False, header=False).value = ap_db

# Write to the SP-hours, SH-hours, ash times, asp times and schools tabs starting at cell A2
swimmingpools_hours_sheet.range('A2').options(index=False, header=False).value = sp_hours_df
sportshalls_hours_sheet.range('A2').options(index=False, header=False).value = sh_hours_df
ash_times_sheet.range('A2').options(index=False, header=False).value = ash_times_df
asp_times_sheet.range('A2').options(index=False, header=False).value = asp_times_df
schools_sheet.range('A2').options(index=False, header=False).value = schools_df

# Write the other DataFrames to the sheet starting at cell A3
artificialgrasspitches_sheet.range('A3').options(index=False, header=False).value = artificialgrasspitches_df
athletics_sheet.range('A3').options(index=False, header=False).value = athletics_df
cycling_sheet.range('A3').options(index=False, header=False).value = cycling_df
golf_sheet.range('A3').options(index=False, header=False).value = golf_df
grasspitches_sheet.range('A3').options(index=False, header=False).value = grasspitches_df
healthandfitnessgym_sheet.range('A3').options(index=False, header=False).value = healthandfitnessgym_df
icerinks_sheet.range('A3').options(index=False, header=False).value = icerinks_df
indoorbowlssheet.range('A3').options(index=False, header=False).value = indoorbowls_df
indoortenniscentre_sheet.range('A3').options(index=False, header=False).value = indoortenniscentre_df
outdoortenniscourts_sheet.range('A3').options(index=False, header=False).value = outdoortenniscourts_df
skislopes_sheet.range('A3').options(index=False, header=False).value = skislopes_df
sportshalls_sheet.range('A3').options(index=False, header=False).value = sportshalls_df
squashcourts_sheet.range('A3').options(index=False, header=False).value = squashcourts_df
studios_sheet.range('A3').options(index=False, header=False).value = studios_df
swimmingpools_sheet.range('A3').options(index=False, header=False).value = swimmingpools_df

# Apply hh:mm format to the specified columns
def apply_time_format(sheet, columns):
    for col in columns:
        sheet.range(col + '1').expand('down').number_format = 'hh:mm'

# Apply formatting to SH-hours sheet
apply_time_format(sportshalls_hours_sheet, ['T', 'U', 'AC', 'AE'])

# Apply formatting to SP-hours sheet
apply_time_format(swimmingpools_hours_sheet, ['X', 'Y'])

# Save and close the workbook
output_xsl = f"{final_df['Local Authority Name'][0]} Indoor Facilities database v15.9.xls"

book.save(output_xsl)
book.close()
app.quit()

# # output cvs for debugging
# ap_db.to_csv("AP DB.csv", index=False)
# sp_hours_df.to_csv("SP hours.csv", index=False)
# sh_hours_df.to_csv("SH hours.csv", index=False)
# asp_times_df.to_csv("ASP times.csv", index=False)
# ash_times_df.to_csv("ASH times.csv", index=False)
# artificialgrasspitches_df.to_csv("artificialgrasspitches_output.csv", index=False)
# athletics_df.to_csv("athletics_output.csv", index=False)
# cycling_df.to_csv("cycling_output.csv", index=False)
# golf_df.to_csv("golf_output.csv", index=False)
# grasspitches_df.to_csv("grasspitches_output.csv", index=False)
# healthandfitnessgym_df.to_csv("healthandfitnessgym_output.csv", index=False)
# icerinks_df.to_csv("icerinks_output.csv", index=False)
# indoorbowls_df.to_csv("indoorbowls_output.csv", index=False)
# indoortenniscentre_df.to_csv("indoortenniscentre_output.csv", index=False)
# outdoortenniscourts_df.to_csv("outdoortenniscourts_output.csv", index=False)
# skislopes_df.to_csv("skislopes_output.csv", index=False)
# sportshalls_df.to_csv("sportshalls_output.csv", index=False)
# squashcourts_df.to_csv("squashcourts_output.csv", index=False)
# studios_df.to_csv("studios_output.csv", index=False)
# swimmingpools_df.to_csv("swimmingpools_output.csv", index=False)
#---------------------------------------------
# final_df.to_csv("final df.csv", index=False)
# schools_df.to_csv("schools.csv", index=False)
#sites_df.to_csv("sites df.csv", index=False)

print ("File saved")

