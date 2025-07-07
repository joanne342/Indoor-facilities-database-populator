How to use the program: https://docs.google.com/document/d/185wcDMdNERGC0vZpbz9RKsZKTDzBjMLv/

What the program does:
https://docs.google.com/document/d/1QdcVxrdVOV7FwuvNj9Jaxe5H7hZD7YIn/

Apprenticeship portfolio report section about this program before it was removed for hitting the same KSBs as the Active Places program section:
https://docs.google.com/document/d/1QZfuGdcVQOoWt3Xu_5kWc_ybxshUnUdu/

# Indoor-facilities-database-populator
Indoor facilities database populator

The program is designed to load, clean, enrich, and combine multiple datasets about sports and leisure facilities within a specified local authority. It:

- Validates user input.<br>
- Filters relevant data.<br>
- Merges multiple data sources.<br>
- Replaces codes with human-readable values.<br>
- Calculates timing durations and accessibility standards.<br>
- And prepares the data for further analysis or reporting.<br>
- Validates and calculates facility usage during key periods.<br>
- Computes capacities based on physical and scheduling data.<br>
- Prepares specialized summaries and facility-specific datasets.<br>
- Outputs clean, ready-to-analyze tables for reporting or further use in Excel.

------------------------------------

What the Program Does

1. Imports necessary libraries:<br>
	- pandas (data manipulation)<br>
	- xlwings (Excel interaction)<br>
	- math<br>
	- glob (file pattern matching)

2. Time conversion utilities:<br>
	- Functions to convert between hh:mm format and minutes.

3. User input:<br>
	- Gets a local authority code from the user.

4. Loads multiple CSV files related to various sports facilities, including:<br>
	- Artificial grass pitches, athletics tracks, cycling, golf, grass pitches, gyms, ice rinks, indoor bowls, tennis centers, sports halls, squash courts, studios, swimming pools, ski slopes, etc.

5. Selects relevant columns from each sports CSV:<br>
	- Focuses on accessibility, facility details, disability access, lighting, trails, pitches, courts, timings, surface types, and many other descriptive attributes.

6. Loads a sites.csv file containing general site info:<br>
	- Including location, authority codes, contact info, closure reasons, etc.

7. Validates the local authority code entered by the user against the sites data.

8. Filters data to keep only rows matching the valid local authority code.

9. Merges sports data and sites data into a combined dataframe.

10. Loads facility timings from facilitytimings.csv and merges it with the combined data.

11. Filters out closed facilities from the dataset.

12. Loads education base data from a CSV with a dynamic date in the filename, picking the latest file if multiple exist.

13. Filters out irrelevant local areas and closed facilities from this dataset.

14. Replaces binary indicators (0/1) with "Yes"/"No" or blank for various facility and disability-related columns to improve readability.

15. Maps coded values to descriptive text for columns like:<br>
	- Facility type and subtype<br>
	- Operational status<br>
	- Surface type<br>
	- Week day<br>
	- Slope type

16. Processes disability access standards by checking which disability features are marked "Yes" or blank and concatenates appropriate indicators.

17. Checks for presence of each facility type per site, e.g., if a site has an athletic track, indoor bowls, gym, etc., marking "Yes" or blank accordingly.

18. Counts the number of activity areas per site based on these facility presences.

19. Concatenates address fields into a single address string.

20. Concatenates open-close timings into one string per facility.

21. Formats certain columns:<br>
	- Adds spaces to car park capacity for readability.<br>
	- Trims timestamps to dates for the last checked column.

22. Calculates duration of facility availability by converting start and end times to timedeltas and subtracting them.

23. Adds columns representing specific time blocks for days and times (e.g., Mon 17:00-22:00), extracting date and time info from column names.

24. Day and Time Validation & Overlap Calculation<br>
	- Checks if days are valid based on predefined acceptable values.<br>
	- Converts start/end times into timedelta objects.<br>
	- Calculates overlap between given time intervals by comparing start and end times.

25. Calculates Peak Period Totals for Facilities<br>
	- Defines specific peak periods for Sports Halls (SH) and Swimming Pools (SP) based on day/time ranges.<br>
	- Sums relevant times into SH Peak Period Total and SP Peak Period Total.<br>
	- Calculates "Peak Period Total for Community Use" only for certain accessibility types.

26. Capacity & Usage Calculations<br>
	- Calculates "Person minutes in pool" depending on pool type.<br>
	- Computes "Pool capacity (visits)" using area, peak period totals, and person minutes.<br>
	- Flags pools larger than 100 m².<br>
	- Derives court counts based on area using a lookup table with floor rounding.<br>
	- Determines if halls are "Main" or "Ancillary" based on clearance and court count.

27. Court Estimates and Equivalents<br>
	- Estimates courts differently depending on hall type.<br>
	- Calculates equivalent courts by comparing court counts and estimates.<br>
	- Computes equivalent court peak period hours.<br>
	- Similarly calculates equivalent courts and peak period hours for community use only.

28. Site-wide Indicators<br>
	- Flags if a site has a large pool or main hall by concatenating site ID and facility info.<br>
	- Checks if these flags appear anywhere in the dataset for the given site.

29. Duration and Sorting<br>
	- Calculates weekly duration by multiplying session durations by the number of applicable days.<br>
	- Sorts the final dataframe by site, facility, and days of the week in a custom order.

30. Data Cleaning and Indexing<br>
	- Removes duplicates based on various columns.<br>
	- Sorts by geographic coordinates (Northing, Easting) and assigns sequential reference numbers (KKP Ref).

31. Facility-specific Filtering<br>
	- Creates filtered dataframes for each facility type (sports halls, swimming pools, athletics, golf, etc.).<br>
	- Adds specific columns to each dataframe depending on the facility type.

32. Output Preparation<br>
	- Drops duplicates in final datasets.<br>
	- Uses xlwings (Python Excel library) to write all these processed dataframes to an Excel workbook in designated sheets and starting cells.

