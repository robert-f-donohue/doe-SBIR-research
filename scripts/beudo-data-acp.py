import pandas as pd
import openpyxl
from openpyxl import load_workbook


def calculate_fines(beudo_df, emissions_factors_df, start_year, end_year):
    fine_rate = 234  # $234 per metric ton of CO2e

    # Filter the dataset to only include rows that are in scope
    # For example, filter based on the 'BEUDO Property Type' or floor area, or any other relevant field
    beudo_df_filtered = beudo_df.dropna(subset=['Property GFA - Self Reported (ft2)'])

    # If there are any other conditions for rows being 'in scope', apply them here
    # For example, filtering by 'Data Year', or 'BEUDO Property Type'
    # Assuming there's a BEUDO Property Type column we can filter by (add other conditions as necessary)
    beudo_df_filtered = beudo_df_filtered[beudo_df_filtered['Primary Property Type - Self Selected'].notnull()]

    # Create an empty list to store the results (we will concatenate this list into a DataFrame later)
    fines_list = []

    for building_id in beudo_df_filtered['Reporting ID'].unique():
        # Subset the building's data
        building_data = beudo_df_filtered[beudo_df_filtered['Reporting ID'] == building_id]

        # Initialize the fine amount
        fine_amount = 0
        # building_floor_area = building_data['Property GFA - Self Reported (ft2)'].values[0]  # Floor area for CEI calculation

        for year in range(start_year, end_year + 1):
            # Get the threshold for the year
            emissions_threshold = building_data[f'Threshold {year}'].values[0]

            # Initialize the total emissions for the building in this year
            total_emissions = 0

            # Get the emissions factors for the current year
            emissions_factors_year = emissions_factors_df[emissions_factors_df['Data Year'] == year]

            # Calculate emissions for each fuel type using the emissions factors and constant fuel usage
            for fuel_type in ['Natural Gas', 'Net Electricity', 'District Hot Water', 'District Chilled Water',
                              'District Steam', 'Fuel Oil 1', 'Fuel Oil 2', 'Fuel Oil 4', 'Fuel Oil 5 and 6', 'Propane',
                              'Diesel Usage', 'Kerosene Usage']:
                if f'{fuel_type} Usage (kBtu)' in building_data.columns:
                    # Get constant fuel usage for the building
                    fuel_usage = building_data[f'{fuel_type} Usage (kBtu)'].values[0] / 1000  # kBtu to MmBtu

                    # Get the emissions factor for the fuel type for this year
                    emissions_factor = emissions_factors_year[f'{fuel_type} Emissions'].values[0]

                    # Calculate the emissions for the fuel type and add it to total emissions
                    fuel_emissions = fuel_usage * emissions_factor / 1000
                    total_emissions += fuel_emissions

            # # Make sure floor area is valid
            # if pd.isna(building_floor_area) or building_floor_area == 0:
            #     continue  # Skip this building if the floor area is invalid

            # # Calculate the CEI for the building for the current year
            # cei = total_emissions * 1000 / building_floor_area

            # If CEI exceeds the threshold, calculate the fine
            if total_emissions > emissions_threshold:
                excess_emissions = total_emissions - emissions_threshold
                fine_amount += excess_emissions * fine_rate
                # print(f"Excess Emissions: {excess_emissions}, Fine for {year}: {fine_amount}")

        fine_amount = round(fine_amount)

            # Add the fine amount to the total fines DataFrame
        fines_list.append({
            'Reporting ID': building_id,
            'Primary Property Type - Self Selected': building_data['Primary Property Type - Self Selected'].values[0],
            'Total Fines': fine_amount
        })

    total_fines = pd.DataFrame(fines_list)

    return total_fines


def aggregate_fines_by_typology(fines_df):
    # Group by building type and sum the fines
    aggregated_fines = fines_df.groupby('Primary Property Type - Self Selected').agg({'Total Fines': 'sum'}).reset_index()
    return aggregated_fines


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Generate DataFrames ------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# File path to BEUDO data & read in CSV
file_path_beudo = '../data-files/beudo_data_files/BEUDO_Data-Thresholds.csv'
df_beudo = pd.read_csv(file_path_beudo)
df_beudo_filtered = df_beudo[(df_beudo['Data Year'] == 2021) &
                             (~df_beudo['Property GFA - Self Reported (ft2)'].isnull()) &
                             (df_beudo['BEUDO Category'] != 'Residential')].copy()

# File path to GHG Emissions factors through 2050 and load in CSV
file_path_emissions = '../data-files/beudo_data_files/beudo-emissions-factors.csv'
df_emissions_factors = pd.read_csv(file_path_emissions)
# Emissions Factors for 2021
df_emissions_factors_2021 = df_emissions_factors[df_emissions_factors['Data Year'] == 2021]

# File path to build summary Excel sheet through 2050 and load in Excel as DF
file_path_excel = '../data-files/beudo_data_files/beudo_building_summary_statistics-acp.xlsx'
xls = pd.ExcelFile(file_path_excel)

# Load the "Building Type Summary" sheet into a DataFrame
df_building_type_summary = pd.read_excel(xls, sheet_name='Building Type Summary')

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Calculate BEUDO Thresholds ----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Add columns for emissions by fuel source for each Data Year/Reporting ID
df_beudo_filtered['Net Electricity Emissions'] = ((df_beudo_filtered['Net Electricity Usage (kBtu)']
                                                             / 1000) * df_emissions_factors_2021['Net Electricity Emissions']) / 1000

df_beudo_filtered['Natural Gas Emissions'] = ((df_beudo_filtered['Natural Gas Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['Natural Gas Emissions']) / 1000

df_beudo_filtered['Fuel Oil 1 Emissions'] = ((df_beudo_filtered['Fuel Oil 1 Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['Fuel Oil 1 Emissions']) / 1000

df_beudo_filtered['Fuel Oil 2 Emissions'] = ((df_beudo_filtered['Fuel Oil 2 Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['Fuel Oil 2 Emissions']) / 1000

df_beudo_filtered['Fuel Oil 4 Emissions'] = ((df_beudo_filtered['Fuel Oil 4 Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['Fuel Oil 4 Emissions']) / 1000

df_beudo_filtered['Fuel Oil 5 and 6 Emissions'] = ((df_beudo_filtered['Fuel Oil 5 and 6 Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['Fuel Oil 5 and 6 Emissions']) / 1000

df_beudo_filtered['Diesel 2 Emissions'] = ((df_beudo_filtered['Diesel 2 Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['Diesel 2 Emissions']) / 1000

df_beudo_filtered['Kerosene Emissions'] = ((df_beudo_filtered['Kerosene Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['Kerosene Emissions']) / 1000

df_beudo_filtered['District Chilled Water Emissions'] = ((df_beudo_filtered['District Chilled Water Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['District Chilled Water Emissions']) / 1000

df_beudo_filtered['District Steam Emissions'] = ((df_beudo_filtered['District Steam Usage (kBtu)'] / 1000)
                                                  * df_emissions_factors_2021['District Steam Emissions']) / 1000
# Fill in all NaN's with 0's
df_beudo_filtered.fillna(0, inplace=True)

# Create a list of all GHG emissions columns to feed into the next operation
GHG_emissions_columns = ['Net Electricity Emissions', 'Natural Gas Emissions',
                         'Fuel Oil 1 Emissions', 'Fuel Oil 2 Emissions',
                         'Fuel Oil 4 Emissions', 'Fuel Oil 5 and 6 Emissions',
                         'Diesel 2 Emissions', 'Kerosene Emissions',
                         'District Chilled Water Emissions', 'District Steam Emissions']

# Sum all emissions data rows to get the total GHG emissions based on established emissions factors
# Important to note that this value is NOT the same as the reported value & should be checked with Samira
df_beudo_filtered['Baseline GHG Emissions'] = df_beudo_filtered[GHG_emissions_columns].sum(axis=1)

# Split the merged data by building GFA (+/- 100,000 SF)
beudo_n_l = df_beudo_filtered[df_beudo_filtered['Property GFA - Self Reported (ft2)'] >= 100000].copy()
beudo_n_s = df_beudo_filtered[df_beudo_filtered['Property GFA - Self Reported (ft2)'] < 100000].copy()

# Define the compliance thresholds for large buildings and their corresponding multipliers
compliance_thresholds_large = {
    (2025, 2029): 0.8,
    (2030, 2034): 0.4,
    (2035, 2050): 0.0
}

# Populate new columns for large building BEUDO Compliance Thresholds (>100,000 sf)
# Iterate over each range and multiplier
for year_range, multiplier in compliance_thresholds_large.items():
    start_year, end_year = year_range  # Unpack the start and end years
    for year in range(start_year, end_year + 1):  # Include the end year
        # Create a new column for each year within the range
        column_name = f'Threshold {year}'
        beudo_n_l[column_name] = beudo_n_l['Baseline GHG Emissions'] * multiplier


# Define the compliance thresholds for small buildings and their corresponding multipliers
compliance_thresholds_small = {
    (2025, 2029): 1.0,
    (2030, 2034): 0.6,
    (2035, 2039): 0.4,
    (2040, 2044): 0.2,
    (2045, 2049): 0.1,
    (2050, 2050): 0.0
}

# Populate new columns for small building BEUDO Compliance Thresholds (25,000 sf - 99,999 sf)
# Iterate over each range and multiplier
for year_range, multiplier in compliance_thresholds_small.items():
    start_year, end_year = year_range  # Unpack the start and end years
    for year in range(start_year, end_year + 1):  # Include the end year
        # Create a new column for each year within the range
        column_name = f'Threshold {year}'
        beudo_n_s[column_name] = beudo_n_s['Baseline GHG Emissions'] * multiplier

df_beudo_combined = pd.concat([beudo_n_s, beudo_n_l], ignore_index=True)
print(df_beudo_combined.shape)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Calculate & Aggregate Fines ----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Calculate fines for 2025-2030 and 2030-2050 using the optimized function
fines_2025_2030 = calculate_fines(df_beudo_combined, df_emissions_factors, 2025, 2030)
fines_2030_2050 = calculate_fines(df_beudo_combined, df_emissions_factors, 2030, 2050)

# Aggregate fines by building type
aggregated_fines_2025_2030 = aggregate_fines_by_typology(fines_2025_2030)
aggregated_fines_2030_2050 = aggregate_fines_by_typology(fines_2030_2050)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Add columns to Building Summary Table ------------------------------
# --------------------------------------------------------------------------------------------------------

# Merge with building summary data and save to Excel
# First merge aggregated fines for 2025-2030
building_type_summary_df = pd.merge(df_building_type_summary, aggregated_fines_2025_2030,
                                    on='Primary Property Type - Self Selected', how='left', suffixes=('', '_2025_2030'))

# Now merge aggregated fines for 2030-2050
building_type_summary_df = pd.merge(building_type_summary_df, aggregated_fines_2030_2050,
                                    on='Primary Property Type - Self Selected', how='left', suffixes=('', '_2030_2050'))

# Save the updated building summary sheet to the Excel file
with pd.ExcelWriter(file_path_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    building_type_summary_df.to_excel(writer, sheet_name='Building Type Summary', index=False)