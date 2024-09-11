import pandas as pd
import openpyxl
from openpyxl import load_workbook


def calculate_fines(source_data_df, emissions_factors_df, start_year, end_year):
    fine_rate = 268  # $268 per metric ton of CO2e

    # Create an empty list to store the results (we will concatenate this list into a DataFrame later)
    fines_list = []

    for building_id in source_data_df['Property Id'].unique():
        # Subset the building data
        building_data = source_data_df[source_data_df['Property Id'] == building_id]

        # Initialize the fine amount
        fine_amount = 0
        building_floor_area = building_data['Gross Floor Area (ft2)'].values[0]  # Floor area for CEI calculation

        for year in range(start_year, end_year + 1):
            # Get the CEI threshold for the year
            cei_threshold = building_data[f'Threshold {year}'].values[0]

            # Initialize the total emissions for the building in this year
            total_emissions = 0

            # Get the emissions factors for the current year
            emissions_factors_year = emissions_factors_df[emissions_factors_df['Data Year'] == year]


            # Calculate emissions for each fuel type using the emissions factors and constant fuel usage
            for fuel_type in ['Natural Gas', 'Electricity', 'District Hot Water', 'District Chilled Water',
                              'District Steam', 'Fuel Oil 1', 'Fuel Oil 2', 'Fuel Oil 4', 'Fuel Oil 5 and 6', 'Propane',
                              'Diesel 2', 'Kerosene']:
                if f'{fuel_type} Usage (kBtu)' in building_data.columns:
                    # Get constant fuel usage for the building
                    fuel_usage = building_data[f'{fuel_type} Usage (kBtu)'].values[0] / 1000  # kBtu to MmBtu

                    # Get the emissions factor for the fuel type for this year
                    emissions_factor = emissions_factors_year[f'{fuel_type} Emissions'].values[0]

                    # Calculate the emissions for the fuel type and add it to total emissions
                    fuel_emissions = fuel_usage * emissions_factor
                    total_emissions += fuel_emissions

            # Make sure floor area is valid
            if pd.isna(building_floor_area) or building_floor_area == 0:
                continue  # Skip this building if the floor area is invalid

            # Calculate the CEI for the building for the current year
            cei = total_emissions / building_floor_area

            # If CEI exceeds the threshold, calculate the fine
            if cei > cei_threshold:
                excess_emissions = ((cei - cei_threshold) * building_floor_area) / 1000  # Convert kg CO2e to MT CO2e
                fine_amount += excess_emissions * fine_rate
                # print(f"Excess Emissions: {excess_emissions}, Fine for {year}: {fine_amount}")

        fine_amount = round(fine_amount)

            # Add the fine amount to the total fines DataFrame
        fines_list.append({
            'BERDO ID': building_id,
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

# File path to BERDO data & read in CSV
file_path_berdo = '../data-files/LL_84_data_files/LL84_Data-Thresholds.csv'
df_ll84 = pd.read_csv(file_path_berdo)

# File path to GHG Emissions factors through 2050 and load in CSV
file_path_emissions = '../data-files/LL_84_data_files/LL84-emissions-factors.csv'
df_emissions_factors = pd.read_csv(file_path_emissions)

# File path to LL84 thresholds
file_path_thresholds = '../data-files/LL_84_data_files/LL84-thresholds.csv'
df_thresholds = pd.read_csv(file_path_thresholds)
df_thresholds['Year'] = df_thresholds['Year'].astype(int)


# File path to build summary Excel sheet through 2050 and load in Excel as DF
file_path_excel = '../data-files/LL_84_data_files/LL84_building_summary_statistics-acp.xlsx'
xls = pd.ExcelFile(file_path_excel)

# Load the "Building Type Summary" sheet into a DataFrame
df_building_type_summary = pd.read_excel(xls, sheet_name='Building Type Summary')


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Calculate LL97 Thresholds ------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Subset df to only include relevant columns
columns_to_keep = [
    'Property Id', 'Primary Property Type - Portfolio Manager-Calculated',
    'Primary Property Type - Self Selected', 'Gross Floor Area (ft2)', 'Site EUI (kBtu/sf)', 'Fuel Oil 1 Usage (kBtu)',
    'Fuel Oil 2 Usage (kBtu)', 'Fuel Oil 4 Usage (kBtu)', 'Fuel Oil 5 and 6 Usage (kBtu)', 'Diesel 2 Usage (kBtu)',
    'Propane Usage (kBtu)', 'District Steam Usage (kBtu)', 'District Hot Water Usage (kBtu)',
    'District Chilled Water Usage (kBtu)', 'Natural Gas Usage (kBtu)', 'Electricity Usage (kBtu)',
    'Property GFA - Calculated (Buildings) (ft2)', 'Property GFA - Calculated (Buildings and Parking) (ft2)',
]

# Drop building types exempt from compliance
building_types_to_drop = ['Worship Facility', 'Police Station', 'Prison/Incarceration', 'Courthouse',
                          'Energy/Power Station', 'Zoo', 'Mailing Center/Post Office', 'Other']

# Energy columns to ensure they remain floats
columns_energy = [
    'Fuel Oil 1 Usage (kBtu)', 'Fuel Oil 2 Usage (kBtu)', 'Fuel Oil 4 Usage (kBtu)', 'Fuel Oil 5 and 6 Usage (kBtu)',
    'Diesel 2 Usage (kBtu)', 'Propane Usage (kBtu)', 'District Steam Usage (kBtu)', 'District Hot Water Usage (kBtu)',
    'District Chilled Water Usage (kBtu)', 'Natural Gas Usage (kBtu)', 'Electricity Usage (kBtu)'
]

# Filter the dataset to only include rows that are in scope
ll84_df_filtered = df_ll84[columns_to_keep].copy()
ll84_df_filtered = ll84_df_filtered.dropna(subset=['Gross Floor Area (ft2)'])
ll84_df_filtered = ll84_df_filtered[ll84_df_filtered['Primary Property Type - Self Selected'].notnull()]
ll84_df_filtered = ll84_df_filtered[ll84_df_filtered['Site EUI (kBtu/sf)'].notnull()]
ll84_df_filtered = (ll84_df_filtered[~ll84_df_filtered['Primary Property Type - Self Selected']
                    .isin(building_types_to_drop)])
ll84_df_filtered['Natural Gas Usage (kBtu)'] = (ll84_df_filtered['Natural Gas Usage (kBtu)']
                                                .replace('Insufficient access', 0))
ll84_df_filtered['Natural Gas Usage (kBtu)'] = ll84_df_filtered['Natural Gas Usage (kBtu)'].astype(float)
ll84_df_filtered['Electricity Usage (kBtu)'] = (ll84_df_filtered['Electricity Usage (kBtu)']
                                                .replace('Insufficient access', 0))
ll84_df_filtered['Electricity Usage (kBtu)'] = ll84_df_filtered['Electricity Usage (kBtu)'].astype(float)


# Fill all NaN with 0's
ll84_df_filtered.fillna(0, inplace=True)

# Melt df_property_thresholds
df_thresholds_melted = df_thresholds.melt(
    id_vars=['Year'], var_name='Primary Property Type - Self Selected', value_name='Threshold')

# Merge on 'Largest Property Type'
df_merged = ll84_df_filtered.merge(df_thresholds_melted, on='Primary Property Type - Self Selected', how='left')

# Pivot the DataFrame to have years as columns
df_pivot = df_merged.pivot_table(index=columns_to_keep, columns='Year', values='Threshold').reset_index()

# Add the Threshold columns and convert to int
df_pivot.columns = (columns_to_keep + [f'Threshold {int(year)}' for year in df_pivot.columns[18:]])

print(df_pivot.dtypes)
df_pivot.to_csv('../data-files/LL_84_data_files/test.csv', index=False)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Calculate & Aggregate Fines ----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Calculate fines for 2025-2030 and 2030-2050 using the optimized function
fines_2025_2030 = calculate_fines(df_pivot, df_emissions_factors, 2025, 2030)
fines_2030_2050 = calculate_fines(df_pivot, df_emissions_factors, 2030, 2050)

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
