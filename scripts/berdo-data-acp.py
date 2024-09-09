import pandas as pd
import openpyxl
from openpyxl import load_workbook


def calculate_fines(berdo_df, emissions_factors_df, start_year, end_year):
    fine_rate = 234  # $234 per metric ton of CO2e

    # Filter the dataset to only include rows that are in scope
    # For example, filter based on the 'BERDO Property Type' or floor area, or any other relevant field
    berdo_df_filtered = berdo_df.dropna(subset=['Reported Gross Floor Area (Sq Ft)'])

    # If there are any other conditions for rows being 'in scope', apply them here
    # For example, filtering by 'Data Year', or 'BERDO Property Type'
    # Assuming there's a BERDO Property Type column we can filter by (add other conditions as necessary)
    berdo_df_filtered = berdo_df_filtered[berdo_df_filtered['BERDO Property Type'].notnull()]

    # Create an empty list to store the results (we will concatenate this list into a DataFrame later)
    fines_list = []

    for building_id in berdo_df_filtered['BERDO ID'].unique():
        # Subset the building's data
        building_data = berdo_df_filtered[berdo_df_filtered['BERDO ID'] == building_id]

        # Initialize the fine amount
        fine_amount = 0
        building_floor_area = building_data['Reported Gross Floor Area (Sq Ft)'].values[0]  # Floor area for CEI calculation

        for year in range(start_year, end_year + 1):
            # Get the CEI threshold for the year
            cei_threshold = building_data[f'Threshold {year}'].values[0]

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
            'BERDO Property Type': building_data['BERDO Property Type'].values[0],
            'Total Fines': fine_amount
        })

    total_fines = pd.DataFrame(fines_list)

    return total_fines


def aggregate_fines_by_typology(fines_df):
    # Group by building type and sum the fines
    aggregated_fines = fines_df.groupby('BERDO Property Type').agg({'Total Fines': 'sum'}).reset_index()
    return aggregated_fines


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Generate DataFrames ------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# File path to BERDO data & read in CSV
file_path_berdo = '../data-files/berdo_data_files/BERDO_Data-Thresholds.csv'
df_berdo = pd.read_csv(file_path_berdo)

# File path to GHG Emissions factors through 2050 and load in CSV
file_path_emissions = '../data-files/berdo_data_files/berdo-emissions-factors.csv'
df_emissions_factors = pd.read_csv(file_path_emissions)

# File path to build summary Excel sheet through 2050 and load in Excel as DF
file_path_excel = '../data-files/berdo_data_files/berdo_building_summary_statistics-acp.xlsx'
xls = pd.ExcelFile(file_path_excel)

# Load the "Building Type Summary" sheet into a DataFrame
df_building_type_summary = pd.read_excel(xls, sheet_name='Building Type Summary')

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Calculate & Aggregate Fines ----------------------------------------
# --------------------------------------------------------------------------------------------------------

# Calculate fines for 2025-2030 and 2030-2050 using the optimized function
fines_2025_2030 = calculate_fines(df_berdo, df_emissions_factors, 2025, 2030)
fines_2030_2050 = calculate_fines(df_berdo, df_emissions_factors, 2030, 2050)

# Aggregate fines by building type
aggregated_fines_2025_2030 = aggregate_fines_by_typology(fines_2025_2030)
aggregated_fines_2030_2050 = aggregate_fines_by_typology(fines_2030_2050)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Add columns to Building Summary Table ------------------------------
# --------------------------------------------------------------------------------------------------------

# Merge with building summary data and save to Excel
# First merge aggregated fines for 2025-2030
building_type_summary_df = pd.merge(df_building_type_summary, aggregated_fines_2025_2030,
                                    on='BERDO Property Type', how='left', suffixes=('', '_2025_2030'))

# Now merge aggregated fines for 2030-2050
building_type_summary_df = pd.merge(building_type_summary_df, aggregated_fines_2030_2050,
                                    on='BERDO Property Type', how='left', suffixes=('', '_2030_2050'))

# Save the updated building summary sheet to the Excel file
with pd.ExcelWriter(file_path_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    building_type_summary_df.to_excel(writer, sheet_name='Building Type Summary', index=False)
