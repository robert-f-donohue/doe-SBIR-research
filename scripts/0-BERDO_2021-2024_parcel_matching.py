import pandas as pd

# relevant features for clustering 2024
features_2024 = [
    'Reported Gross Floor Area (Sq Ft)',
    'Site EUI (Energy Use Intensity kBtu/ft2)',
    'Total Site Energy Usage (kBtu)',
    'Estimated Total GHG Emissions (kgCO2e)'
]

# relevant features for clustering 2021
features_2021 = [
    'Tax Parcel ID',
    'Year Built'
]

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Loading Data -------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# File path to 2024 BERDO data
file_path_berdo_2024 = '../data-files/berdo_data_files/BERDO_Data-2024-clean.csv'
# File path to 2021 BERDO data
file_path_berdo_2021 = '../data-files/berdo_data_files/BERDO_Data-2021.csv'

# create DataFrames for 2021 and 2024 data
df_2024 = pd.read_csv(file_path_berdo_2024)
df_2021 = pd.read_csv(file_path_berdo_2021)

column_to_compare = 'Tax Parcel ID'

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Data Cleaning 2024 -------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# keep only multifamily
multifamily_df = df_2024[df_2024['Largest Property Type'] == 'Multifamily Housing'].dropna(subset=features_2024)

# convert columns to numeric 2024
for feature in features_2024:
    multifamily_df[feature] = pd.to_numeric(multifamily_df[feature], errors='coerce')

# drop any rows with GSF below 20,000 SF
min_gsf = 20000
multifamily_df = multifamily_df[multifamily_df['Reported Gross Floor Area (Sq Ft)'] >= min_gsf]

# drop any rows with EUI below 15 kBtu/sf
min_eui = 15.0
multifamily_df = multifamily_df[multifamily_df['Site EUI (Energy Use Intensity kBtu/ft2)'] >= min_eui]

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Filtering --- ------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# add Tax Parcel ID column to df_2021
df_2021[column_to_compare] = df_2021['Tax Parcel']

# filter 2021 data for columns that have an entry for year built
df_2021_filtered = df_2021[df_2021['Year Built'].notna() & (df_2021['Year Built'] != '')]
df_2021_filtered = df_2021_filtered[[column_to_compare, 'Year Built']]

# convert columns to numeric
for feature in features_2021:
    df_2021_filtered[feature] = pd.to_numeric(df_2021_filtered[feature], errors='coerce')

# find number of matching values
matching_values = df_2021[column_to_compare].isin(multifamily_df[column_to_compare])
# get number of matches
num_matches = matching_values.sum()
print(f"Number of matches: {num_matches}")

# find number of matching values with a data year entry
matching_values = df_2021_filtered[column_to_compare].isin(multifamily_df[column_to_compare])
# get number of matches
num_matches = matching_values.sum()
print(f"Number of matches with Year Built: {num_matches}")

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Merging ------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# perform left-join from 2021 data into 2024 data
df_2024_enhanced = pd.merge(df_2024, df_2021_filtered, on=column_to_compare, how='left')

print(f'Size of original DataFrame: {df_2024.shape[0]}')
print(f'Size of filtered DataFrame: {df_2024_enhanced.shape[0]}')

# # add to csv
# csv_location = '../data-files/berdo_data_files/BERDO_Data-2024-clean-merged.csv'
# # df_2024_enhanced.to_csv(csv_location, index=False)
#
# print(f"Merge successful, file saved to: {csv_location}!")