import pandas as pd
import datetime
import matplotlib.pyplot as plt
import seaborn as sns

columns_to_keep = [
    '_id', 'worktype', 'declared_valuation', 'total_fees', 'issued_date', 'occupancytype',
    'sq_feet', 'property_id', 'parcel_id'
]

columns_to_keep_worktype = [
    'ADDITION', 'NEWCON',
]

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Loading Data -------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# File path to renovation schedule
file_path_renovations = '../data-files/berdo_data_files/renovations-since-2009.csv'

# create DataFrame
df = pd.read_csv(file_path_renovations)

# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Cleaning Data ------------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Remove '$' from anywhere in the string
df['total_fees'] = df['total_fees'].str.replace(r'\$', '', regex=True)
df['declared_valuation'] = df['declared_valuation'].str.replace(r'\$', '', regex=True)

# make them numeric
df['total_fees'] = pd.to_numeric(df['total_fees'], errors='coerce')
df['declared_valuation'] = pd.to_numeric(df['declared_valuation'], errors='coerce')
df['issued_date'] = pd.to_datetime(df['issued_date'], errors='coerce').dt.date

# keep relevant columns
df_relevant = df[columns_to_keep]
df_occupancy = df_relevant[df_relevant['occupancytype'].isin(['MIXED', '7unit', 'Other'])]
df_area = df_occupancy[df_occupancy['sq_feet'] >= 2000]

# check shapes
print(df.shape)
print(df_occupancy.shape)
print(df_area.shape)
# check occupancy types
print(df_occupancy.dtypes)
print(df_occupancy['worktype'].unique())

df_area.to_csv('../data-files/berdo_data_files/renovations-since-2009-smaller.csv', index=False)


# --------------------------------------------------------------------------------------------------------
# ----------------------------------- Visualize Distributions --------------------------------------------
# --------------------------------------------------------------------------------------------------------

# # Distributions
# plt.figure(figsize=(20, 10))
#
# # Occupancy distributions
# plt.subplot(1, 2, 1)
# sns.histplot(df_occupancy['occupancytype'], bins=30, kde=True)
# plt.title('Occupancy Type Distribution')
# plt.xlabel('Occupancy Type')
# plt.ylabel('Count')
#
# # Work Type distributions
# plt.subplot(1, 2, 2)
# sns.histplot(df_occupancy['worktype'], bins=30, kde=True)
# plt.title('Work Type Distribution')
# plt.xlabel('Work Type')
# plt.ylabel('Count')
#
# plt.tight_layout()
# plt.show()