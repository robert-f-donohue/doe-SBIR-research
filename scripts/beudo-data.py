import pandas as pd

file_path = '../data-files/BEUDO_Data.csv'

df = pd.read_csv(file_path)

df_2022 = df[(df['Data Year'] == 2022) & (~df['Property GFA - Self Reported (ft2)'].isnull())]

df_filtered = df[(df['Data Year'] != 2022) & (~df['Property GFA - Self Reported (ft2)'].isnull())]

df_filtered = df_filtered[~df_filtered.index.isin(df_2022.index)]

df_filtered = df_filtered.sort_values(by='Data Year', ascending=True)

df_filtered_2018 = df_filtered[(df_filtered['Data Year'] == 2018)]
df_filtered_2019 = df_filtered[(df_filtered['Data Year'] == 2019)]
df_filtered_2020 = df_filtered[(df_filtered['Data Year'] == 2020)]
df_filtered_2021 = df_filtered[(df_filtered['Data Year'] == 2021)]
df_filtered_2022 = df_filtered[(df_filtered['Data Year'] == 2022)]

print(df.shape)
# print(df_2022.shape)
# print(df_filtered.shape)
# print(df_filtered_2018.shape)
# print(df_filtered_2019.shape)
# print(df_filtered_2020.shape)
print(df_filtered_2021.shape)
# print(df_filtered_2022.shape)

df_2021_non_res = df_filtered_2021[(df_filtered_2021['BEUDO Category'] == 'Non-Residential')]
df_2021_non_res_large = df_2021_non_res[df_2021_non_res["Property GFA - Self Reported (ft2)"] >= 100000]
df_2021_non_res_small = df_2021_non_res[df_2021_non_res["Property GFA - Self Reported (ft2)"] < 100000]

df_2021_municipal = df_filtered_2021[(df_filtered_2021['BEUDO Category'] == 'Municipal')]
df_2021_res = df_filtered_2021[(df_filtered_2021['BEUDO Category'] == 'Residential')]

print(f'\nThe total amount of buildings is {df_filtered_2021.shape[0]}')
print(f'The total GFA of BEUDO reported buildings is {df_filtered_2021["Property GFA - Self Reported (ft2)"].sum()
                                                      .astype(int)} SF\n')

print("------------------------------------------------------------------------")
print("------------------------------------------------------------------------")

print(f'The total amount of Non-Residential buildings is {df_2021_non_res.shape[0]}')
print(f'The total GFA of Non-Residential buildings is {df_2021_non_res["Property GFA - Self Reported (ft2)"].sum()
                                                        .astype(int)} SF\n')

print("------------------------------------------------------------------------")

print(f'The total number of Non-Residential buildings with GFA >100,000 is {df_2021_non_res_large.shape[0]}')
print(f'The total GFA of Non-Residential buildings with GFA >100,000 is {df_2021_non_res_large[
    "Property GFA - Self Reported (ft2)"].sum().astype(int)} SF\n')

print("------------------------------------------------------------------------")

print(f'The total number of Non-Residential buildings with GFA <100,000 is {df_2021_non_res_small.shape[0]}')
print(f'The total GFA of Non-Residential buildings with GFA <100,000 is {df_2021_non_res_small[
    "Property GFA - Self Reported (ft2)"].sum().astype(int)} SF\n')

print("------------------------------------------------------------------------")
print("------------------------------------------------------------------------")

print(f'The total amount of Municipal buildings is {df_2021_municipal.shape[0]}')
print(f'The total GFA of Municipal buildings is {df_2021_municipal["Property GFA - Self Reported (ft2)"].sum()
                                                    .astype(int)} SF\n')

print("------------------------------------------------------------------------")
print("------------------------------------------------------------------------")

print(f'The total amount of Residential buildings is {df_2021_res.shape[0]}')
print(f'The total GFA of Residential buildings is {df_2021_res["Property GFA - Self Reported (ft2)"].sum()
                                                    .astype(int)} SF')
