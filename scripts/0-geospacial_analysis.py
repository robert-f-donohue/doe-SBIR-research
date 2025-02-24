import numpy as np
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import folium
from folium.plugins import MarkerCluster


# function to geocode an address
def geocode_address(address):
    try:
        location = geocode(address)
        if location:
            return pd.Series(location.latitude, location.longitude)
        else:
            return pd.Series([None, None])
    except:
        return pd.Series([None, None])

# Function to get color for a given cluster
def get_cluster_color(cluster_label):
    return cluster_colors.get(cluster_label, 'gray')  # Default to gray if cluster not found


# load in clustered data
data_file = '../data-files/berdo_data_files/exploratory-analysis-results/cross_referenced_clustered_data.csv'
df = pd.read_csv(data_file)

# --------------------------------------------------------------------------------------------------------
# ---------------------------- Color Mapping by Cluster --------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# Define color palette for clusters (adjust as needed)
cluster_colors = {
    0: 'blue',
    1: 'green',
    2: 'red',
    3: 'purple',
    4: 'orange',
    5: 'darkred',
    6: 'lightblue',
    7: 'darkgreen'
}

# --------------------------------------------------------------------------------------------------------
# ---------------------------- Combine Address and City --------------------------------------------------
# --------------------------------------------------------------------------------------------------------

# ensure consistent formatting in address columns
df['Building Address'] = df['Building Address'].str.strip().str.title()
df['Building Address City'] = df['Building Address City'].str.strip().str.title()
df['Building Address Zip  Code'] = df['Building Address Zip  Code']\
                                    .astype(float)\
                                    .astype(int)\
                                    .astype(str)\
                                    .str.zfill(5)

# combine address and city columns into one column
df['Full Address'] = (df['Building Address'] + ', ' +
                      df['Building Address City'] + ', ' + 'MA ' +
                      df['Building Address Zip  Code'])

# --------------------------------------------------------------------------------------------------------
# ---------------------------- Geocode the Combined Address ----------------------------------------------
# --------------------------------------------------------------------------------------------------------

# initialize the Nominatim Geocoder
geolocator = Nominatim(user_agent='berdo-geocoder')
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

# apply geocoding
df[['Latitude', 'Longitude']] = df['Full Address'].apply(geocode_address)

# drop rows where geocoding failed
df.dropna(subset=['Latitude', 'Longitude'], inplace=True)

# save the geocoded data
df.to_csv('../data-files/berdo_data_files/exploratory-analysis-results/geocoded_clustered_data.csv', index=False)
print("Geocoded data saved to geocoded_clustered_data.csv")

# --------------------------------------------------------------------------------------------------------
# ---------------------------- Map the Clusters on Boston Map --------------------------------------------
# --------------------------------------------------------------------------------------------------------

# create a base map centered on boston
boston_map = folium.Map(location=[42.3601, -71.0589], zoom_start=12)
marker_cluster = MarkerCluster().add_to(boston_map)

# add markers for each building
for idx, row in df.iterrows():
    # find appropriate cluster
    cluster_label = row['Cluster_8']
    marker_color = get_cluster_color(cluster_label)

    # create market for map
    folium.Marker(
        location=[row['Latitude'], row['Longitude']],
        popup=f"Cluster: {row['cluster_label']}\nEUI: {row['Site EUI (Energy Use Intensity kBtu/ft2)']:.2f}",
        icon=folium.Icon(color='blue', icon='info-sign')
    ).add_to(marker_cluster)

# save the map as an HTML file
boston_map.save('../data-files/berdo_data_files/exploratory-analysis-results/boston_map_8_clusters.html')
print("Map saved as boston_map_8_clusters.html")




