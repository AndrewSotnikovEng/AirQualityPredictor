import shutil
import requests
import sys
from Services.WeatherService import WeatherService
from Services.LocationService import LocationService
from Services.ExcelService import ExcelService
from Services.BackupManager import BackupManager

# Check connection
try:
    response = requests.get("http://www.google.com", timeout=5)
    if response.status_code == 200:
        print("Connected to the internet")
    else:
        print("Failed to connect to the internet")
        sys.exit(0)
except requests.ConnectionError:
    print("Failed to connect to the internet")
    sys.exit(0)

# Perform backup before any operation
BackupManager.create_backup()

# Process all data
locations = LocationService.get_location_list()

location_names = [location["name"] for location in locations]

# Create or load the Excel file and ensure the necessary sheets exist
ExcelService.create_file(location_names)

# Fetch and write data for each location
for location in locations:
    try:
        aiq_string = WeatherService.take_data_by_url(location["station_url"])
        if aiq_string:
            aiq_item = WeatherService.parse_air_quality_data(aiq_string, location["name"])
            ExcelService.format_tab(location["name"])
            ExcelService.write_to_file([aiq_item])
    except Exception as e:
        print(f"Failed to process data for {location['name']}: {e}")

print("Data processing completed.")
