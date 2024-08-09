import shutil
import requests
import sys
from Services.WeatherService import WeatherService
from Services.LocationService import LocationService
from Services.ExcelService import ExcelService
from Services.BackupManager import BackupManager

# Get current location example
'''
url =  LocationService.get_aqi_url()
aiq = WeatherService.take_data_by_url(url)

print(WeatherService.parse_air_quality_data(aiq))
'''

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
# if BackupManager.is_file_valid('data.xlsx'):
#     ExcelService.create_file()
# else:
#     print("data.xlsx is corrupted. Attempting to restore the latest backup.")
#     BackupManager.restore_latest_backup()

#process all data
ExcelService.create_file()
locations = LocationService.get_location_list()

for location in locations:
    aiq_string = WeatherService.take_data_by_url(location["station_url"])
    aiq = WeatherService.parse_air_quality_data(aiq_string)
    ExcelService.create_tab(location["name"])
    ExcelService.format_tab(location["name"])
    ExcelService.write_to_file(location["name"], aiq)
    
