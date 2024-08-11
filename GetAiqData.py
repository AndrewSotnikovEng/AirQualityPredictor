import time
import requests
import sys
import logging
from Services.WeatherService import WeatherService
from Services.LocationService import LocationService
from Services.ExcelService import ExcelService
from Services.BackupManager import BackupManager

logging.basicConfig(
    filename='processing_report.log.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Check connection
try:
    response = requests.get("http://www.google.com", timeout=5)
    if response.status_code == 200:
        print("Connected to the internet")
        logging.info("Connected to the internet")
    else:
        print("Failed to connect to the internet")
        logging.error("Connection error: Failed to connect to the internet")
        sys.exit(0)
except requests.ConnectionError:
    print("Failed to connect to the internet")
    logging.error("Connection error: Failed to connect to the internet")
    sys.exit(0)

# Prepare the Excel file (restore from backup if necessary)
locations = LocationService.get_location_list()
location_names = [location["name"] for location in locations]
ExcelService.prepare_excel_file(location_names)

# Perform backup before any operation
BackupManager.create_backup()
logging.info("Backup created successfully")

total_locations = len(locations)
successful_locations = 0

# Fetch and write data for each location
for index, location in enumerate(locations, start=1):
    print(f"Processing station {index}/{total_locations}: {location['name']}")
    logging.info(f"Processing station {index}/{total_locations}: {location['name']}")
    try:
        aiq_string = WeatherService.take_data_by_url(location["station_url"])
        if aiq_string:
            aiq_item = WeatherService.parse_air_quality_data(aiq_string, location["name"])
            ExcelService.format_tab(location["name"])
            ExcelService.write_to_file([aiq_item])
            successful_locations += 1
            print(f"Data successfully written for station: {location['name']}")
        else:
            logging.warning(f"Station not available: {location['name']}")
    except Exception as e:
        print(f"Failed to process data for {location['name']}: {e}")
        logging.error(f"Failed to process data for {location['name']}: {e}")
    time.sleep(3)

# Log the final status
if successful_locations == total_locations:
    logging.info("Completely successful: All data is available and written to file.")
else:
    logging.warning(f"Processed {successful_locations}/{total_locations} stations successfully.")

print("Data processing completed.")
logging.info("Data processing completed.")

# Add a separator to distinguish between
separator = '-' * 80
logging.info(separator)
