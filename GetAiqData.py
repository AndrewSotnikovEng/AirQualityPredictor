from Services.WeatherService import WeatherService
from Services.LocationService import LocationService
from Services.ExcelService import ExcelService

#Get current location example
'''url =  LocationService.get_aqi_url()
aiq = WeatherService.take_data_by_url(url)

print(WeatherService.parse_air_quality_data(aiq) )
'''

ExcelService.create_file()
#process all data
locations = LocationService.get_location_list()

for location in locations:
    aiq_string = WeatherService.take_data_by_url(location["station_url"])
    aiq = WeatherService.parse_air_quality_data(aiq_string)
    ExcelService.create_tab(location["name"])
    ExcelService.format_tab(location["name"])
    ExcelService.write_to_file(location["name"], aiq)
    
