from Services.WeatherService import WeatherService
from Services.LocationService import LocationService

url =  LocationService.get_aqi_url()
aiq = WeatherService.take_data_by_url(url)

print(aiq)
print(WeatherService.parse_air_quality_data(aiq) )

