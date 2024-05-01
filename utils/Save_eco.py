from Services.WeatherService import WeatherService
from Services.LocationService import LocationService


# Example usage:
url_to_parse = "https://www.saveecobot.com/station/19104"
service = WeatherService()
aiq = service.take_data_by_url(url_to_parse)
print(LocationService.get_location_by_IP())
print(aiq)
print(service.parse_air_quality_data(aiq) )
