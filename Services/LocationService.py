import urllib.request
import json

class LocationService:

    @staticmethod
    def get_location_by_IP():
        with urllib.request.urlopen("https://geolocation-db.com/json") as url:
            data = json.loads(url.read().decode())
            return data

    @staticmethod
    def get_aqi_url():

        postal = LocationService.get_location_by_IP()["postal"]

        f = open('locations.json')

        # returns JSON object as
        # a dictionary
        data = json.load(f)

        # Iterating through the json
        # list
        for elem in data:
            if (elem["postal"] == postal):
                return elem["station_url"]


    @staticmethod
    def get_location_list():
        f = open('locations.json')
        # returns JSON object as
        # a dictionary
        data = json.load(f)
        return data
