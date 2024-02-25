class AiqItem:
    full_address = ""
    station_name = ""
    region = ""
    date = ""
    pm2_5 = 0
    pm10 = 0
    temperature = 0
    humidity = 0
    pressure = 0
    heca_temperature = 0
    heca_humidity = 0

    def __str__(self):
        return f'''Air Quality Item:
Full address: {self.full_address}
Date: {self.date}
PM2.5: {self.pm2_5}
PM10: {self.pm10}
Temperature: {self.temperature}
Humidity: {self.humidity}
Pressure: {self.pressure}
HECA Temperature: {self.heca_temperature}
HECA Humidity: {self.heca_humidity}'''