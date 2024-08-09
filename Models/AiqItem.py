class AiqItem:
    full_address = ""
    station_name = ""
    region = ""
    date = ""
    pm1 = 0
    pm2_5 = 0
    pm10 = 0
    temperature = 0
    humidity = 0
    pressure = 0
    heca_temperature = 0
    heca_humidity = 0
    no = 0
    no2 = 0
    co2 = 0
    o3 = 0
    h2s = 0
    so2 = 0
    ch2o = 0

    def __str__(self):
        return f'''Air Quality Item:
Full address: {self.full_address}
Date: {self.date}
PM1: {self.pm1}
PM2.5: {self.pm2_5}
PM10: {self.pm10}
Temperature: {self.temperature}
Humidity: {self.humidity}
Pressure: {self.pressure}
HECA Temperature: {self.heca_temperature}
HECA Humidity: {self.heca_humidity}
NO: {self.no}
NO2: {self.no2}
CO2: {self.co2}
O3: {self.o3}
H2S: {self.h2s}
SO2: {self.so2}
CH2O: {self.ch2o}'''
