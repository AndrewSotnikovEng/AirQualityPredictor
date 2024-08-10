import re
import requests
from bs4 import BeautifulSoup
from Models.AiqItem import AiqItem


class WeatherService:

    @staticmethod
    def parse_air_quality_data(data, station_name):
        air_quality_item = AiqItem()

        # Set station name
        air_quality_item.station_name = station_name

        # Extract city and street
        address_match = re.search(r'Рівень забруднення атмосферного повітря за адресою:\s*(.*)', data)
        if address_match:
            air_quality_item.full_address = address_match.group(1)

        # Extract date and time
        date_match = re.search(r'Первинні дані на (\d+ \w+ \d{4}, \d{2}:\d{2})', data)
        if date_match:
            air_quality_item.date = date_match.group(1).strip()

        # Extract PM1
        pm1_match = re.search(r'PM1: ([\d.]+) мкг/м³', data)
        if pm1_match:
            air_quality_item.pm1 = pm1_match.group(1).replace(',', '.')

        # Extract PM2.5
        pm2_5_match = re.search(r'PM2.5: ([\d.]+) мкг/м³', data)
        if pm2_5_match:
            air_quality_item.pm2_5 = pm2_5_match.group(1).replace(',', '.')

        # Extract PM10
        pm10_match = re.search(r'PM10: ([\d.]+) мкг/м³', data)
        if pm10_match:
            air_quality_item.pm10 = pm10_match.group(1).replace(',', '.')

        # Extract temperature
        temperature_match = re.search(r'Температура: ([\d.-]+) °C', data)
        if temperature_match:
            air_quality_item.temperature = temperature_match.group(1).replace(',', '.')

        # Extract humidity
        humidity_match = re.search(r'Відносна вологість: ([\d.]+) %', data)
        if humidity_match:
            air_quality_item.humidity = humidity_match.group(1).replace(',', '.')

        # Extract pressure
        pressure_match = re.search(r'Атмосферний тиск: ([\d.]+) гПа', data)
        if pressure_match:
            air_quality_item.pressure = pressure_match.group(1).replace(',', '.')

        # Extract HECA temperature
        heca_temperature_match = re.search(r'HECA – Температура: ([\d.-]+) °C', data)
        if heca_temperature_match:
            air_quality_item.heca_temperature = heca_temperature_match.group(1).replace(',', '.')

        # Extract HECA humidity
        heca_humidity_match = re.search(r'HECA – Відносна вологість: ([\d.]+) %', data)
        if heca_humidity_match:
            air_quality_item.heca_humidity = heca_humidity_match.group(1).replace(',', '.')

        # Extract NO
        no_match = re.search(r'Оксид азоту \(NO\): ([\d.]+) мкг/м³', data)
        if no_match:
            air_quality_item.no = no_match.group(1).replace(',', '.')

        # Extract NO2
        no2_match = re.search(r'Діоксид азоту \(NO₂\): ([\d.]+) мкг/м³', data)
        if no2_match:
            air_quality_item.no2 = no2_match.group(1).replace(',', '.')

        # Extract CO2
        co2_match = re.search(r'Вуглекислий газ \(CO₂\): ([\d.]+) мг/м³', data)
        if co2_match:
            air_quality_item.co2 = co2_match.group(1).replace(',', '.')

        # Extract O3
        o3_match = re.search(r'Озон \(O₃\): ([\d.]+) мкг/м³', data)
        if o3_match:
            air_quality_item.o3 = o3_match.group(1).replace(',', '.')

        # Extract H2S
        h2s_match = re.search(r'Сірководень \(H₂S\): ([\d.]+) мкг/м³', data)
        if h2s_match:
            air_quality_item.h2s = h2s_match.group(1).replace(',', '.')

        # Extract SO2
        so2_match = re.search(r'Діоксид сірки \(SO₂\): ([\d.]+) мкг/м³', data)
        if so2_match:
            air_quality_item.so2 = so2_match.group(1).replace(',', '.')

        # Extract CH2O
        ch2o_match = re.search(r'Формальдегід \(CH₂O\): ([\d.]+) мкг/м³', data)
        if ch2o_match:
            air_quality_item.ch2o = ch2o_match.group(1).replace(',', '.')

        return air_quality_item

    @staticmethod
    def take_data_by_url(url):
        response = requests.get(url)

        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            data1 = soup.select_one('h1.text-center')
            data2 = soup.select_one('div.col-md-6:nth-child(1) > h4:nth-child(14)')
            data3 = soup.select_one('div.col-md-6:nth-child(1) > p:nth-child(15)')

            return data1.text.strip() + "\n" + data2.text.strip() + "\n" + data3.text.strip()
        else:
            print("Failed to retrieve the page. Status code:", response.status_code)
            return None
