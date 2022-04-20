import datetime

import requests

import config

broker_token = config.BROKER_TOKEN
weather_token = config.WEATHER_TOKEN
weather_coord = config.WEATHER_COORD
broker_api = 'https://www.alphavantage.co'


def get_weather():
    url = f'https://api.openweathermap.org/data/2.5/onecall?lat={weather_coord[0]}&lon={weather_coord[1]}&units=metric&appid={weather_token}'
    print(f"Getting weather: {url}")
    weather_data = requests.get(url).json()
    forecast = weather_data['current']
    return f'Temp: {round(forecast["temp"])} °C Max: {round(weather_data["daily"][0]["temp"]["max"])}°C Pressure: {forecast["pressure"]} hPa'


def get_open(symbol):
    url = f'{broker_api}/query?function=TIME_SERIES_INTRADAY&symbol={symbol}&interval=60min&apikey={broker_token}'
    print(f"Getting open: {url}")
    r = requests.get(url)
    cur_series = r.json()
    cur_series = cur_series['Time Series (60min)']
    cur_latest = list(cur_series.keys())[0]
    return cur_series[cur_latest]['1. open']


def get_crypto(currency):
    url = f"{broker_api}/query?function=CURRENCY_EXCHANGE_RATE&from_currency={currency}&to_currency=USD&apikey={broker_token}"
    print(f"Getting Crypto: {url}")
    r = requests.get(url)
    cur_series = r.json()
    return cur_series['Realtime Currency Exchange Rate']['5. Exchange Rate']


def get_forex():
    url = f"{broker_api}/query?function=FX_INTRADAY&from_symbol=USD&to_symbol=EUR&interval=60min&apikey={broker_token}"
    print(f"Getting forex: {url}")
    r = requests.get(url)
    cur_series = r.json()
    cur_series = cur_series['Time Series FX (60min)']
    cur_latest = list(cur_series.keys())[0]
    return cur_series[cur_latest]['1. open']


def main():
    def b(inp):
        return str(round(float(inp) * float(ratio), 2)).ljust(3) + " EUR"

    ratio = get_forex()
    sInfo = [
        f"{'Calendar week:':<18}" + str(datetime.datetime.now().isocalendar()[1]),
        f"{'Bitcoins worth:':<18}" + b(get_crypto('BTC')),
        f"{'PayPal worth:':<18}" + b(get_open('PYPL')),
        f"{'Apple worth:':<18}" + b(get_open('AAPL')),
        f"{'Forecast:':<18}" + get_weather(),
        f"{'Refreshed:':<18}" + str(datetime.datetime.now().strftime("%d.%m.%y"))
    ]

    print(sInfo)
    with open("sInfo.txt", mode='w', encoding="utf-8-sig") as file:
        for line in sInfo:
            file.write(str(line) + "\n")


if __name__ == "__main__":
    main()
else:
    print("This script cannot be imported")
