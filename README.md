# Update WP

This will update your Wallpaper, such that the bottom-right corner sayss the following text:

```txt
Calendar week:    <week>
Bitcoins worth:   <BTC> EUR
PayPal   worth:   <PYPL> EUR
Apple    worth:   <APLE> EUR
Forecast:         Temp: <currentTemp> °C Max: <maxTemp> °C Pressure: <pressure> hPa
Refreshed:        <time>
```

where `<>` are inserted respectivly.

You'll need [BGInfo64](https://docs.microsoft.com/en-us/sysinternals/downloads/bginfo) to run a scheduled Task, that updates your Wallpaper.

## Configuration

add a `config.py` with the following values:

```python
BROKER_TOKEN  = "" # Get token at: <https://www.alphavantage.co>
WEATHER_TOKEN = "" # Get Token at: <https://api.openweathermap.org>
WEATHER_COORD = <xCoord>, <yCoord> # coordinates of your city
```
## Tasks

### General

#### Security Options

- When running the taskm use the following user account: `username`
- Run with highest privileges
- Configure for: `Windows 10`

### Triggers

- begin the task: `At log on`
  - `Any user`
- Advanced settings:
  -  Repeat Task every `1 hour`
  -  for a duration of `Indefinitely`
  - Enabled

### Actions

#### Start a program

- Program/script: `<path>\UpdateWP\wp.pyw`
- Args: 
- Start in: `<path>`

#### Start a program

- Program/script: `<path>\UpdateWP\Bginfo64.exe`
- Args: `<path>\UpdateWP\conf.bgi /TIMER:0 /SILENT`
- Start in: `<path>`
