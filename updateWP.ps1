[CmdletBinding()]
param (
    [switch]$regTask
)
Start-Transcript -Path "$PSScriptroot\updateWP.log" -Force

$settings = Get-Content 'config.json' | ConvertFrom-Json
$broker_Token = $settings.broker_Token
$weather_Token = $settings.weather_Token
$weather_latlon = $settings.weather_latlon

$global:broker_data = Invoke-RestMethod -Uri "https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=BTC&to_currency=EUR&apikey=$broker_Token" -Method GET
$global:weather_data = Invoke-RestMethod -Uri "https://api.openweathermap.org/data/2.5/onecall?lat=$($weather_latlon[0])&lon=$($weather_latlon[1])&units=metric&appid=$weather_Token" -Method GET
$btc = $broker_data.'Realtime Currency Exchange Rate'.'5. Exchange Rate'.substring(0, $broker_data.'Realtime Currency Exchange Rate'.'5. Exchange Rate'.indexof('.') + 3)
$forecast = $weather_data.current
$forecaststring = "Temp: " + [math]::Round($forecast.temp) + "°C Max: " + [math]::Round($weather_data.daily[0].temp.max) + "°C Pressure: " + $forecast.pressure + " hPa"
Write-Verbose -Message "$btc"

"Calendar week:	$(Get-Date -UFormat %V)
Bitcoins worth:	$btc EUR
Forecast:	$forecaststring
Refreshed:	$(Get-Date -Format 'HH:mm:ss')" | Out-File "$PSScriptRoot\sInfo.txt"

if ($regTask) {
    # Create Hourly scedule
    if (!(Test-Path "$PSScriptRoot\Update Wallpaper.xml")) { Write-Error -Message "Couldnt find xml data"; return 1 }
    [xml]$XMLData = Get-Content "$PSScriptRoot\Update Wallpaper.xml"
    if (Get-ScheduledTask -TaskName ($XMLData.Task.RegistrationInfo.URI).trimstart("\") -ErrorAction SilentlyContinue) {
        Write-Verbose -Message "$($($XMLData.Task.RegistrationInfo.URI).trimstart(""\"")) already exists"
        Stop-Transcript
        return 0
    }

    Write-Verbose -Message "Create Scheduled Task?"
    cmd /c choice /T 8 /M "Would you like run every hour?" /C yn /D n
    if ($LASTEXITCODE -eq 2) { Stop-Transscript; exit 0 }
    if (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
        Write-Warning -Message "You need Admin rights to Create a Scheduled Task"
    }

    Start-Process PowerShell -Verb Runas -ArgumentList "-File ""$PSScriptRoot\regTask.ps1"" -verbose" -WindowStyle Hidden
    return $global:data
}
else {
    Write-Verbose -Message "Wallpaper only switch set"
    Stop-Transcript
    return 0
}
