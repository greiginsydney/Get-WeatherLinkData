# Get-WeatherLinkData


Get-WeatherLinkData.ps1 is PowerShell script to query data from your [Davis Instruments weather station](https://www.davisinstruments.com/weather-monitoring/), primarily as a means of getting that data into a monitoring solution like [PRTG](https://www.paessler.com/prtg).

The default output is PRTG's XML format, but command-line switches add the options for csv and a PowerShell object.

Add a filename and the output will be saved to that file.

Add the "-metric" switch and the values will be converted to km/h and &#8451;.

You'll find more information on my blog, including detailed "how-to" steps to get your weather station data into PRTG:
https://greiginsydney.com/Get-WeatherLinkData.ps1
</br>
\- Greig.
