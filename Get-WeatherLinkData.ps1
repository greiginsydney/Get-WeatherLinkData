<# 
.SYNOPSIS
	This script queries your Davis WeatherLink weather station and outputs the primary values in a range of formats.

.DESCRIPTION
	This script queries your Davis WeatherLink weather station and outputs the primary values in a range of formats.
	It reports in Imperial by default, but will output metric with the "-metric" switch.
	The default output is a PowerShell object to the screen, with other options being CSV & the XML format for PRTG.
	Add a Filename and the same output will be saved to the named file.

.NOTES
	Version				: 0.0
	Date				: TBA 2020
	Author				: Greig Sheridan
	See the credits at the bottom of the script
	
	WISH-LIST / TODO:
	
	KNOWN ISSUES:
	
	Revision History 	:
				v0.0 TBA 2020
					Initial release

.LINK
	https://greiginsydney.com/Get-WeatherLinkData.ps1 - also https://github.com/greiginsydney/Get-WeatherLinkData

.EXAMPLE
	.\Get-WeatherLinkData.ps1

	Description
	-----------
	Without an IP address, the script will err.

.EXAMPLE
	.\Get-WeatherLinkData.ps1 -IpAddress '10.10.10.10'

	Description
	-----------
	Queries the Weatherlink device at 10.10.10.10 and displays the output to screen and pipeline as an Object.

.EXAMPLE
	.\Get-WeatherLinkData.ps1 -IpAddress '10.10.10.10' -OutputFormat PRTG

	Description
	-----------
	Queries the Weatherlink device at 10.10.10.10 and displays the output on screen in XML format for PRTG.

.EXAMPLE
	.\Get-WeatherLinkData.ps1 -IpAddress '10.10.10.10' -OutputFormat csv -FileName WeatherlinkData.csv

	Description
	-----------
	Queries the Weatherlink device at 10.10.10.10 and displays the output on screen in csv format. The same output is written to the file at WeatherlinkData.csv.
	
.EXAMPLE
	.\Get-WeatherLinkData.ps1 -IpAddress '10.10.10.10' -Metric 

	Description
	-----------
	Queries the Weatherlink device at 10.10.10.10 and displays the output on screen and pipeline as PRTG-format XML data. Values are displayed in metric: degrees in Celsius and speeds in km/h.


.PARAMETER IpAddress
	String. The IP address of your weather station.

.PARAMETER Port
	String. The default port is 22222, but here's the place to change it if you have reason to.

.PARAMETER OutputFormat
	Optional string. Follow with one of Object,PRTG,CSV. The default is PRTG. (See Examples.)

.PARAMETER FileName
	File name (and path if you wish) of a file to which the script will write the data. Any existing file of the same name will be over-written without prompting.

.PARAMETER Metric
	Switch. The data out of the Weatherlink is Imperial by default. If $True (or simply present), the script will convert to degrees Celsius & speeds in km/h.

.PARAMETER Retries
	Integer. How many attempts will be made to get a good read out of the weather station. The default is 4. Great for where you have bad connectivity.

#>

[CmdletBinding(SupportsShouldProcess = $False)]
param(
	[parameter(ValueFromPipeline, ValueFromPipelineByPropertyName = $true, Mandatory)]
	[alias('Ip')][string]$IpAddress,
	[string]$Port="22222",
	[ValidateSet('Object', 'PRTG', 'CSV')]
	[alias('Output')][String]$OutputFormat='PRTG',
	[alias('File')][string]$FileName,
	[switch] $Metric,
	[int]$Retries=4

)

$Error.Clear()		#Clear PowerShell's error variable
$Global:Debug = $psboundparameters.debug.ispresent


#--------------------------------
# START CONSTANTS ---------------
#--------------------------------

[uint16[]]$crc_table = @(
   0x0, 0x1021, 0x2042, 0x3063, 0x4084, 0x50a5, 0x60c6, 0x70e7,`
0x8108, 0x9129, 0xa14a, 0xb16b, 0xc18c, 0xd1ad, 0xe1ce, 0xf1ef,`
0x1231, 0x0210, 0x3273, 0x2252, 0x52b5, 0x4294, 0x72f7, 0x62d6,`
0x9339, 0x8318, 0xb37b, 0xa35a, 0xd3bd, 0xc39c, 0xf3ff, 0xe3de,`
0x2462, 0x3443, 0x0420, 0x1401, 0x64e6, 0x74c7, 0x44a4, 0x5485,`
0xa56a, 0xb54b, 0x8528, 0x9509, 0xe5ee, 0xf5cf, 0xc5ac, 0xd58d,`
0x3653, 0x2672, 0x1611, 0x0630, 0x76d7, 0x66f6, 0x5695, 0x46b4,`
0xb75b, 0xa77a, 0x9719, 0x8738, 0xf7df, 0xe7fe, 0xd79d, 0xc7bc,`
0x48c4, 0x58e5, 0x6886, 0x78a7, 0x0840, 0x1861, 0x2802, 0x3823,`
0xc9cc, 0xd9ed, 0xe98e, 0xf9af, 0x8948, 0x9969, 0xa90a, 0xb92b,`
0x5af5, 0x4ad4, 0x7ab7, 0x6a96, 0x1a71, 0x0a50, 0x3a33, 0x2a12,`
0xdbfd, 0xcbdc, 0xfbbf, 0xeb9e, 0x9b79, 0x8b58, 0xbb3b, 0xab1a,`
0x6ca6, 0x7c87, 0x4ce4, 0x5cc5, 0x2c22, 0x3c03, 0x0c60, 0x1c41,`
0xedae, 0xfd8f, 0xcdec, 0xddcd, 0xad2a, 0xbd0b, 0x8d68, 0x9d49,`
0x7e97, 0x6eb6, 0x5ed5, 0x4ef4, 0x3e13, 0x2e32, 0x1e51, 0x0e70,`
0xff9f, 0xefbe, 0xdfdd, 0xcffc, 0xbf1b, 0xaf3a, 0x9f59, 0x8f78,`
0x9188, 0x81a9, 0xb1ca, 0xa1eb, 0xd10c, 0xc12d, 0xf14e, 0xe16f,`
0x1080, 0x00a1, 0x30c2, 0x20e3, 0x5004, 0x4025, 0x7046, 0x6067,`
0x83b9, 0x9398, 0xa3fb, 0xb3da, 0xc33d, 0xd31c, 0xe37f, 0xf35e,`
0x02b1, 0x1290, 0x22f3, 0x32d2, 0x4235, 0x5214, 0x6277, 0x7256,`
0xb5ea, 0xa5cb, 0x95a8, 0x8589, 0xf56e, 0xe54f, 0xd52c, 0xc50d,`
0x34e2, 0x24c3, 0x14a0, 0x0481, 0x7466, 0x6447, 0x5424, 0x4405,`
0xa7db, 0xb7fa, 0x8799, 0x97b8, 0xe75f, 0xf77e, 0xc71d, 0xd73c,`
0x26d3, 0x36f2, 0x0691, 0x16b0, 0x6657, 0x7676, 0x4615, 0x5634,`
0xd94c, 0xc96d, 0xf90e, 0xe92f, 0x99c8, 0x89e9, 0xb98a, 0xa9ab,`
0x5844, 0x4865, 0x7806, 0x6827, 0x18c0, 0x08e1, 0x3882, 0x28a3,`
0xcb7d, 0xdb5c, 0xeb3f, 0xfb1e, 0x8bf9, 0x9bd8, 0xabbb, 0xbb9a,`
0x4a75, 0x5a54, 0x6a37, 0x7a16, 0x0af1, 0x1ad0, 0x2ab3, 0x3a92,`
0xfd2e, 0xed0f, 0xdd6c, 0xcd4d, 0xbdaa, 0xad8b, 0x9de8, 0x8dc9,`
0x7c26, 0x6c07, 0x5c64, 0x4c45, 0x3ca2, 0x2c83, 0x1ce0, 0x0cc1,`
0xef1f, 0xff3e, 0xcf5d, 0xdf7c, 0xaf9b, 0xbfba, 0x8fd9, 0x9ff8,`
0x6e17, 0x7e36, 0x4e55, 0x5e74, 0x2e93, 0x3eb2, 0x0ed1, 0x1ef0
);

#--------------------------------
# END CONSTANTS -----------------
#--------------------------------


#--------------------------------
# START FUNCTIONS ---------------
#--------------------------------


Function Get-Telnet
{
	Param (
		[Parameter(ValueFromPipeline=$true)]
		[String[]]$Commands = @(),
		[string]$RemoteHost = "",
		[string]$Port = "23",
		[int]$WaitTime = 1000
	)
	#Attach to the remote device, setup streaming requirements
	try
	{
		$Socket = New-Object System.Net.Sockets.TcpClient($RemoteHost, $Port)
	}
	catch {}
	If ($null -ne $Socket)
	{
		$Stream = $Socket.GetStream()
		$Writer = New-Object System.IO.StreamWriter($Stream)
		$Buffer = New-Object System.Byte[] 100
		$Encoding = New-Object System.Text.UnicodeEncoding

		#Now start issuing the commands
		ForEach ($Command in $Commands)
		{
			$Writer.WriteLine($Command)
			$Writer.Flush()
			Start-Sleep -Milliseconds $WaitTime
		}
		#All commands issued, but since the last command is usually going to be
		#the longest let's wait a little longer for it to finish
		Start-Sleep -Milliseconds ($WaitTime)
		$Result = @()
		#Save all the results
		While($Stream.DataAvailable)
		{
			$Read = $Stream.Read($Buffer, 0, 100)
			$Result += ($Encoding.GetString($Buffer, 0, $Read))
		}
		$Socket.Close()
	}
	Else
	{
		write-warning "Unable to connect to host: $($RemoteHost):$($Port)"
		$result = $null
	}
	return $result
}


function Convert-F2C
{
	Param (
		[Parameter(ValueFromPipeline=$true)]
		[int]$Temp
	)

	return [int16](($Temp-32)*(5/9))

}


#--------------------------------
# END FUNCTIONS -----------------
#--------------------------------


$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path -Path $scriptpath
$LogFile = (Join-Path -path $dir -childpath "GetWeatherlinkLOG-")
$LogFile += (Get-Date -format "yyyyMMdd-HHmm") + ".log"

if ($FileName)
{
	#If the user only provided a filename, add the script's path for an absolute reference:
	if ([IO.Path]::IsPathRooted($FileName))
	{
		#It's absolute. Safe to leave.
	}
	else
	{
		#It's relative.
		$FileName = [IO.Path]::GetFullPath((Join-Path -path $dir -childpath $FileName))
	}
}


$Success = $false
$Attempt = 0
:nextAttempt while ($retries - $attempt -ge 0)
{
	$attempt ++
	write-verbose "Attempt #$($attempt)"
	$enc = [system.Text.Encoding]::Unicode
	$result = (Get-Telnet -RemoteHost $IpAddress -port $Port -Commands "LOOP 1")
	[System.Collections.ArrayList]$data = $null
	try
	{
		$data = $enc.GetBytes($result)
	}
	catch {}

	switch ($data.count)
	{
		99
		{
			write-verbose "Good read. Read 99 bytes"
		}

		100
		{
			if ($data[0] -eq 0x06)
			{
				# 'ACK' - we're good
				write-verbose "Read started with an ACK."
				$data.RemoveAt(0)
			}
			elseif ($data[0] -eq 0x21)
			{
				write-warning "Read started with a NAK."
				continue nextAttempt 
			}
			else
			{
				write-warning "Bad read. Read $($data.count ) bytes from the weather station"
				continue nextAttempt 
			}
		}

		default
		{
			write-warning "Bad read. Read $($data.count ) bytes from the weather station"
			continue nextAttempt 
		}
	}

	#Check the CRC
	[uint16]$crc = 0
	foreach ($byte in $data)
	{
		$crc = ($crc_table[([byte]($crc -shr 8) -bxor $byte)]) -bxor [uint16]($crc -shl 8)
	}
	if ($crc -eq 0)
	{
		#We're good!
		$Success = $true
		write-verbose "CRC OK on attempt #$($attempt)"
		$data | out-file -filepath "CRC-good.log" -encoding "UTF8"
		break
	}
	else
	{
		write-warning "CRC Fail on attempt #$($attempt)"
		$data | out-file -filepath $LogFile -encoding "UTF8"
	}
}


if ($Success)
{
	#These values are decoded now as they potentially require metric conversion. All others are decoded as the output object is created:
	$BarometricPressure = [System.BitConverter]::ToInt16($data,7) / 1000
	$InsideTemp =  [System.BitConverter]::ToInt16($data,9) / 10
	$OutsideTemp = [System.BitConverter]::ToInt16($data,12) / 10
	$AvgWindSpeed = [int]$data[15]
	$RainRate = ([System.BitConverter]::ToInt16($data,41)) * 0.01

	if ($Metric)
	{
		$InsideTemp  = Convert-F2C $InsideTemp
		$OutsideTemp = Convert-F2C $OutsideTemp
		$BarometricPressure = [math]::round(($BarometricPressure * 33.86389),1)
		$AvgWindSpeed = [math]::round($AvgWindSpeed * 1.609,1)
		$RainRate = ([System.BitConverter]::ToInt16($data,41)) * 0.2
	}

	$info = [ordered]@{
		"Success" = $True;
		"Metric" = $Metric;
		"InsideTemperature" = "{0:f1}" -f $InsideTemp
		"InsideHumidity" = [int]$data[11]
		"OutsideTemperature" = "{0:f1}" -f $OutsideTemp
		"OutsideHumidity" = [int]$data[33]
		"AvgWindSpeed" = "{0:f1}" -f $AvgWindSpeed
		"WindDirection" = ([System.BitConverter]::ToInt16($data,16))
		"BarometricPressure" = "{0:f1}" -f $BarometricPressure
		"RainRate" = ([System.BitConverter]::ToInt16($data,41))
	}
}
else
{
	$info = [ordered]@{
		"Success" = $False;
		"Metric" = $Metric;
		"InsideTemperature" = $null;
		"InsideHumidity" = $null;
		"OutsideTemperature" = $null;
		"OutsideHumidity" = $null;
		"AvgWindSpeed" = $null;
		"WindDirection" = $null;
		"BarometricPressure" = $null;
		"RainRate" = $null;
	}
}

$resultInfo = New-Object -TypeName PSObject -Property $info

switch ($OutputFormat)
{
	'object'
	{
		$resultInfo
		if ($FileName) { $resultInfo | out-File -FilePath $Filename -Encoding "UTF8" }
	}

	'csv'
	{
		$csvResult = $resultInfo | ConvertTo-Csv -NoTypeInformation
		$csvResult
		if ($FileName) { $CsvResult | out-File -FilePath $Filename -Encoding "UTF8" }
	}

	'prtg'
	{
		[xml]$Doc = New-Object System.Xml.XmlDocument
		$dec = $Doc.CreateXmlDeclaration("1.0","UTF-8",$null)
		$doc.AppendChild($dec) | Out-Null
		$root = $doc.CreateNode("element","prtg",$null)
		$resultInfo.PsObject.Properties | foreach-object `
		{
			if (($_.Name -eq 'Success') -or ($_.Name -eq 'Metric')) { return } #Suppress these in the output to PRTG 
			$child = $doc.CreateNode("element","Result",$null)
			$ChannelElement = $doc.CreateElement('Channel')
			$UnitElement = $doc.CreateElement('customUnit')
			$FloatElement = $doc.CreateElement('float');
			$ValueElement = $doc.CreateElement('value');
			$ChartElement = $doc.CreateElement('showChart');
			$TableElement = $doc.CreateElement('showTable');
			
			switch ($_.Name)
			{
				'InsideTemperature'
				{
					$ChannelElement.InnerText = 'Inside Temperature'
					$UnitElement.InnerText = 'Degrees'
					$FloatElement.InnerText = "1";
					$ChartElement.InnerText = '1';
					$TableElement.InnerText = '1';
				}
				'InsideHumidity'
				{
					$channelelement.innertext = 'Inside Humidity';
					$UnitElement.InnerText = "%";
					$FloatElement.InnerText = "1";
					$ChartElement.InnerText = '1';
					$TableElement.InnerText = '1';
				}
				'OutsideTemperature'
				{
					$channelelement.innertext = 'Outside Temperature';
					$UnitElement.InnerText = "Degrees";
					$FloatElement.InnerText = "1";
					$ChartElement.InnerText = '1';
					$TableElement.InnerText = '1';
				}
				'OutsideHumidity'
				{
					$channelelement.innertext = 'Outside Humidity';
					$UnitElement.InnerText = "%";
					$FloatElement.InnerText = "1";
					$ChartElement.InnerText = '1';
					$TableElement.InnerText = '1';
				}
				'AvgWindSpeed'
				{
					$channelelement.innertext = 'Average Wind Speed';
					$UnitElement.InnerText = if ($metric) { "km/h" } else { "mph" };
					$FloatElement.InnerText = "1";
					$ChartElement.InnerText = '1';
					$TableElement.InnerText = '1';
				}
				'WindDirection'
				{
					$channelelement.innertext = 'Wind Direction';
					$UnitElement.InnerText = "Degrees";
					$FloatElement.InnerText = "1";
					$ChartElement.InnerText = '1';
					$TableElement.InnerText = '1';
				}
				'BarometricPressure'
				{
					$channelelement.innertext = 'Barometric Pressure';
					$UnitElement.InnerText = if ($metric) { "hPa" } else { "Hg" };
					$FloatElement.InnerText = "1";
					$ChartElement.InnerText = '1';
					$TableElement.InnerText = '1';
				}
				'RainRate'
				{
					$channelelement.innertext = 'Rain Rate';
					$UnitElement.InnerText = if ($metric) { "mm/h" } else { "inches/h" };
					$FloatElement.InnerText = "1";
					$ChartElement.InnerText = '1';
					$TableElement.InnerText = '1';
				}
				default { continue }
			}
			$child.AppendChild($ChannelElement)	| Out-Null;
			$child.AppendChild($UnitElement)	| out-null;
			$child.AppendChild($FloatElement)	| out-null;
			$ValueElement.InnerText = $_.Value
			$child.AppendChild($ValueElement)	| out-null;
			$child.AppendChild($ChartElement)	| out-null;
			$child.AppendChild($TableElement)	| out-null;
			#append to root
			$root.AppendChild($child) | Out-Null
		}
		$doc.AppendChild($root) | Out-Null
		$doc.InnerXML
		if ($FileName) { $doc.Save($Filename)}
	}
}


# CREDITS:

# Thank you Martin Pugh for Get-Telnet: https://thesurlyadmin.com/2013/04/04/using-powershell-as-a-telnet-client/
# See also https://community.spiceworks.com/scripts/show/1887-get-telnet-telnet-to-a-device-and-issue-commands
# Hand-crafting XML, thanks to Jeff Hicks: https://www.petri.com/creating-custom-xml-net-powershell

