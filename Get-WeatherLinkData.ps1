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
	0x0000, 0x1021, 0x2042, 0x3063, 0x4084, 0x50A5, 0x60C6, 0x70E7,
	0x8108, 0x9129, 0xA14A, 0xB16B, 0xC18C, 0xD1AD, 0xE1CE, 0xF1EF,
	0x1231, 0x0210, 0x3273, 0x2252, 0x52B5, 0x4294, 0x72F7, 0x62D6,
	0x9339, 0x8318, 0xB37B, 0xA35A, 0xD3BD, 0xC39C, 0xF3FF, 0xE3DE,
	0x2462, 0x3443, 0x0420, 0x1401, 0x64E6, 0x74C7, 0x44A4, 0x5485,
	0xA56A, 0xB54B, 0x8528, 0x9509, 0xE5EE, 0xF5CF, 0xC5AC, 0xD58D,
	0x3653, 0x2672, 0x1611, 0x0630, 0x76D7, 0x66F6, 0x5695, 0x46B4,
	0xB75B, 0xA77A, 0x9719, 0x8738, 0xF7DF, 0xE7FE, 0xD79D, 0xC7BC,
	0x48C4, 0x58E5, 0x6886, 0x78A7, 0x0840, 0x1861, 0x2802, 0x3823,
	0xC9CC, 0xD9ED, 0xE98E, 0xF9AF, 0x8948, 0x9969, 0xA90A, 0xB92B,
	0x5AF5, 0x4AD4, 0x7AB7, 0x6A96, 0x1A71, 0x0A50, 0x3A33, 0x2A12,
	0xDBFD, 0xCBDC, 0xFBBF, 0xEB9E, 0x9B79, 0x8B58, 0xBB3B, 0xAB1A,
	0x6CA6, 0x7C87, 0x4CE4, 0x5CC5, 0x2C22, 0x3C03, 0x0C60, 0x1C41,
	0xEDAE, 0xFD8F, 0xCDEC, 0xDDCD, 0xAD2A, 0xBD0B, 0x8D68, 0x9D49,
	0x7E97, 0x6EB6, 0x5ED5, 0x4EF4, 0x3E13, 0x2E32, 0x1E51, 0x0E70,
	0xFF9F, 0xEFBE, 0xDFDD, 0xCFFC, 0xBF1B, 0xAF3A, 0x9F59, 0x8F78,
	0x9188, 0x81A9, 0xB1CA, 0xA1EB, 0xD10C, 0xC12D, 0xF14E, 0xE16F,
	0x1080, 0x00A1, 0x30C2, 0x20E3, 0x5004, 0x4025, 0x7046, 0x6067,
	0x83B9, 0x9398, 0xA3FB, 0xB3DA, 0xC33D, 0xD31C, 0xE37F, 0xF35E,
	0x02B1, 0x1290, 0x22F3, 0x32D2, 0x4235, 0x5214, 0x6277, 0x7256,
	0xB5EA, 0xA5CB, 0x95A8, 0x8589, 0xF56E, 0xE54F, 0xD52C, 0xC50D,
	0x34E2, 0x24C3, 0x14A0, 0x0481, 0x7466, 0x6447, 0x5424, 0x4405,
	0xA7DB, 0xB7FA, 0x8799, 0x97B8, 0xE75F, 0xF77E, 0xC71D, 0xD73C,
	0x26D3, 0x36F2, 0x0691, 0x16B0, 0x6657, 0x7676, 0x4615, 0x5634,
	0xD94C, 0xC96D, 0xF90E, 0xE92F, 0x99C8, 0x89E9, 0xB98A, 0xA9AB,
	0x5844, 0x4865, 0x7806, 0x6827, 0x18C0, 0x08E1, 0x3882, 0x28A3,
	0xCB7D, 0xDB5C, 0xEB3F, 0xFB1E, 0x8BF9, 0x9BD8, 0xABBB, 0xBB9A,
	0x4A75, 0x5A54, 0x6A37, 0x7A16, 0x0AF1, 0x1AD0, 0x2AB3, 0x3A92,
	0xFD2E, 0xED0F, 0xDD6C, 0xCD4D, 0xBDAA, 0xAD8B, 0x9DE8, 0x8DC9,
	0x7C26, 0x6C07, 0x5C64, 0x4C45, 0x3CA2, 0x2C83, 0x1CE0, 0x0CC1,
	0xEF1F, 0xFF3E, 0xCF5D, 0xDF7C, 0xAF9B, 0xBFBA, 0x8FD9, 0x9FF8,
	0x6E17, 0x7E36, 0x4E55, 0x5E74, 0x2E93, 0x3EB2, 0x0ED1, 0x1EF0
)

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
		$Result = New-Object System.Byte[] 100

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
		#Save all the results
		While($Stream.DataAvailable)
		{
			$Read = $Stream.Read($Result, 0, 100)
		}
		$Writer.Close()
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

	return ([math]::round(($Temp-32)*(5/9),1))

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
	[System.Collections.ArrayList]$data = (Get-Telnet -RemoteHost $IpAddress -port $Port -Commands "LOOP 1")

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
				write-warning "Bad read. Read 100 bytes from the weather station"
				$data | out-file -filepath $LogFile
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
	$BarometricPressure = [double][System.BitConverter]::ToInt16($data,7) / 1000
	$InsideTemp =  [System.BitConverter]::ToInt16($data,9) / 10
	$OutsideTemp = [System.BitConverter]::ToInt16($data,12) / 10
	$WindSpeed = [double]$data[14]
	$RainRate = ([System.BitConverter]::ToInt16($data,41)) * 0.01

	if ($Metric)
	{
		$InsideTemp  = Convert-F2C $InsideTemp
		$OutsideTemp = Convert-F2C $OutsideTemp
		$BarometricPressure = $BarometricPressure * 33.86389
		$WindSpeed = [math]::round($WindSpeed * 1.609,0)
		$RainRate = ([System.BitConverter]::ToInt16($data,41)) * 0.2
	}

	$info = [ordered]@{
		"Success" = $True;
		"Metric" = $Metric;
		"InsideTemperature" = [math]::round($InsideTemp,1)
		"InsideHumidity" = [double]$data[11]
		"OutsideTemperature" = [math]::round($OutsideTemp,1)
		"OutsideHumidity" = [double]$data[33]
		"WindSpeed" = $WindSpeed
		"WindDirection" = [double]([System.BitConverter]::ToInt16($data,16))
		"BarometricPressure" = [math]::round($BarometricPressure,3)
		"RainRate" = $RainRate
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
		"WindSpeed" = $null;
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
		if ($Success)
		{
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
						$UnitElement.InnerText = if ($metric) { "&#8451;" } else { "&#8457;" };
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
						$UnitElement.InnerText = if ($metric) { "&#8451;" } else { "&#8457;" };
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
					'WindSpeed'
					{
						$channelelement.innertext = 'Wind Speed';
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
						$ChartElement.InnerText = '0';
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
		}
		else
		{
			$child = $doc.CreateNode("element","error",$null)
			$child.InnerText = '1';
			$root.AppendChild($child) | Out-Null
			$child = $doc.CreateNode("element","text",$null)
			$child.InnerText = 'error';
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
# Cross-checked CRC table from here: https://stackoverflow.com/questions/17196743/crc-ccitt-implementation
# NetworkStream.Read Method: https://docs.microsoft.com/en-us/dotnet/api/system.net.sockets.networkstream.read?view=netframework-4.8
