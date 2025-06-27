#requires -version 5.1
#requires -module CimCmdlets

#this is a sample PowerShell script showing how you might embed a CSS style sheet in your output
#This script will generate a report for the localhost by default

[cmdletbinding()]
Param(
    #The output path for your file. It will use the format <Computername>-SystemReport.htm
    [string]$Path = "."
)

$fragments = @()`
#ComputerSystem
$cs = Get-CimInstance Win32_ComputerSystem -ov c |
    Select-Object Manufacturer, Model, SystemFamily, SystemSKUNumber, SystemType, NumberOf*, TotalPhysicalMemory

$fragments += "<H2>$($c.CimClass.CimClassName.split("_")[1])</H2>"
$fragments += $cs | ConvertTo-Html -Fragment -as Table

#volumes
$vol = Get-CimInstance win32_volume -ov c | Select-Object Name, Label, Freespace, Capacity
$fragments += "<H2>$($c.CimClass.CimClassName.split("_")[1])</H2>"
$fragments += $vol | ConvertTo-Html -Fragment -as Table

#processor
$cpu = Get-CimInstance win32_processor -ov c | Select-object DeviceID, Name, Caption, MaxClockSpeed, *CacheSize, NumberOf*, SocketDesignation, *Width, Manufacturer
$fragments += "<H2>$($c.CimClass.CimClassName.split("_")[1])</H2>"
$fragments += $cpu | ConvertTo-Html -Fragment -as List

#memory
$mem = Get-CimInstance win32_PhysicalMemory -ov c| Select-Object BankLabel, Capacity, DataWidth, Speed
$fragments += "<H2>$($c.CimClass.CimClassName.split("_")[1])</H2>"
$fragments += $mem | ConvertTo-Html -Fragment -as Table

#network adapter
$net = Get-NetAdapter -Physical -ov c | Select-Object Name, InterfaceDescription, LinkSpeed
$fragments += "<H2>NetworkAdapter</H2>"
$fragments += $net | ConvertTo-Html -Fragment -as Table

#USB
$usb = Get-CimInstance Win32_USBController -ov c | Select-Object Name, Manufacturer, Description
$fragments += "<H2>$($c.CimClass.CimClassName.split("_")[1])</H2>"
$fragments += $usb | ConvertTo-Html -Fragment -as Table

#video
$display = Get-CimInstance Win32_DisplayConfiguration -ov c | Select-Object DeviceName, BitsPerPel, DisplayFrequency, DriverVersion, Pels*
$video = Get-CimInstance CIM_PCVideoController | Select-Object Name, Adapter*, Driver*, VideoModeDescription
$fragments += "<H2>Video</H2>"
$fragments += $display | ConvertTo-Html -Fragment -as Table
$fragments += $video | ConvertTo-Html -Fragment -as Table

#sound
$sound = Get-CimInstance Win32_SoundDevice -ov c | Select-Object Name, Manufacturer
$fragments += "<H2>$($c.CimClass.CimClassName.split("_")[1])</H2>"
$fragments += $sound | ConvertTo-Html -Fragment -as Table

#pointer
$point = Get-CimInstance Win32_PointingDevice -ov c | Select-Object Name, Manufacturer, HardwareType
$fragments += "<H2>$($c.CimClass.CimClassName.split("_")[1])</H2>"
$fragments += $point | ConvertTo-Html -Fragment -as Table

#keyboard
$kbd = Get-CimInstance Win32_Keyboard -ov c | Select-Object Name, Layout, NumberOfFunctionKeys, Description
$fragments += "<H2>$($c.CimClass.CimClassName.split("_")[1])</H2>"
$fragments += $kbd | ConvertTo-Html -Fragment -as Table

#create HTML report
$title = "System Configuration Report"
#this must be left justified
#embed the CSS in between the <style> tags
$head = @"
<Title>$title</Title>
<style>
body { background-color:#FFFF8F;
       font-family:Tahoma;
       font-size:10pt; }
td, th { border:1px solid black;
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px }
tr:nth-child(odd) {background-color: LightGray}
table { width:95%;margin-left:5px; margin-bottom:20px;}
</style>
<br>
<H1>$Title</H1>
"@

$out = Join-Path -Path $path -ChildPath "$($env:computername)-SystemReport.htm"
ConvertTo-Html -Head $head -Body $fragments | Out-File -FilePath $out
Get-Item $out
