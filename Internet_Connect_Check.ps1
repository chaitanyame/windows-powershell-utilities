

<#
    .SYNOPSIS
    Script for checking internet service is working or not
    
    .DESCRIPTION
      this script has to be run under task scheduler as background process.
	  
    AUTHOR: chaitanya
    

#>

Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer


$speak.SelectVoice('Microsoft Zira Desktop')


#region WriteLog function

function WriteLog($LogMessage, $LogDateTime, $LogType) 
{

	write-host 
	"$LogType, ["+ $LogDateTime +"]: "+ $LogMessage | Add-Content -Path $LogFilepath
}
#endregion

# Get Start Time
$startTime = (Get-Date)

$RunTime =get-date -Format "MMdyyyhhmmss"

# Get build folder parent directory
$scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $scriptpath


# Get Application folder path in the build folder

$LogFolderPath = $ScriptDir + "\" + "Logs" 


# Check if Log  folder already exists. If not create a folder for logging purposes 

if(!(Test-Path $LogFolderPath))
{
	New-Item -ItemType directory -Path $LogFolderPath 
}


[string] $logdate =get-date -Format "yyyyMMdd"


$LogFolderFilepath =$LogFolderPath + "\" + "$logdate" 

if(!(Test-Path $LogFolderFilepath))
{
	New-Item -ItemType directory -Path $LogFolderFilepath 
}


# creating logfile path string
$LogFilepath =$LogFolderFilepath +"\"+ "InternetLogfile.txt"


$LogDateTime = get-date

WriteLog "***Checking Internet Connectivity Status ..." $LogDateTime "Information" 


$currentDate = (Get-Date).tostring(“dd-MM-yyyy”)


if (-not(test-Connection -ComputerName www.google.com -Count 4 -Quiet))
{


$speak.Speak('Hello Chaitanya... There is Internet Connectivity issue. I am Unable to Connect ,please check')

$LogDateTime = get-date
WriteLog "**Lookslike Internet Connectivity is Lost . Please check it...."  $LogDateTime "Information"

}

else
{

$LogDateTime = get-date

WriteLog "**Internet Connectivity is good , Exiting ...."  $LogDateTime "Information"

return

}
