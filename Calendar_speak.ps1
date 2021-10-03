
Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer


Add-Type -AssemblyName System.Windows.Forms
Add-Type -Name ConsoleUtils -Namespace WPIA -MemberDefinition @'
   [DllImport("Kernel32.dll")]
   public static extern IntPtr GetConsoleWindow();
   [DllImport("user32.dll")]
   public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'@

# Hide Powershell window
$hWnd = [WPIA.ConsoleUtils]::GetConsoleWindow()
[WPIA.ConsoleUtils]::ShowWindow($hWnd, 0)




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
$LogFilepath =$LogFolderFilepath +"\"+ "CalendarLog.txt"



$currentDate = (Get-Date).tostring(“dd-MM-yyyy”)

 $LogDateTime = get-date
 WriteLog "**Starting the script   $LogDateTime " $LogDateTime "Information" 

 

     WriteLog "**Hello ....  $LogDateTime " $LogDateTime "Information" 

    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

    $speak.SelectVoice('Microsoft Zira Desktop')


    # load the required .NET types
    Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
   
    
    # access Outlook object model
    $outlook = New-Object -ComObject outlook.application

      

    # connect to the appropriate location
    $namespace = $outlook.GetNameSpace('MAPI')

    $chaiMailBox=$namespace.folders| Where-Object {$_.Name -eq 'chaitanya.talasila@techvedika.com'}

    $currentDatetime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"

    #Start-Sleep 5  # sleeping for 45 seconds

    WriteLog "**Hello ....  $LogDateTime " $LogDateTime "Information" 

     WriteLog "**Hello $($outlook) ....  $LogDateTime " $LogDateTime "Information" 

    foreach ($mailfolder in $chaiMailBox.Folders)

    {
       
        $LogDateTime = get-date

        WriteLog "**Traversing the folder path $($mailfolder.FolderPath)   $LogDateTime " $LogDateTime "Information" 


        if ($mailfolder.FolderPath -like '*Calendar*')

        {

            

             $appts=$mailfolder.items             # Per referenced article
             $appts.Sort("[Start]")               # the sequence of these
             $appts.IncludeRecurrences = $true    # is very important

             $start = (Get-Date)
             $end = $start.AddDays(1)
             $filter = "[Start] >= `"$($start.toString('g'))`" and [End] <= `"$($end.toString('g'))`" "
             $someitems = $appts.restrict($filter)
             $AllCalendarItems=$someitems | Select-Object -Property Subject, Start, End, Duration, Location
            
       
         }
        
    }


    if ($AllCalendarItems)

    {

    $CalendarItems=$AllCalendarItems | where {$_.Start -ge $currentDatetime}

     $LogDateTime = get-date
     WriteLog "**Calender Item records found :$($CalendarItems.count)  $LogDateTime " $LogDateTime "Information" 
         
    $speak.Speak("Hello Chaitanya... you have $($CalendarItems.count) events to attend.")

    foreach ($item in $CalendarItems)

        {

        $speak.Speak("At $($item.Start.ToShortTimeString()), there is  $($item.Subject)")

    
        }


     }






