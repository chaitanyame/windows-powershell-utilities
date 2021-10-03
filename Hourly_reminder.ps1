Add-Type -AssemblyName System.speech

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

Import-Module -Name 'AudioDeviceCmdlets'

$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

$speak.SelectVoice('Microsoft Zira Desktop')

$dict=@{}

[int]$curDefaultAudio = [int](Get-AudioDevice -PlaybackVolume).split('%')[0]

$maxVolume=100


# Get build folder parent directory
$scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $scriptpath


$currHour=(get-date).Hour

while ($currHour -in 10..22)
{


    if ($dict[$currHour] -eq 1)
    {

    Start-Sleep 10
    $currHour=(get-date).Hour
    continue

    }


    if ($currHour -ge 22)
        {

            $speak.Speak('Hello Chaitanya... Time is now 10 PM. please log off from the work')

            exit

        }
    else 
    
        {
            if((get-date).Minute -eq 0)
            {

                $dict.Add($currHour,1)
                $PlayWav=New-Object System.Media.SoundPlayer

                #Sets volume to 100%
                Set-AudioDevice -PlaybackVolume $maxVolume

                $filepath=$ScriptDir+'\mixkit-bell-notification-933.wav'

                $PlayWav.SoundLocation=$filepath

                $PlayWav.playsync()

                Start-Sleep 2

                $currHour=(get-date).Hour

                #Sets volume to 100%
                Set-AudioDevice -PlaybackVolume $curDefaultAudio

                continue

            }

            else
            {

                continue

            }
        }



}


 



