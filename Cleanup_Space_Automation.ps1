


<#
    .SYNOPSIS
    Script for Cleanup space 
    
    .DESCRIPTION
     Script for Cleanup space 

    AUTHOR: chaitanya
    

#>


$DaysToDeleteWin=1


$tempfolders = @(“C:\Windows\Temp\*”, “C:\Windows\Prefetch\*”, “C:\Documents and Settings\*\Local Settings\temp\*”, “C:\Users\*\Appdata\Local\Temp\*”)


Get-ChildItem $tempfolders  -Recurse -Force -Verbose -ErrorAction SilentlyContinue| `
Where-Object { ($_.LastWriteTime -lt $(Get-Date).AddDays(-$DaysToDeleteWin)) } | `
remove-item -force -Verbose -recurse -ErrorAction SilentlyContinue 


if ( get-process -name chrome)
{

taskkill /F /IM "chrome.exe"
Start-Sleep -Seconds 5

}


$DaysToDelete = 4

$crLauncherDir = "C:\Documents and Settings\%USERNAME%\Local Settings\Application Data\Chromium\User Data\Default"
$chromeDir = "C:\Users\*\AppData\Local\Google\Chrome\User Data\Default"
$chromeSetDir = "C:\Users\*\Local Settings\Application Data\Google\Chrome\User Data\Default"

$chromeCachePaths = @("*Archived History*", "*Cache*", "*Cookies*", "*History*", "*Web Data*")

$chromeCachePaths|ForEach-Object {
$item = $_ 

#for directories

Get-ChildItem $chromeDir -Directory -Force -Recurse -ErrorAction SilentlyContinue |`
    Where-Object {$_.Name -like $item}
}| ForEach-Object -Process {

 get-childitem -path $_.fullname -Recurse |
 Where-Object {($_.LastWriteTime -lt $(Get-Date).AddDays(-$DaysToDelete))}| `  
 remove-item -force -Verbose -recurse -ErrorAction SilentlyContinue 

}



if ( get-process -name firefox)
{

taskkill /F  /im "firefox.exe"
Start-Sleep -Seconds 5

}


$firefoxdir='C:\Users\*\AppData\Local\Mozilla\Firefox\Profiles'
$firefoxdir2='C:\Users\*\AppData\Roaming\Mozilla\Firefox\Profiles'

Get-ChildItem $firefoxdir -Recurse -Force -ErrorAction SilentlyContinue | 
    Where-Object { ($_.LastWriteTime -lt $(Get-Date).AddDays(-$DaysToDelete)) } | ForEach-Object -Process { Remove-Item $_ -force -Verbose -recurse -ErrorAction SilentlyContinue }


Get-ChildItem $firefoxdir2 -Recurse -Filter *.sqlite -Force -ErrorAction SilentlyContinue | 
    Where-Object { ($_.LastWriteTime -lt $(Get-Date).AddDays(-$DaysToDelete)) } | ForEach-Object -Process { Remove-Item $_ -force -Verbose -recurse -ErrorAction SilentlyContinue }





if ( get-process -name msedge)
{

taskkill /F  /im "msedge.exe"
Start-Sleep -Seconds 5

}

$EdgeAppData = "C:\Users\*\AppData\Local\Microsoft\Edge\User Data\Default"
$possibleCachePaths = @('Cache','Cache2\entries','Cookies','History','Top Sites','Visited Links','Web Data','Media History','Cookies-Journal')


$possibleCachePaths|  ForEach-Object {
$itemname = $_ 

#for directories

Get-ChildItem $EdgeAppData -Directory -Force -Recurse -ErrorAction SilentlyContinue |`
    Where-Object {$_.Name -like $itemname}
}| ForEach-Object -Process {

  get-childitem -path $_.fullname -Recurse |
 Where-Object {($_.LastWriteTime -lt $(Get-Date).AddDays(-$DaysToDelete))}| `  
 remove-item -force -Verbose -recurse -ErrorAction SilentlyContinue 

}







$recycleBinContents=(New-Object -ComObject Shell.Application).NameSpace(0x0a).Items()

if ($recycleBinContents)

{

Clear-RecycleBin -Force

}
else
{

return

}



