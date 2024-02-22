Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;

Stop-Process -Name "InfoBoard" -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5
$dir00=(ls C:\Users\$env:Username\AppData\Roaming\NEC *InfoBoard* -Recurse -Directory).FullName
$dir01=(ls C:\Users\$env:Username\AppData\Local *InfoBoard* -Recurse -Directory).FullName
$dir1=(ls "C:\Program Files\WindowsApps\*"  *InfoBoard* -Recurse -Directory).FullName
$dirapp=$dir1+"\infoboard.exe"

if($dir1.length -gt 0 -or $dir01.length -gt 0)
{start-process "$dirapp"
start-sleep -s 10
Stop-Process -Name "InfoBoard" -ErrorAction SilentlyContinue
start-sleep -s 2

 if($dir00.length -gt 0){
 $dirdata00=$dir00+"\Data"
 Copy-Item C:\Jumpstart\batch\IFBsettings\* -Destination $dirdata00 -Force
 }
 
 if($dir01.length -gt 0){
 $dirdata01=$dir01+"\Data"
 Copy-Item C:\Jumpstart\batch\IFBsettings\* -Destination $dirdata01 -Force
 }

 

start-sleep -s 5
start-process "$dirapp"
start-sleep -s 10

}
else{
echo "no Infoboard Installed"}

exit