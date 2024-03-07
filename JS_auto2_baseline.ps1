Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms

$wshell = New-Object -ComObject Wscript.shell
$shell=New-Object -ComObject shell.application
$mySI= (get-Process cmd -ErrorAction SilentlyContinue|Sort-Object StartTime -ea SilentlyContinue |Select-Object -first 1).SI
$checkcmd=((get-process cmd*  -ErrorAction SilentlyContinue)|Where-Object{$_.SI -eq $mySI}).HandleCount.count
$checkwinscp=((get-process winscp*)|Where-Object{$_.SI -eq $mySI}).HandleCount.count
$winv= ([System.Environment]::OSVersion.Version).Build

$pshid0= (Get-Process Powershell |Sort-Object StartTime -ea SilentlyContinue |Select-Object -last 1).id

function Set-WindowState {
	<#
	.LINK
	https://gist.github.com/Nora-Ballard/11240204
	#>

	[CmdletBinding(DefaultParameterSetName = 'InputObject')]
	param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
		[Object[]] $InputObject,

		[Parameter(Position = 1)]
		[ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE',
					 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED',
					 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
		[string] $State = 'SHOW'
	)

	Begin {
		$WindowStates = @{
			'FORCEMINIMIZE'		= 11
			'HIDE'				= 0
			'MAXIMIZE'			= 3
			'MINIMIZE'			= 6
			'RESTORE'			= 9
			'SHOW'				= 5
			'SHOWDEFAULT'		= 10
			'SHOWMAXIMIZED'		= 3
			'SHOWMINIMIZED'		= 2
			'SHOWMINNOACTIVE'	= 7
			'SHOWNA'			= 8
			'SHOWNOACTIVATE'	= 4
			'SHOWNORMAL'		= 1
		}

		$Win32ShowWindowAsync = Add-Type -MemberDefinition @'
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
'@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru

		if (!$global:MainWindowHandles) {
			$global:MainWindowHandles = @{ }
		}
	}

	Process {
		foreach ($process in $InputObject) {
			if ($process.MainWindowHandle -eq 0) {
				if ($global:MainWindowHandles.ContainsKey($process.Id)) {
					$handle = $global:MainWindowHandles[$process.Id]
				} else {
					Write-Error "Main Window handle is '0'"
					continue
				}
			} else {
				$handle = $process.MainWindowHandle
				$global:MainWindowHandles[$process.Id] = $handle
			}

			$Win32ShowWindowAsync::ShowWindowAsync($handle, $WindowStates[$State]) | Out-Null
			Write-Verbose ("Set Window State '{1} on '{0}'" -f $MainWindowHandle, $State)
		}
	}
}
##
if((get-process "cmd" -ea SilentlyContinue) -ne $Null){ 
$lastid=  (Get-Process cmd |Where-Object{$_.SI -eq $mySI}|sort-object StartTime -ea SilentlyContinue |select-object -last 1).id
 Get-Process -id $lastid  | Set-WindowState -State MINIMIZE
}
##>

$cSource = @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class Clicker
{
//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646270(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct INPUT
{ 
    public int        type; // 0 = INPUT_MOUSE,
                            // 1 = INPUT_KEYBOARD
                            // 2 = INPUT_HARDWARE
    public MOUSEINPUT mi;
}

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646273(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct MOUSEINPUT
{
    public int    dx ;
    public int    dy ;
    public int    mouseData ;
    public int    dwFlags;
    public int    time;
    public IntPtr dwExtraInfo;
}

//This covers most use cases although complex mice may have additional buttons
//There are additional constants you can use for those cases, see the msdn page
const int MOUSEEVENTF_MOVED      = 0x0001 ;
const int MOUSEEVENTF_LEFTDOWN   = 0x0002 ;
const int MOUSEEVENTF_LEFTUP     = 0x0004 ;
const int MOUSEEVENTF_RIGHTDOWN  = 0x0008 ;
const int MOUSEEVENTF_RIGHTUP    = 0x0010 ;
const int MOUSEEVENTF_MIDDLEDOWN = 0x0020 ;
const int MOUSEEVENTF_MIDDLEUP   = 0x0040 ;
const int MOUSEEVENTF_WHEEL      = 0x0080 ;
const int MOUSEEVENTF_XDOWN      = 0x0100 ;
const int MOUSEEVENTF_XUP        = 0x0200 ;
const int MOUSEEVENTF_ABSOLUTE   = 0x8000 ;

const int screen_length = 0x10000 ;

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646310(v=vs.85).aspx
[System.Runtime.InteropServices.DllImport("user32.dll")]
extern static uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

public static void LeftClickAtPoint(int x, int y)
{
    //Move the mouse
    INPUT[] input = new INPUT[3];
    input[0].mi.dx = x*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width);
    input[0].mi.dy = y*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);
    input[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;
    //Left mouse button down
    input[1].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    //Left mouse button up
    input[2].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    SendInput(3, input, Marshal.SizeOf(input[0]));
}
}
'@
Add-Type -TypeDefinition $cSource -ReferencedAssemblies System.Windows.Forms,System.Drawing

$dwidth=([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize).Width
$dhight=([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize).Height


################################# Check Powershell Windows and Get the latest Results ##########################################################

$pscount=(Get-Process Powershell).count

$vcount=(Get-Process video.UI -ErrorAction SilentlyContinue).count
$vrespond=(Get-Process video.UI -ErrorAction SilentlyContinue).Responding


if ($pscount -eq 1 -or ($vcount -eq 1 -and $vrespond -eq $true)){

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 

Start-Sleep -s 5

$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2_baseline.bat" 
$etime=(Get-Date).AddMinutes(5)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 

exit

}


if ($pscount -eq 2){
  
do{

start-sleep -s 30

$pshid=  (Get-Process Powershell -ErrorAction SilentlyContinue |sort-object StartTime |select-object -first 1).id |Where-Object{$_ -notmatch $pshid0}


start-sleep -s 30

$pshid2=  (Get-Process Powershell -ErrorAction SilentlyContinue |sort-object StartTime |select-object -first 1).id |Where-Object{$_ -notmatch $pshid0}


}until($pshid -ne $null -and $pshid2 -ne $null -and $pshid -eq $pshid2)


}

 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
$wshell.SendKeys("~")
start-sleep -s 10

 $wait_check=test-path C:\Jumpstart\batch\wait_pswindow.txt
 if($wait_check -eq $false){
   set-content C:\Jumpstart\batch\wait_pswindow.txt -value 0}


do{

  $waitt=get-content C:\Jumpstart\batch\wait_pswindow.txt
  $waitt=[Int64]($waitt|Out-String)
    
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
$wshell.SendKeys("^a")
start-sleep -s 2

 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
$wshell.SendKeys("^c")
start-sleep -s 2
$content=Get-Clipboard
$waitt++

if($content -match "Executing system preparation"){

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f'
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 
$etime=(Get-Date).AddMinutes(60)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 

exit

}

if(-not($content[-1] -match "Enter selected option \(or type Q to exit\)")){

$waitt2=$waitt+1

set-content  C:\Jumpstart\batch\wait_pswindow.txt -value $waitt2 -Force

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f'
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 
$etime=(Get-Date).AddMinutes(3)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 

exit
}

}until ($content[-1] -match "Enter selected option \(or type Q to exit\)"  -or $waitt -gt 120)


########### Get the latest Results #################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS" -f'
start-sleep -s 5 

$ini=""

 $wait_check=test-path C:\Jumpstart\batch\wait_pswindow.txt
 if($wait_check -eq $true){
  $waitt=get-content C:\Jumpstart\batch\wait_pswindow.txt
  $waitt=[Int64]($waitt|Out-String)}
  else{$waitt=0}

if($waitt -le 120){

<########### Check if baseline Environment ##################################
$env_check=test-path $env:USERPROFILE\Desktop\削除されたアプリ.html
if ($env_check -eq $false){
[System.Windows.Forms.MessageBox]::Show($this,"Not Baseline Enviroment, Please check!")
exit
}
####################>


########### Check if Need Prepare ##################################

$prep_flag="na"
if ($content -match "current status\: prep needed"){
$prep_flag="yes"

}

########### Check if Modern Standby ##################################
$modern_flag="yes"
if($content -match "Execute Standby"){
$modern_flag="na"
}


########### Check if Battery Mode ##################################
$Battery_flag="yes"
if(-not($content -match "Execute BatteryLife")){
$Battery_flag="na"
}

########### initialize ##############################

if($content -match "current\: OEM"){

New-Item -Name TempResult -Force -ItemType directory -Path C:\Jumpstart\batch\ |out-null

########### initialize the baseline needed testitems  ########
$result_csv=(get-childitem -path C:\Jumpstart\performance\results\OEM_*results.csv|select-object -Last 1).fullname
$result_content=import-csv $result_csv -Encoding UTF8
$result_items=$result_content.assessment|Sort-Object|Get-Unique

if($winv -ge 22000){
#$testitems=@("FastStartup","Standby","Edge","BatteryLife")
if($modern_flag -eq "yes" -and $Battery_flag -eq "yes"){$testitems=@("FastStartup","Edge","BatteryLife")}
if($modern_flag -eq "yes" -and $Battery_flag -eq "na"){$testitems=@("FastStartup","Edge")}
if($modern_flag -eq "na" -and $Battery_flag -eq "yes"){$testitems=@("FastStartup","Standby","Edge","BatteryLife")}
if($modern_flag -eq "na" -and $Battery_flag -eq "na"){$testitems=@("FastStartup","Standby","Edge")}

}
else{
if($modern_flag -eq "yes" -and $Battery_flag -eq "yes"){$testitems=@("FastStartup","BatteryLife")}
if($modern_flag -eq "yes" -and $Battery_flag -eq "na"){$testitems=@("FastStartup")}
if($modern_flag -eq "na" -and $Battery_flag -eq "yes"){$testitems=@("FastStartup","Standby","BatteryLife")}
if($modern_flag -eq "na" -and $Battery_flag -eq "na"){$testitems=@("FastStartup","Standby")}
}


foreach($testitem in $testitems){
   $setfile=$testitem+"_baseline.txt"
   $setvaluefile=$testitem+"_baseline_value.txt" 

if($testitem -eq "FastStartup" -and $result_items -like "*$testitem*"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem -and $_."testcase" -eq "Total"})."metricvalue"| Measure-Object -Minimum).Minimum
$item_status=($result_content|Where-Object {$_."assessment" -eq $testitem -and $_."testcase" -eq "Total" -and $_."metricvalue" -eq $item_value}).status
$ini_check = test-path C:\Jumpstart\batch\$setfile

if($item_status -ne "Pass" -and $ini_check -eq $false){
 set-content  C:\Jumpstart\batch\$setfile -value 0 -Force
 new-item  -path C:\Jumpstart\batch\$setvaluefile -Force |out-null
}
}


if($testitem -eq "BatteryLife"  -and $result_items -like "*$testitem*"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem})."metricvalue"| Measure-Object -Maximum).Maximum
$item_status=($result_content|Where-Object {$_."assessment" -eq $testitem  -and $_."metricvalue" -eq $item_value}).status

#####
$ini_check = test-path C:\Jumpstart\batch\$setfile
if($item_status -ne "Pass" -and $ini_check -eq $false){
 set-content  C:\Jumpstart\batch\$setfile -value 0 -Force
 new-item  -path C:\Jumpstart\batch\$setvaluefile -Force |out-null
}
###>
}


if(($testitem -eq "Edge"-or $testitem -eq "Standby")  -and $result_items -like "*$testitem*"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem})."metricvalue" | Measure-Object -Minimum).Minimum
$item_status=($result_content|Where-Object {$_."assessment" -eq $testitem  -and $_."metricvalue" -eq $item_value}).status

$ini_check = test-path C:\Jumpstart\batch\$setfile

if($item_status -ne "Pass" -and $ini_check -eq $false){
   $setfile=$testitem+"_baseline.txt"
 set-content  C:\Jumpstart\batch\$setfile -value 0 -Force
 new-item  -path C:\Jumpstart\batch\$setvaluefile -Force |out-null
}
}

}


########### Rename as baseline  ##################################

 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 
 start-sleep -s 2

$wshell.SendKeys("T")
start-sleep -s 5
$wshell.SendKeys("~")
start-sleep -s 10
$wshell.SendKeys("Baseline")
start-sleep -s 5
$wshell.SendKeys("~")
start-sleep -s 10
$wshell.SendKeys("~")

 $ini= "ok"

}

########### check after baseline testing ##############################

if($content -match "current\: Baseline"){

$result_csv=(get-childitem -path C:\Jumpstart\performance\results\OEM_*results.csv|select-object -Last 1).fullname
$result_content=import-csv $result_csv -Encoding UTF8
$oem_result_paths=$result_content.xmlpath|ForEach-Object{($_.split("\"))[-2]}|Sort-Object|Get-Unique

get-childitem -path C:\Jumpstart\performance\results\JobResults* | ForEach-Object{

$axe=$_.fullname+"\AxeLog.txt"
$bl_folder_content=get-content -path $axe
if($bl_folder_content -match "baseline"){
$_.fullname
$lastesetfolder=$lastesetfolder+@($_.fullname)
}
}
$latest_bl= $lastesetfolder|Sort-Object|select-object -last 1

get-childitem -path C:\Jumpstart\performance\results\JobResults* | ForEach-Object{


if($_.name -notin $oem_result_paths -and $latest_bl -notmatch $_.name ){
move-item $_.fullname  C:\Jumpstart\batch\TempResult\  -Force
}

}

########### generate-tested data ##################################

 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null

$wshell.SendKeys("a")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2

$waitt2=0
do{

start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
$wshell.SendKeys("^a")

start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
$wshell.SendKeys("^c")

start-sleep -s 2
$content2=Get-Clipboard
$waitt2++
}until ($content2 -match "Enter selected option \(or type Q to exit\)" -or $waitt2 -gt 20)

if($waitt2 -gt 20){

 [System.Windows.Forms.MessageBox]::Show($this,"Data is not Ready after 100 secs!")

exit 
} 


if($content2 -match "Current results path \:" -and $content2 -match "Test (\w+)\: Baseline"){

start-sleep -s 2
$wshell.SendKeys("a")

start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5

do{
   start-sleep -s 5
  [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
   Get-Process -id $pshid2 | Set-WindowState -State MAXIMIZE
     [Clicker]::LeftClickAtPoint(1,1)
     start-sleep -s 2
     $wshell.SendKeys("E")
     start-sleep -s 2
     $wshell.SendKeys("S")
     start-sleep -s 2
     $wshell.SendKeys("~")
     start-sleep -s 5
     $content3=Get-Clipboard
}until($content3[-1] -match "Press any key to continue . . .")

$wshell.SendKeys("~")
 start-sleep -s 5

 
 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null

start-sleep -s 2
$wshell.SendKeys("g")
start-sleep -s 2
$wshell.SendKeys("~")
 
 start-sleep -s 2

  [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 
 Get-Process -id $pshid2 | Set-WindowState -State MAXIMIZE
 [Clicker]::LeftClickAtPoint(1,1)
 start-sleep -s 2
$wshell.SendKeys("E")
start-sleep -s 2
$wshell.SendKeys("S")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5
$content=Get-Clipboard

if($content -match "\[O\]verwrite or \[A\]ppend to CSV file\?"){

start-sleep -s 2

$wshell.SendKeys("O")

$wshell.SendKeys("~")

}

start-sleep -s 2

do{
     start-sleep -s 5
     [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
     Get-Process -id $pshid2 | Set-WindowState -State MAXIMIZE
     [Clicker]::LeftClickAtPoint(1,1)
     start-sleep -s 2
     $wshell.SendKeys("E")
     start-sleep -s 2
     $wshell.SendKeys("S")
     start-sleep -s 2
     $wshell.SendKeys("~")
     start-sleep -s 5
     $content=Get-Clipboard
}until($content -match "Press any key to continue . . .")


$timenow=get-date -Format yyMMdd_HHmm
set-content -path "C:\Jumpstart\batch\tempResult\tempResult_$timenow.txt" -value $content

start-sleep -s 2
$wshell.SendKeys("~")


start-sleep -s 2
$wshell.SendKeys("U")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5
$wshell.SendKeys("q")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5

###########check the baseline needed testitems  ########
$result_csv=(get-childitem -path C:\Jumpstart\performance\results\Baseline_*results.csv|select-object -Last 1).fullname
$result_content=import-csv $result_csv -Encoding UTF8
$result_items=$result_content.assessment|Sort-Object|Get-Unique


if($winv -ge 22000){
#$testitems=@("FastStartup","Standby","Edge","BatteryLife")
if($modern_flag -eq "yes" -and $Battery_flag -eq "yes"){$testitems=@("FastStartup","Edge","BatteryLife")}
if($modern_flag -eq "yes" -and $Battery_flag -eq "na"){$testitems=@("FastStartup","Edge")}
if($modern_flag -eq "na" -and $Battery_flag -eq "yes"){$testitems=@("FastStartup","Standby","Edge","BatteryLife")}
if($modern_flag -eq "na" -and $Battery_flag -eq "na"){$testitems=@("FastStartup","Standby","Edge")}

}
else{
if($modern_flag -eq "yes" -and $Battery_flag -eq "yes"){$testitems=@("FastStartup","BatteryLife")}
if($modern_flag -eq "yes" -and $Battery_flag -eq "na"){$testitems=@("FastStartup")}
if($modern_flag -eq "na" -and $Battery_flag -eq "yes"){$testitems=@("FastStartup","Standby","BatteryLife")}
if($modern_flag -eq "na" -and $Battery_flag -eq "na"){$testitems=@("FastStartup","Standby")}
}


foreach($testitem in $testitems){
$setfile=$testitem+"_baseline.txt"
$setvaluefile=$testitem+"_baseline_value.txt"
 $retestpath=test-path C:\Jumpstart\batch\$setfile
 if  ($retestpath -eq $true){

if($testitem -eq "FastStartup" -and $result_items -like "*$testitem*"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem -and $_."testcase" -eq "Total"})."metricvalue"| Measure-Object -Minimum).Minimum
$item_status=($result_content|Where-Object {$_."assessment" -eq $testitem -and $_."testcase" -eq "Total" -and $_."metricvalue" -eq $item_value}).status

$ini_check = test-path C:\Jumpstart\batch\$setfile
if ($ini_check -eq $true){$bl_count=get-content C:\Jumpstart\batch\$setfile}

$bl_value=(get-content -Path (get-childitem C:\Jumpstart\batch\TempResult\tempResult*|select-object -last 1).fullname)

foreach ($bll in $bl_value){
if($bll -match $testitem -and $bll -match "Total" ){$linec=$bl_value.indexof($bll) 
break}
}


$valuea=($bl_value[$linec]).split(" ")|foreach-object{
if($_ -match "\d{1,}\.\d{1,}"){
$valueX=$_
}
}

  add-content -path C:\Jumpstart\batch\$setvaluefile -value $valueX

$newv= ("""1"",""2"",""3"",""4"",""5""") 
 $va1=get-content "C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.csv"|ForEach-Object{
 $newv=$newv+"`n"+$_
 }

$newv|Set-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv -Encoding UTF8
 
$table_content=import-csv "C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv" -Encoding UTF8
$value0=$table_content[6].3
if($value0 -eq $null){$table_content[6].3=$valueX}
if($value0 -ne $null -and $valueX -lt $value0){$table_content[6].3=$valueX}


 if($item_status -eq "Pass" ){$table_content[6].4="Pass"}
$table_content|export-csv C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt -Encoding UTF8 -NoTypeInformation
(Get-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt | Select-Object -Skip 1) | Set-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.csv
Remove-Item C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt -Force
Remove-Item C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv -Force 

  if($item_status -eq "Pass" -or $bl_count -ge 5){ remove-item  C:\Jumpstart\batch\$setfile -Force}

}


if($testitem -eq "BatteryLife"  -and $result_items -like "*$testitem*"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem})."metricvalue"| Measure-Object -Maximum).Maximum
$item_status=($result_content|Where-Object {$_."assessment" -eq $testitem  -and $_."metricvalue" -eq $item_value}).status

$ini_check = test-path C:\Jumpstart\batch\$setfile
if ($ini_check -eq $true){$bl_count=get-content C:\Jumpstart\batch\$setfile }

$bl_value=(get-content -Path (get-childitem C:\Jumpstart\batch\TempResult\tempResult*|select-object -last 1).fullname)
foreach ($bll in $bl_value){
if($bll -match $testitem -and $bll -match "FullDrain" ){$linec=$bl_value.indexof($bll)
break}
}
$valuea=($bl_value[$linec]).split(" ")|ForEach-Object{
if($_ -match "\d{3,}"){
$valueX=$_
}
}

  add-content -path C:\Jumpstart\batch\$setvaluefile -value $valueX

$newv= ("""1"",""2"",""3"",""4"",""5""") 
 $va1=get-content "C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.csv"|ForEach-Object{
 $newv=$newv+"`n"+$_
 }

$newv|Set-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv -Encoding UTF8
 
$table_content=import-csv "C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv" -Encoding UTF8

$value0=$table_content[8].3
if($value0 -eq $null){$table_content[8].3=$valueX}
if($value0 -ne $null -and $valueX -gt $value0){$table_content[8].3=$valueX}

if($item_status -eq "Pass"){$table_content[8].4="Pass"}
$table_content|export-csv C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt -Encoding UTF8 -NoTypeInformation
(Get-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt | Select-Object -Skip 1) | Set-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.csv
Remove-Item C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt -Force
Remove-Item C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv -Force 

 if($item_status -eq "Pass" -or $bl_count -ge 3){remove-item  C:\Jumpstart\batch\$setfile -Force}

}

if(($testitem -eq "Edge"-or $testitem -eq "Standby")  -and $result_items -like "*$testitem*"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem})."metricvalue" | Measure-Object -Minimum).Minimum
$item_status=($result_content|Where-Object {$_."assessment" -eq $testitem  -and $_."metricvalue" -eq $item_value}).status

$ini_check = test-path C:\Jumpstart\batch\$setfile
if ($ini_check -eq $true){$bl_count=get-content C:\Jumpstart\batch\$setfile }

$bl_value=(get-content -Path (get-childitem C:\Jumpstart\batch\TempResult\tempResult*|select-object -last 1).fullname)
foreach ($bll in $bl_value){
if($bll -match $testitem -and ($bll -match "Total" -or $bll -match "LargestContentfulPaint") ){$linec=$bl_value.indexof($bll)
break}
}
$valuea=($bl_value[$linec]).split(" ")|ForEach-Object{
if($_ -match "\d{1,}\.\d{1,}"){
$valueX=$_
}
}

  add-content -path C:\Jumpstart\batch\$setvaluefile -value $valueX

$newv= ("""1"",""2"",""3"",""4"",""5""") 
 $va1=get-content "C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.csv"|ForEach-Object{
 $newv=$newv+"`n"+$_
 }

$newv|Set-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv -Encoding UTF8
 
$table_content=import-csv "C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv" -Encoding UTF8

if($testitem -eq "Edge"){

$value0=$table_content[7].3
if($value0 -eq $null){$table_content[7].3=$valueX}
if($value0 -ne $null -and $valueX -lt $value0){$table_content[7].3=$valueX}

 if($item_status -eq "Pass"){$table_content[7].4="Pass"}
}
if($testitem -eq  "Standby"){
$value0=$table_content[9].3
if($value0 -eq $null){$table_content[9].3=$valueX}
if($value0 -ne $null -and $valueX -lt $value0){$table_content[9].3=$valueX}
 if($item_status -eq "Pass"){$table_content[9].4="Pass"}
}

$table_content|export-csv C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt -Encoding UTF8 -NoTypeInformation
(Get-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt | Select-Object -Skip 1) | Set-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.csv
Remove-Item C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt -Force
Remove-Item C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.csv -Force 


  if($item_status -eq "Pass" -or $bl_count -ge 5){remove-item  C:\Jumpstart\batch\$setfile -Force}

}

}
}

$ini="ok"

}



else{

 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2
$wshell.SendKeys("q")

start-sleep -s 2
$wshell.SendKeys("~")

$ini="ok"
}


$testcheck=(get-childitem "C:\Jumpstart\batch\*baseline.txt").count 

if($testcheck -eq 0){

#######Finish##########


######delete Schedule #####
start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS" -f' 
Start-Sleep -s 5
start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp"  -f' 
Start-Sleep -s 5
     ######Show Messages #####
  
move-item  C:\Jumpstart\batch\TempResult\JobResults* C:\Jumpstart\performance\results\  -Force

 remove-item C:\Jumpstart\batch\wait_pswindow.txt -Force

 Get-Process -id $pshid2 | Set-WindowState -State MINIMIZE
 #[System.Windows.Forms.MessageBox]::Show($this,"Baseline Test Complete!")
  #exit
   ######  Report generate #####
   invoke-expression -Command C:\Jumpstart\batch\get_report.ps1

}
}

}

else{


 $datenow=get-date -Format yyMMdd_HHmm
 move-item C:\Jumpstart\batch\wait_pswindow.txt C:\Jumpstart\batch\wait_pswindow_$datenow.txt -Force
 [System.Windows.Forms.MessageBox]::Show($this,"No Program Ready has been found after 10 mins!")

exit 
} 

#################################Check Enviroment settings ####################################################


$testcheck=(get-childitem "C:\Jumpstart\batch\*baseline.txt").count 
       
if ( $ini -eq "ok" -and $testcheck -ne 0){

################################# Check Hajime and turn off re-open ####################################################

$rcdid= (Get-Process RcdSettings -ErrorAction SilentlyContinue).id
start-sleep -seconds 2

if($rcdid -ne $null){

 [Microsoft.VisualBasic.interaction]::AppActivate($rcdid)|out-null
  
 Get-Process -id $rcdid | Set-WindowState -State MAXIMIZE
  [Clicker]::LeftClickAtPoint(1,1)

start-sleep -seconds 2
$wshell.SendKeys('{TAB}')
start-sleep -seconds 2
$wshell.SendKeys('{Enter}')
start-sleep -seconds 2
$wshell.SendKeys('{TAB}')
start-sleep -seconds 2
$wshell.SendKeys('{Enter}')
start-sleep -seconds 2

$wshell.SendKeys('{TAB}')
start-sleep -seconds 2
$wshell.SendKeys('{TAB}')
start-sleep -seconds 2
$wshell.SendKeys('{TAB}')
start-sleep -seconds 2
$wshell.SendKeys(' ')
start-sleep -seconds 2

 [Microsoft.VisualBasic.interaction]::AppActivate($rcdid)|out-null
$wshell.SendKeys('%" "')
Stop-Process -name RcdSettings
}
#################################  Power Settings ####################################################


$powerset=get-content C:\Jumpstart\batch\Power.txt
if($powerset -match 2){
 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
start-sleep -s 2

$wshell.SendKeys("c")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2

$wshell.SendKeys("^a")
start-sleep -s 2
$wshell.SendKeys("^c")
start-sleep -s 2
$content=Get-Clipboard

if ($content -match "\[current status\: on"){

do{
start-sleep -s 2
$wshell.SendKeys("J")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2

$wshell.SendKeys("^a")
start-sleep -s 2
$wshell.SendKeys("^c")
start-sleep -s 2

} until ($content -match "\[current status\: off")

}

$wshell.SendKeys("q")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
set-content C:\Jumpstart\batch\Power.txt -value "3"

}

###>

################################# Need Prepare  ####################################################

if ($prep_flag -match "yes"){

#############################Schedule Setup after login ####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2_baseline.bat" 

$trigger = New-JobTrigger -AtLogOn -RandomDelay 00:05:00

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS" 
#Register-ScheduledTask JS_Auto -Action $Sta -Trigger $Stt -Settings $Stset -Force


start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
$etime=(Get-Date).AddMinutes(60)
$trigger2 = New-ScheduledTaskTrigger -Once -At $etime 

Register-ScheduledTask -Action $action -Trigger $trigger2 -Settings $Stset -Force -TaskName "Auto_JS_temp" 



############################# AutoLogon ####################################


 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2

$wshell.SendKeys("S")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("1111")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5


############################# Start to Preapre ####################################

 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2
 
$wshell.SendKeys("P")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("Y")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("N")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5


exit


}


else{

######delete Schedule #####
start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS" -f' 
Start-Sleep -s 5
start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp"  -f' 
Start-Sleep -s 5

}


################################# start baseline test and record test rounds ####################################################

$testcheck=(get-childitem "C:\Jumpstart\batch\*baseline.txt").count

if($testcheck -gt 0){

$testitems2=@("FastStartup_baseline.txt","Standby_baseline.txt","Edge_baseline.txt","BatteryLife_baseline.txt")

(get-childitem "C:\Jumpstart\batch\*baseline.txt").name|Sort-Object{$testitems2.IndexOf($_)} |ForEach-Object{
$testgo=$_
$itemfile="C:\Jumpstart\batch\$_"
$testcount=get-content -path $itemfile

#########run FSU########

if($testgo -match "FastStartup" -and $testcount -le 5){


#############################Schedule Setup after FSU login ####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS" -f' 
Start-Sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2_baseline.bat" 

$trigger = New-JobTrigger -AtLogOn -RandomDelay 00:05:00

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS" 
#Register-ScheduledTask JS_Auto -Action $Sta -Trigger $Stt -Settings $Stset -Force


############################# AutoLogon ####################################3


 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2

$wshell.SendKeys("S")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("1111")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5

$testcount=[int64]$testcount+1
set-content  $itemfile -value $testcount -Force

start-sleep -s 2
[Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
start-sleep -s 2
$wshell.SendKeys("2")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("g")
start-sleep -s 2

exit


}


#########run standby########

if($testgo -match "Standby" -and $testcount -le 5){



#############################Schedule Setup  2 min later####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2_baseline.bat" 
$etime=(Get-Date).AddMinutes(3)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 

 
$testcount=[int64]$testcount+1
set-content  $itemfile -value $testcount -Force

start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2
 
$wshell.SendKeys("4")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2

exit

}


#########run edge########

if($testgo -match "Edge" -and $testcount -le 5){


#############################Schedule Setup  2 min later####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2_baseline.bat" 
$etime=(Get-Date).AddMinutes(3)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 

$testcount=[int64]$testcount+1
set-content  $itemfile -value $testcount -Force
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2
 
$wshell.SendKeys("3")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2

exit

}

#########run BatteryLife########

if($testgo -match "BatteryLife" -and $testcount -le 3){

#############################Schedule Setup  2 min later####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5

$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2_baseline.bat" 
$etime=(Get-Date).AddMinutes(60)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 

$testcount=[int64]$testcount+1
set-content  $itemfile -value $testcount -Force
start-sleep -s 2
 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2
 
$wshell.SendKeys("5")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5

do{
start-sleep -s 5
$bid= (get-process energy*).id

}until($bid -ne $Null)

 start-sleep -s 10
 [Microsoft.VisualBasic.interaction]::AppActivate($bid)|out-null
 $wshell.SendKeys("~")
start-sleep -s 2

exit

}

}
}


}
