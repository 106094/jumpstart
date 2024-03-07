Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 #$checkdouble=(get-process cmd*).HandleCount.count
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms

$wshell = New-Object -ComObject Wscript.shell
$shell=New-Object -ComObject shell.application
$mySI= (get-Process cmd -ErrorAction SilentlyContinue|sort-object StartTime -ea SilentlyContinue |select-object -first 1).SI
#$checkcmd=((get-process cmd*  -ErrorAction SilentlyContinue)|where-object{$_.SI -eq $mySI}).HandleCount.count
$winv= ([System.Environment]::OSVersion.Version).Build
#$testWin10=test-path C:\Jumpstart\performance\Win10\
#$testWin11=test-path C:\Jumpstart\performance\Win11\
$testfolder=test-path "C:\Jumpstart\performance\results\"

##### copy JS tool to local path ####

if(!$testfolder){
 if($winv -ge 22000){$A1= (get-childitem -path C:\Jumpstart\performance\Win11\*.zip).fullname }
 else{$A1= (get-childitem -path C:\Jumpstart\performance\Win10\*.zip).fullname }
 write-output "waiting 1～3 min for JS tool file extracting to local path"
  #Expand-Archive -LiteralPath $A1  -DestinationPath C:\Jumpstart\performance\
   $shell.NameSpace("C:\Jumpstart\performance\").copyhere($shell.NameSpace($A1).Items(),4)

  do{
  start-sleep -s 5
  $test_ass=test-path C:\Jumpstart\performance\results\
  }until($test_ass -eq $true)
  start-sleep -s 30

  remove-item C:\Jumpstart\performance\Win10 -Recurse -Force -ea SilentlyContinue
   remove-item C:\Jumpstart\performance\Win11 -Recurse -Force -ea SilentlyContinue

$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 
$etime=(Get-Date).AddMinutes(2)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 

 exit

}

$pshid0= (Get-Process Powershell |sort-object StartTime -ea SilentlyContinue |select-object -last 1).id

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
if((get-process "cmd" -ea SilentlyContinue)){ 
$lastid=  (Get-Process cmd |where-object{$_.SI -eq $mySI}|sort-object StartTime -ea SilentlyContinue |select-object -last 1).id
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

#$dwidth=([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize).Width
#$dhight=([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize).Height


################################# Check Powershell Windows and Get the latest Results ##########################################################

$pscount=(Get-Process Powershell).count

$vcount=(Get-Process video.UI -ea SilentlyContinue).count
$vrespond=(Get-Process video.UI -ea SilentlyContinue).Responding

if ($pscount -eq 1 -or ($vcount -eq 1 -and $vrespond -eq $true)){

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 
$etime=(Get-Date).AddMinutes(5)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 

exit

}


if ($pscount -eq 2){
  
do{

start-sleep -s 30

$pshid=  (Get-Process Powershell -ErrorAction SilentlyContinue |sort-object StartTime |select-object -first 1).id |where-object{$_ -notmatch $pshid0}


start-sleep -s 30

$pshid2=  (Get-Process Powershell -ErrorAction SilentlyContinue |sort-object StartTime |select-object -first 1).id |where-object{$_ -notmatch $pshid0}


}until($pshid -and $pshid2 -and $pshid -eq $pshid2)


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

########### Check if Need Prepare ##################################

$prep_flag="na"
if ($content -match "current status\: prep needed"){
$prep_flag="yes"

}

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

if($content2 -match "Current results path \:"){


$wshell.SendKeys("g")
start-sleep -s 2
$wshell.SendKeys("~")
 
 start-sleep -s 10

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

set-content -path "C:\Jumpstart\batch\tempResult.txt" -value $content

start-sleep -s 2
$wshell.SendKeys("~")

start-sleep -s 2
$wshell.SendKeys("q")

start-sleep -s 2
$wshell.SendKeys("~")

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
}

else{
  
 $datenow=get-date -Format yyMMdd_HHmm
 move-item C:\Jumpstart\batch\wait_pswindow.txt C:\Jumpstart\batch\wait_pswindow_$datenow.txt -Force
 [System.Windows.Forms.MessageBox]::Show($this,"No Program Ready has been found after 10 mins!")

exit}  ####### For Prepare use ############3


################################# Check History Results ####################################################

$csvexist=get-childitem -path "C:\Jumpstart\performance\results\OEM*results.csv"|select-object -Last 1

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


$runflag="waitcheck"

remove-item "C:\Jumpstart\batch\runflag.txt" -force -ErrorAction SilentlyContinue

if ($csvexist.count -eq 0){
$runflag="FastStartup"
set-content -path "C:\Jumpstart\batch\runflag.txt" -value $runflag

}


if ($csvexist.count -ne 0){ 

foreach($testitem in $testitems){

$newcsv=$csvexist.fullname 
$csv_content=import-csv $newcsv -Encoding UTF8

if($testitem -eq "FastStartup"){$item_content=($csv_content|where-object{$_."assessment" -eq $testitem -and $_."testcase" -eq "Total"})."status"}
else{$item_content=($csv_content|where-object{$_."assessment" -eq $testitem})."status"}


if ($item_content.Count -gt 0){

if("Pass" -notin $item_content){

if($testitem -eq "FastStartup"){
$failc=($csv_content|where-object{$_."assessment" -eq $testitem -and $_."testcase" -eq "Total" -and $_."status" -match "baseline"}).count
$failc2=($csv_content|where-object{$_."assessment" -eq $testitem -and $_."testcase" -eq "Total" -and $_."status" -match "fail"}).count

}
else{
 $failc=($csv_content|where-object{$_."assessment" -eq $testitem -and $_."status" -match "baseline"}).count
 $failc2=($csv_content|where-object{$_."assessment" -eq $testitem -and $_."status" -match "fail"}).count
 }
 
 if( ($testitem -ne "BatteryLife" -and $failc -lt 5) -or ($testitem -eq "BatteryLife" -and $failc -lt 3)){

write-output  "$testitem,$failc"
 $runflag=$testitem
 
set-content -path "C:\Jumpstart\batch\runflag.txt" -value $runflag
break
}

### Fail Maximun 2 ###########
 if($failc2 -gt 0 -and $failc2 -lt 2){

 write-output "$testitem,$failc"
 $runflag=$testitem
 
set-content -path "C:\Jumpstart\batch\runflag.txt" -value $runflag
break
}

 }
  }
if ($item_content.Count -eq 0){
 $runflag=$testitem
set-content -path "C:\Jumpstart\batch\runflag.txt" -value $runflag

break

}

   }
}


 $passflagr=test-path "C:\Jumpstart\batch\runflag.txt"
  if($passflagr -eq $false){

    ######delete Schedule #####
    start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS" -f' 
    start-sleep -s 5
     ######Show Messages #####

 remove-item C:\Jumpstart\batch\wait_pswindow.txt -Force
 Get-Process -id $pshid2 | Set-WindowState -State MINIMIZE
 #[System.Windows.Forms.MessageBox]::Show($this,"OEM Test Complete!")

   ######  Report generate #####
 invoke-expression -Command C:\Jumpstart\batch\get_report.ps1

  }


       
if ( $ini -eq "ok" -and $passflagr -eq $true){

 
$testgo= get-content -path "C:\Jumpstart\batch\runflag.txt"
################################# Check Hajime and turn off re-open ####################################################


$rcdid= (Get-Process RcdSettings).id
start-sleep -seconds 2

if($rcdid){

 [Microsoft.VisualBasic.interaction]::AppActivate($rcdid)|out-null
  Get-Process -id $rcdid | Set-WindowState -State MAXIMIZE
  [Clicker]::LeftClickAtPoint(1,1)

start-sleep -seconds 2
$wshell.SendKeys('{TAB}')
start-sleep -seconds 1
$wshell.SendKeys('{Enter}')
start-sleep -seconds 1
$wshell.SendKeys('{TAB}')
start-sleep -seconds 1
$wshell.SendKeys('{Enter}')
start-sleep -seconds 1

$wshell.SendKeys('{TAB}')
start-sleep -seconds 2
$wshell.SendKeys('{TAB}')
start-sleep -seconds 1
$wshell.SendKeys('{TAB}')
start-sleep -seconds 1
$wshell.SendKeys(' ')
start-sleep -seconds 1

 [Microsoft.VisualBasic.interaction]::AppActivate($rcdid)|out-null
$wshell.SendKeys('%" "')
Stop-Process -name RcdSettings
}



#################################  Power Settings ####################################################



$powerset=get-content C:\Jumpstart\batch\Power.txt
if($powerset -match 1){
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
set-content C:\Jumpstart\batch\Power.txt -value "2"

}



################################# Need Prepare  ####################################################

if ($prep_flag -match "yes"){


#############################Schedule Setup after login or wait Prepare w/O reboot ####################################

#start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS" -f' 

$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 

$trigger = New-JobTrigger -AtLogOn -RandomDelay 00:05:00

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS" 
#Register-ScheduledTask JS_Auto -Action $Sta -Trigger $Stt -Settings $Stset -Force



start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 
$etime=(Get-Date).AddMinutes(60)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 
$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 


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
start-sleep -s 5
start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp"  -f' 
start-sleep -s 5
}



#########run FSU########

if($testgo -match "FastStartup"){



#############################Schedule Setup after login  ####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 

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

[Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null

$wshell.SendKeys("2")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2
$wshell.SendKeys("g")
start-sleep -s 2

exit



}


#########run standby########


if($testgo -match "Standby"){



#############################Schedule Setup  2 min later####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 
$etime=(Get-Date).AddMinutes(3)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 


 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2
 
$wshell.SendKeys("4")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2

exit


}

#########run edge########

if($testgo -match "Edge"){


#############################Schedule Setup  2 min later####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 
$etime=(Get-Date).AddMinutes(3)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 


 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2
 
$wshell.SendKeys("3")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 2

exit

}

#########run BatteryLife########

if($testgo -match "BatteryLife"){


#############################Schedule Setup  10 min later####################################

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_JS_temp" -f' 
start-sleep -s 5
$action = New-ScheduledTaskAction -Execute "C:\Jumpstart\batch\JS_auto2.bat" 
$etime=(Get-Date).AddMinutes(60)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 

$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_JS_temp" 


 [Microsoft.VisualBasic.interaction]::AppActivate($pshid2)|out-null
 start-sleep -s 2
 
$wshell.SendKeys("5")
start-sleep -s 2
$wshell.SendKeys("~")
start-sleep -s 5

do{
start-sleep -s 5
$bid= (get-process energy*).id

}until($bid)

 start-sleep -s 10
 [Microsoft.VisualBasic.interaction]::AppActivate($bid)|out-null
 $wshell.SendKeys("~")
start-sleep -s 2

exit

}

}