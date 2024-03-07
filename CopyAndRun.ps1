Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
Add-Type -AssemblyName System.Windows.Forms
$shell=New-Object -ComObject shell.application
$ping = New-Object System.Net.NetworkInformation.Ping

if($PSScriptRoot.length -eq 0){
    $scriptRoot="D:\0227\_JS_8.6P_test"
    }
    else{
    $scriptRoot=$PSScriptRoot
    }
 

#$desktophtml="$env:USERPROFILE/desktop/*.html"
$tagfloder="C:\Jumpstart\Tag"
$logfloder="C:\Jumpstart\log"

#region create folder
    if(!(test-path $tagfloder)){
    new-item -ItemType Directory -Path $tagfloder -Force |Out-Null
}
if(!(test-path $logfloder)){
    new-item -ItemType Directory -Path $logfloder -Force |Out-Null
}

#region link test.net
 [System.Windows.Forms.MessageBox]::Show($this,"請以wifi連接測網,連線後按下OK") |out-null
 start-sleep -s 20
 $timex1=get-date
do{
    start-sleep -s 10
    try{
    $checkconnect=($ping.Send("www.google.com", 1000)).Status
    }
    catch{
    [System.Windows.Forms.MessageBox]::Show($this,"請確認wifi連接測網") |out-null
    }
   

    $timepassed=(New-TimeSpan -start  $timex1 -End (get-date)).TotalSeconds
} until($checkconnect -match "Success" -or $timepassed -ge 100)

if($timepassed -ge 100){
[System.Windows.Forms.MessageBox]::Show($this,"連接測網失敗") |out-null
exit
}
#region change DNS

$getinfo=ipconfig
$checkline=($getinfo -match "192.168.")[0]

$linenu=$getinfo.IndexOf($checkline)
($getinfo|Select-Object -Skip $linenu|Select-Object -First 1) -match "\d{1,}\.\d{1,}\.\d{1,}\.\d{1,}" |Out-Null
$ipout=$matches[0]

Get-NetIPAddress|ForEach-Object{
    if($_.IPAddress -match $ipout -and $_.AddressFamily -eq "IPv4"){
    $adtname=$_.InterfaceAlias 
    }
}

#$dnsconfigtxt="C:\Jumpstart\tag\dnsconfig_before.txt"
#Get-DnsClientServerAddress -InterfaceAlias $adtname|Where-object{$_.AddressFamily -eq "2"}|Out-String|set-content $dnsconfigtxt

$wificonfig="C:\Jumpstart\tag\wifiname.txt"
$adtname|set-content $wificonfig

#$adtname
$dnsip="127.0.0.1"
Set-DnsClientServerAddress -InterfaceAlias $adtname -ServerAddresses $dnsip
Clear-DnsClientCache

        $ping = New-Object System.Net.NetworkInformation.Ping
        $timex1=get-date
        do{
            start-sleep -s 10
            $checkconnect=($ping.Send("www.google.com", 1000)).Status
            $timepassed=(New-TimeSpan -start  $timex1 -End (get-date)).TotalSeconds
        } until(!($checkconnect -match "Success") -or $timepassed -ge 100)
        if($timepassed -ge 100){
        [System.Windows.Forms.MessageBox]::Show($this,"測網斷線失敗") |out-null
        exit
        }


#endregion

try{
#region cmd and windows termial settings for windows 11 ###

$Version = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\'

if($Version.CurrentBuildNumber -ge 22000){

## wt json file ##

$wtcu=(get-process WindowsTerminal -ErrorAction SilentlyContinue).Id
start-process wt
start-sleep -s 5
Stop-Process -id (get-process WindowsTerminal|Where-Object{$_.id -notin $wtcu}).Id

$wtsettings=get-content "$env:USERPROFILE\AppData\Local\Packages\Microsoft.WindowsTerminal*\localState\settings.json"

function ConvertTo-Hashtable {
    [CmdletBinding()]
    [OutputType('hashtable')]
    param (
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    process {
        ## Return null if the input is null. This can happen when calling the function
        ## recursively and a property is null
        if ($null -eq $InputObject) {
            return $null
        }

        ## Check if the input is an array or collection. If so, we also need to convert
        ## those types into hash tables as well. This function will convert all child
        ## objects into hash tables (if applicable)
        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            $collection = @(
                foreach ($object in $InputObject) {
                    ConvertTo-Hashtable -InputObject $object
                }
            )

            ## Return the array but don't enumerate it because the object may be pretty complex
            Write-Output -NoEnumerate $collection
        } elseif ($InputObject -is [psobject]) { ## If the object has properties that need enumeration
            ## Convert it to its own hash table and return it
            $hash = @{}
            foreach ($property in $InputObject.PSObject.Properties) {
                $hash[$property.Name] = ConvertTo-Hashtable -InputObject $property.Value
            }
            $hash
        } else {
            ## If the object isn't an array, collection, or other object, it's already a hash table
            ## So just return it.
            $InputObject
        }
    }
}

$wshash=$wtsettings | ConvertFrom-Json | ConvertTo-HashTable
$guidcmd=($wshash.profiles.list|Where-Object{$_.name -match "command"}).guid

$wtsettings|ForEach-Object{
$wtset=$_
if($wtset -match "defaultProfile"){
    $wtset= """defaultProfile"": ""$guidcmd"","
}
$wtset

}|set-content "$env:USERPROFILE\AppData\Local\Packages\Microsoft.WindowsTerminal*\localState\settings.json"

###default app settings
## https://support.microsoft.com/en-us/windows/command-prompt-and-windows-powershell-for-windows-11-6453ce98-da91-476f-8651-5c14d5777c20

$RegPath = 'HKCU:\Console\%%Startup'
if(!(test-path $RegPath)){New-Item -Path $RegPath}
$RegKey = 'DelegationConsole'
$RegValue = '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
set-ItemProperty -Path $RegPath -Name $RegKey -Value $RegValue -Force | Out-Null

$RegKey2 = 'DelegationTerminal'
$RegValue2 = '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
set-ItemProperty -Path $RegPath -Name $RegKey2 -Value $RegValue2 -Force | Out-Null
}

#endregion
}
catch{
Write-Host "change cmd"
}
#endregion


######## copy importExcel Module to PS path ####
Write-Output "install excel module first < 1 min, please wait"

$PSfolder=(($env:PSModulePath).split(";")|Where-Object{$_ -match "Users" -and $_ -match "WindowsPowerShell"})+"\"+"importexcel"
$checkPSfolder=(Get-ChildItem $PSfolder  -Recurse -file -Filter ImportExcel.psd1 -ErrorAction SilentlyContinue).count

if($checkPSfolder -eq 0){
### remind if block by trendmicro ###
[System.Windows.Forms.MessageBox]::Show($this, "若執行途中遇到掃毒擋powershell請許可")   
 if (!(test-path $PSfolder )){
 New-Item -ItemType directory $PSfolder |Out-Null
 } 
 $A1=(Get-ChildItem "$scriptRoot\toc\batch\importexcel*.zip").fullname
 $shell.NameSpace($PSfolder).copyhere($shell.NameSpace($A1).Items(),4)
}

$checkPSfolder=(Get-ChildItem $PSfolder  -Recurse -file -Filter ImportExcel.psd1).count
if($checkPSfolder -eq 0){
[System.Windows.Forms.MessageBox]::Show($this, "ImportExcel Package Tool 安裝失敗，請重新執行")   
exit

}

#region get module.log

#$testmode1="C:\Jumpstart\Tag\Testmodel1.log"
$testmodelog="C:\Jumpstart\Tag\Testmodel.log"
$qualog="C:\Jumpstart\Tag\Quarter.log"
$prdlinelog="C:\Jumpstart\tag\Productline.log"

if(!(test-path $testmodelog)){
    new-item -ItemType File -Path $testmodelog -Force |Out-Null
    new-item -ItemType File -Path $qualog -Force |Out-Null
    new-item -ItemType File -Path $prdlinelog -Force |Out-Null

    $moduleslog="C:\Windows\modules.log"
    $revfolder="c:\Windows\REVISION"
    
    $modulelogs=get-content -path $moduleslog
    #new-item -ItemType File -Path $testmode1 -Force |Out-Null
    $xmlname=$modulelogs -match "XML name is\:"
    if($xmlname.count -gt 1){
        $xmlname=$xmlname[0]
    }
    $model=(($xmlname -replace "XML name is\:","").split("-"))[2]
    $qua=(($xmlname -replace "XML name is\:","").split("-"))[1]
       
    if (test-path $revfolder) {
      $prd="Comm"
    }
    else{
        $prd="Cons"
    }
     set-content -path $testmodelog -Value $model -Force
     set-content -path $qualog -Value $qua -Force
     set-content -path $prdlinelog -Value $prd -Force
     }    
#endregion


#region DMI code
$dmi1log="C:\Jumpstart\tag\DMI1.log"
$dmi2log="C:\Jumpstart\tag\DMI2.log"

if(!(test-path $dmi1log)){
   
    start-process msinfo32 -ArgumentList /report, C:\Jumpstart\tag\msinfo.log -wait
    $msinfo=get-content C:\Jumpstart\tag\msinfo.log
    (($msinfo -match "システムモデル") -replace "システムモデル","").trim() | Out-File -FilePath $dmi1log 
     (($msinfo -match "システム SKU") -replace "システム SKU","").trim()  | Out-File -FilePath $dmi2log
     #-> systemmodel
     # -> system SKU


}

#endregion

Write-Output "Hello!"
Start-Sleep -s 1
Write-Output "This is going to install the Jumpstart AutoRunner tool."
write-host -BackgroundColor yellow  -ForegroundColor Black "-------[1-1]Copy Batch File-------"

try{
    xcopy "$scriptRoot\toc\batch\*.*" c:\Jumpstart\batch\ /s /v /h /y
    xcopy "$scriptRoot\CopyAndRun*" c:\Jumpstart\batch\ /y
    }
    catch{
        write-host -BackgroundColor red "[1-1] Failed with the error occurred: $_"
    }

 write-host -BackgroundColor yellow -ForegroundColor Black "-------[1-2]Copy Performance Folder-------"
 try{
    xcopy "$scriptRoot\toc\performance" c:\Jumpstart\performance\ /s /v /h /y
    }
    catch{
        write-host -BackgroundColor red "[1-2] Failed with the error occurred: $_"
    }

 write-host -BackgroundColor yellow -ForegroundColor Black "-------[1-3]Copy Shortcut Folder-------"
 try{
    xcopy "$scriptRoot\toc\lnk" "$env:userprofile\desktop\" /s /v /h /y
    }
    catch{
        write-host -BackgroundColor red "[1-3] Failed with the error occurred: $_"
    }

  #region  1,2 BIOS/OS check

  write-host -BackgroundColor DarkGreen -ForegroundColor White "Let's check basic information first before copying tool~~"   
  $BIOS=systeminfo | findstr /I /c:bios
  
  $name=(Get-WmiObject Win32_OperatingSystem).caption
  $bit=(Get-WmiObject Win32_OperatingSystem).OSArchitecture
  $Versiona=(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion).DisplayVersion
  $Version = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\'
  $OScheck=" $name, $bit, $Versiona (OS Build $($Version.CurrentBuildNumber).$($Version.UBR))"
 
   write-host "1. BIOS: `n $BIOS"
   write-host "2. OS Version: `n $OScheck"
   write-host ""

   $biosans=read-host "BIOS and OS version is 1. correct 2. NOT correct  [select 1 or 2] then press Enter"
   if($biosans -match "2"){
   write-host -ForegroundColor Red "!!Important!! Please correct the system environment!"
   pause
   exit
   }
   
  #endregion

  #region  3 yellowbang check
  $ye=Get-WmiObject Win32_PnPEntity|Where-Object{ $_.ConfigManagerErrorCode -ne 0}|Select-Object Name,Description, DeviceID, Manufacturer
  if($ye.DeviceID.Count -gt 0){
  Start-Process devmgmt.msc
  write-host -ForegroundColor Red "!!Important!! Please fix driver problem first! `n"
  pause
  exit
  }
  
  write-host "3. Driver status: `n No yellow bang, continue `n"
  #endregion

  #region  4
  write-host "4. FSU Power Setting"
  if(!(test-path C:\Jumpstart\batch\Power.txt)){
      new-item -ItemType File -Path C:\Jumpstart\batch\Power.txt -Force |Out-Null 
  $pwset=read-host "1. Change FSU Power setting (C/J)  2. No Change Power setting (C/J) [select 1 or 2] then press Enter"
  if( $pwset -match "1"){
      Write-Host "Need Change the Power setting"
      set-content C:\Jumpstart\batch\Power.txt -Value "1"
  }
  if( $pwset -match "2"){
      Write-Host "No Power setting Change needed"
      set-content C:\Jumpstart\batch\Power.txt -Value "0"
  }
  }
  #endregion

  #region  5 #
  write-host "5. check the tester name"
  $filePath = 'C:\Jumpstart\Tag\Tester.log'
  if (Test-Path $filePath) { 
    $tester = Get-Content $filePath -First 1
     Write-Host ('Hello ' + $tester) } 
  else { 
    do{
        $tester = Read-Host 'What is your name'
    }until($tester.length -gt 0)
    
     Write-Host ('Hello ' + $tester)
     new-item 'C:\Jumpstart\Tag\Tester.log' -value $tester  -force |out-null
    }
    write-host ""
  #endregion  5 #  

   #region  6 #
   
  $CurrentUser=$env:USERNAME
  write-host "6. check the test ID"
  $checkid= Read-Host "If The Test ID is [$($CurrentUser)]  (should be like [No.01]) `n 1. correct 2. NOT correct  [select 1 or 2] then press Enter"
  if($checkid -match "2"){
    do{
    do{
    $newname= Read-Host "Enter the current Test ID"
    }until($newname.length -gt 0)
    $checkid= Read-Host "If The Test ID is [$($newname)]  (should be like [No.01]) `n 1. correct 2. NOT correct  [select 1 or 2] then press Enter"
    }until($checkid -eq 1)
    Rename-LocalUser -Name  $CurrentUser -NewName $newname
  }
  
  #endregion  6 # 
  
    $projlog="C:\Jumpstart\Tag\Project.log"
    if(!(test-path $projlog)){
        new-item -ItemType File -Path $projlog -Force |Out-Null
    }
    set-content -Path $projlog -Value "Jumpstart"
  
   #region Phase.log
   $phaselog="C:\Jumpstart\Tag\Phase.log"
   write-host "6. Test phase"
   if(!(test-path $phaselog)){
       $phaseselections=@("1 - FC","2 - Final","3 - F1","4 - F2","5 - Other")
       new-item -ItemType File -Path $phaselog -Force |Out-Null
       write-host "Please 1, 2, 3, 4, 5 to choose pattern"
       $phaseselections|ForEach-Object{
        write-host $_
       }
       $phasen=Read-Host "select"
       if($phasen -match "5"){
           $phasename=Read-Host "Please input the phase"
       }
       else{
        $phasename=(($phaseselections|Where-Object{$_ -match $phasen}).split("-")[1]).trim()
       }
    
       set-content -Path $phaselog -value $phasename -Force
   }
   else{
       $phasename=get-content -Path $phaselog
   }
   
   write-host "Phase: $phasename `n"
   #endregion Phase.log
   
   #region MID

   $midlog="C:\Jumpstart\Tag\MID.log"
   write-host "7. Machine ID"
   if(!(test-path $midlog)){
    new-item -ItemType File -Path $midlog -Force |Out-Null
    $mid=Read-Host "Please input the MID"
    set-content -Path $midlog -value $mid -Force
     }
     else{
        $mid=get-content -Path $midlog
     }
   write-host "MID: $mid `n"

   #endregion MID
   
   #region SKU
   $skulog="C:\Jumpstart\Tag\SKU.log"
   write-host "8. SKU Info"
   if(!(test-path $skulog)){
    new-item -ItemType File -Path $skulog -Force |Out-Null
    $sku=Read-Host "Enter SKU"
    set-content -Path $skulog -value $sku -Force
     }
     else{
        $sku=get-content -Path $skulog
     }
   write-host "SKU: $sku `n"

   #endregion MID
   

   #region Report Date
   $repdatelog="C:\Jumpstart\batch\Report_Date.txt"
   write-host "9. Report Date"
   if(!(test-path $repdatelog)){
    new-item -ItemType File -Path $repdatelog -Force |Out-Null
    $repdate=Read-Host "6. Input the Report Date (ex: 2022/02/11) "
    set-content -Path $repdatelog -value $repdate -Force
     }
     else{
        $repdate=get-content -Path $repdatelog
     }
   write-host "Report Date: $repdate `n"

   #endregion

   #region test mode

   $testmodelog="C:\Jumpstart\Tag\Mode.log"
   write-host "10. Test Mode"
   if(!(test-path $testmodelog)){
    new-item -ItemType File -Path $testmodelog -Force |Out-Null
    $testmode=Read-Host "Enter Mode (1 - OEM, 2 - Baseline)"
    if($testmode -match "1"){
        $testmode="OEM"
    }
    if($testmode -match "2"){
        $testmode="Baseline"
    }
    set-content -Path $testmodelog -value $testmode -Force
     }
     else{
        $testmode=get-content -Path $testmodelog
     }
   write-host "Test Mode: $testmode `n"

   #endregion
   
   
    #region uac and reboot
    Set-ItemProperty -Path HKLM:Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA  -Value 0 -Force     
    write-host "Disable uac: change done, need reboot `n"
    xcopy "$scriptRoot\toc\batch\startup" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\" /s /v /h /y
    Write-Output "Now, let's reboot system then start to run tool~"
    start-sleep -s 30
    shutdown /r /t 0
    #endregion

