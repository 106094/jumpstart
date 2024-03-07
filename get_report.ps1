 Add-Type -AssemblyName System.Windows.Forms

$winv= ([System.Environment]::OSVersion.Version).Build
if($winv -ge 22000){$spec=import-csv C:\Jumpstart\batch\Spec\target_Win11.csv}
else{$spec=import-csv C:\Jumpstart\batch\Spec\target_Win10.csv}

$modln=get-content C:\Jumpstart\Tag\Testmodel.log
$rdate=get-content C:\Jumpstart\batch\Report_Date.txt
$tempcsv="C:\Jumpstart\batch\Spec\Template.csv"
$temptxt="C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt"
$asstempcsv="C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.csv"
$asstempxlsx="C:\Jumpstart\batch\Spec\Performance_Assessment_Toolkit_Reporting_Template.xlsx"
$assxlsx="C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.xlsx"
$OEMresults=""

$result_csvall=(Get-ChildItem -path C:\Jumpstart\performance\results\OEM*results.csv).fullname
$result_baseline=(Get-ChildItem -path C:\Jumpstart\performance\results\baseline*results.csv).fullname

if($result_csvall.count -eq 0 -and $result_baseline.count -eq 0){
 [System.Windows.Forms.MessageBox]::Show($this,"No result csv has been found, please check!")|out-null
  exit
}

#baseline remove old results
if($result_baseline.count -gt 1){
    $result_baseline_csv=(Get-ChildItem -path C:\Jumpstart\performance\results\OEM*results.csv|Sort-Object LastWriteTime |Select-Object -Last 1).fullname
    $result_baseline -ne $result_baseline_csv|Remove-Item -Force
}

# generator temp csv before baseline test
if($result_csvall.count -ne 0 -and $result_baseline.count -eq 0){

    $result_csv=(Get-ChildItem -path C:\Jumpstart\performance\results\OEM*results.csv|Sort-Object LastWriteTime |Select-Object -Last 1).fullname
   
    ## remove the old OEM results#
    if($result_csvall.count -gt 1){
        $result_csvall -ne $result_csv|Remove-Item -Force 
    }
    
$table_content=import-csv $tempcsv -Encoding UTF8
$result_content=import-csv $result_csv -Encoding UTF8

$table_content[1].2=$modln
$table_content[3].2=$rdate

#######Get Configuration####
$config=$null
$config= "Configuration "+($result_content."config"|Sort-Object|Get-Unique)
$table_content[2].2=$config

#######Get KitVersion####

#######GetTesting Date####

$testitems=@("FastStartup","Edge","BatteryLife","Standby")

foreach($testitem in $testitems){
$row=$testitems.IndexOf($testitem)+6
$row_spec=$testitems.IndexOf($testitem)
#$testitem
#$table_content[$row].2
#$spec[$row_spec].$config

if($testitem -eq "FastStartup"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem -and $_."testcase" -eq "Total"})."metricvalue"| Measure-Object -Minimum).Minimum
$pass_flag=($result_content|Where-Object {$_."assessment" -eq $testitem -and $_."metricvalue" -eq $item_value})."status"

$table_content[$row].2=$item_value
$table_content[$row].4=$pass_flag

$unit=$table_content[$row].5
$spectarget=$spec[$row_spec].$config+" "+$unit
$table_content[$row].5=$spectarget
}

if($testitem -eq "BatteryLife"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem})."metricvalue"| Measure-Object -Maximum).Maximum
$pass_flag=($result_content|Where-Object {$_."assessment" -eq $testitem -and $_."metricvalue" -eq $item_value})."status"

$table_content[$row].2=$item_value
$table_content[$row].4=$pass_flag

$unit=$table_content[$row].5
$spectarget=$spec[$row_spec].$config+" "+$unit
$table_content[$row].5=$spectarget
}

if($testitem -eq "Edge"-or $testitem -eq "Standby"){
$item_value=(($result_content|Where-Object {$_."assessment" -eq $testitem})."metricvalue" | Measure-Object -Minimum).Minimum
$pass_flag=($result_content|Where-Object {$_."assessment" -eq $testitem -and $_."metricvalue" -eq $item_value})."status"

$table_content[$row].2=$item_value
$table_content[$row].4=$pass_flag

$unit=$table_content[$row].5
$spectarget=$spec[$row_spec].$config+" "+$unit
$table_content[$row].5=$spectarget
}
}

$table_content |export-csv $temptxt -Encoding UTF8 -NoTypeInformation
#(Get-Content $temptxt | Select-Object -Skip 1) | Set-Content $asstempcsv
$csvcontent=Get-Content $temptxt 
if($csvcontent -match "BaselineRequired"){
    $OEMresults="fail"
}
$csvcontent |  Set-Content $asstempcsv
Remove-Item $temptxt -Force

}

if($OEMresults -eq "fail" -and $result_baseline.count -eq 0){   
    [System.Windows.Forms.MessageBox]::Show($this,"OEM test failed, need continue to test baseline condition") |out-null
    exit
   }
else{
    #copy csv data to excel file when OEM passed or baseline test finished
Copy-Item $asstempxlsx -Destination $assxlsx -Force
Import-Csv -Path $asstempcsv | Export-Excel -Path $assxlsx -WorksheetName 'Performance-Assessment-Results' -NoNumberConversion 3,4

$excelPackage = Open-ExcelPackage -Path $assxlsx
$worksheet = $excelPackage.Workbook.Worksheets['Performance-Assessment-Results'] # Assuming it's the first worksheet
$columnIndexs = @(2,3) # Specify the column index

foreach ($columnIndex in $columnIndexs) {
# Loop through each cell in the specified column
foreach ($row in 1..$worksheet.Dimension.End.Row) {
    $cellValue = $worksheet.Cells[$row, $columnIndex].Text
    if ($cellValue -match "\.") {
        # If the cell contains a decimal point, format as decimal
        $worksheet.Cells[$row, $columnIndex].Style.Numberformat.Format = "0.##"
    } else {
        # If the cell does not contain a decimal point, format as integer
        $worksheet.Cells[$row, $columnIndex].Style.Numberformat.Format = "0"
    }
} 
}

foreach ($column in 1..$worksheet.Dimension.End.Column) {
    $worksheet.Cells[1, $column].Value = $null
}

$excelPackage.Save()
$excelPackage.Dispose()

remove-item $asstempcsv -Force

#get the zip file name
$DMI=((get-content -path "C:\Jumpstart\tag\DMI1.log")|Out-String).trim()
$MID=$env:USERNAME
$mtype=$DMI.Substring(4,1)
$typename="Note."
if($mtype -eq "D"){
    $typename="Mate."
}
$zipname=$typename+$MID+"_"+$DMI

$sourceFolder = "C:\Jumpstart\performance\results"  
$zipFile = "$env:userprofile\desktop\"+$zipname+".zip"    

try {
       # Compress the folder
       Compress-Archive -Path $sourceFolder -DestinationPath $zipFile -Force
       Write-Host "Folder '$sourceFolder' has been compressed to '$zipFile'."
       $sizecheck=0
       do{
        start-sleep -s 10
        $sizecheckold=$sizecheck
        $sizecheck=(get-childitem $zipFile).length
       }until($sizecheckold -eq $sizecheck -and $sizecheck -ne 0)
             
    Write-Output "data compress is completed "
    #[System.Windows.Forms.MessageBox]::Show($this,"Please check $zipFile") |out-null
    
    #region change DNS to auto
            
        $wificonfig="C:\Jumpstart\tag\wifiname.txt"
        $adtname=get-content  $wificonfig

        Set-DnsClientServerAddress -InterfaceAlias $adtname -ResetServerAddresses
        Clear-DnsClientCache                
        #check connection
        $ping = New-Object System.Net.NetworkInformation.Ping
        $timex1=get-date
        do{
            start-sleep -s 10
            $checkconnect=($ping.Send("www.google.com", 1000)).Status
            $timepassed=(New-TimeSpan -start  $timex1 -End (get-date)).TotalSeconds
        } until($checkconnect -match "Success" -or $timepassed -ge 100)
        if($timepassed -ge 100){
        [System.Windows.Forms.MessageBox]::Show($this,"測網連線失敗") |out-null
        exit
        }

    #endregion

    ### upload to server ##
    # AutoUploadLogsToServer_Rev3.bat


    [System.Windows.Forms.MessageBox]::Show($this,"$zipFile upload ok")

}
catch {
     Write-Output "data compress failed "
      [System.Windows.Forms.MessageBox]::Show($this,"fail to compress result folder, please check!") |out-null
} 

}