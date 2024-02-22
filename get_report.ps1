 Add-Type -AssemblyName System.Windows.Forms

$result_csv=(gci -path C:\Jumpstart\performance\results\*results.csv|sort LastWriteTime |select -Last 1).fullname
$winv= ([System.Environment]::OSVersion.Version).Build
if($winv -eq 22000){$spec=import-csv C:\Jumpstart\batch\Spec\target_Win11.csv}
else{$spec=import-csv C:\Jumpstart\batch\Spec\target_Win10.csv}
$modln=get-content C:\Jumpstart\batch\Model_Name.txt
$rdate=get-content C:\Jumpstart\batch\Report_Date.txt

if($result_csv.count -ne 0){
$table_content=import-csv C:\Jumpstart\batch\Spec\Template.csv -Encoding UTF8
$result_content=import-csv $result_csv -Encoding UTF8

$table_content[1].2=$modln
$table_content[3].2=$rdate

#######Get Configuration####
$config=$null
$config= "Configuration "+($result_content."config"|sort|Get-Unique)
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
$item_value=(($result_content|? {$_."assessment" -eq $testitem -and $_."testcase" -eq "Total"})."metricvalue"| measure -Minimum).Minimum
$pass_flag=($result_content|? {$_."assessment" -eq $testitem -and $_."metricvalue" -eq $item_value})."status"

$table_content[$row].2=$item_value
$table_content[$row].4=$pass_flag

$unit=$table_content[$row].5
$spectarget=$spec[$row_spec].$config+" "+$unit
$table_content[$row].5=$spectarget
}

if($testitem -eq "BatteryLife"){
$item_value=(($result_content|? {$_."assessment" -eq $testitem})."metricvalue"| measure -Maximum).Maximum
$pass_flag=($result_content|? {$_."assessment" -eq $testitem -and $_."metricvalue" -eq $item_value})."status"

$table_content[$row].2=$item_value
$table_content[$row].4=$pass_flag

$unit=$table_content[$row].5
$spectarget=$spec[$row_spec].$config+" "+$unit
$table_content[$row].5=$spectarget
}

if($testitem -eq "Edge"-or $testitem -eq "Standby"){
$item_value=(($result_content|? {$_."assessment" -eq $testitem})."metricvalue" | measure -Minimum).Minimum
$pass_flag=($result_content|? {$_."assessment" -eq $testitem -and $_."metricvalue" -eq $item_value})."status"

$table_content[$row].2=$item_value
$table_content[$row].4=$pass_flag

$unit=$table_content[$row].5
$spectarget=$spec[$row_spec].$config+" "+$unit
$table_content[$row].5=$spectarget
}

}

$table_content |export-csv C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt -Encoding UTF8 -NoTypeInformation
(Get-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt | Select-Object -Skip 1) | Set-Content C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template.csv
Remove-Item C:\Jumpstart\performance\results\Performance_Assessment_Toolkit_Reporting_Template0.txt -Force
   
 [System.Windows.Forms.MessageBox]::Show($this,"Please check ""Performance_Assessment_Toolkit_Reporting_Template.csv"" file in the results Folder!") |out-null
}
else{

 [System.Windows.Forms.MessageBox]::Show($this,"No result csv has been found, please check!")|out-null

}