# -----------------------------------------------------------------------------############################
# Script Function: This sample Windows powershell script calculates free disk spaces
# in multiple servers and emails copy of  csv report. 
# Author: Joel Souza 
# Date: 06/20/2015 
# 
# ----------------------------------------------------------------------------- ############################
$erroractionpreference = "SilentlyContinue"


## BEGGINNIBF OF SCRIPT ###

#Set execution policy to Unrestricted (-Force suppresses any confirmation)
#Execution policy stopped the script from running via task scheduler
#As a work-around, I added an action in the task scheduler to run first before this script runs
# Set-ExecutionPolicy Unrestricted -Force

Set-ExecutionPolicy Unrestricted -Force

#delete reports older than 7 days

$OldReports = (Get-Date).AddDays(-7)

#edit the line below to the location you store your disk reports# It might also
#be stored on a local file system for example, D:\ServerStorageReport\DiskReport

Get-ChildItem D:\DiskReport
Where-Object { $_.LastWriteTime -le $OldReports} | `
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue  

#Create variable for log date

$LogDate = get-date -f yyyyMMddhhmm


#Define location of text file containing your servers. It might also
#be stored on a local file system for example, D:\ServerStorageReport\DiskReport

$File = Get-Content -Path E:\lhaarmann\DiskReport\CheckFreeSpace\Servers.txt

#Define admin account variable (Uncommented it and the line in Get-WmiObject
#Line 44 below is commented out because I run this script via task schedule.
#If you wish to run this script manually, please uncomment line 44.


#$RunAccount = get-Credential  

#the disk $DiskReport variable checks all servers returned by the $File variable (line 37)

$DiskReport = ForEach ($Servernames in ($File)) 

{

Get-WmiObject win32_logicaldisk `
-ComputerName $Servernames -Filter "Drivetype=3" `
-ea SilentlyContinue 

#return only disks with
#free space less  
#than or equal to 0.1 (10%)

Where-Object {($_.freespace/$_.size) -lt .1}

} 


#create reports

$DiskReport | 

Select-Object @{Label = "Server Name";Expression = {$_.SystemName}},
@{Label = "Drive Letter";Expression = {$_.DeviceID}},
@{Label = 'SizeInGB';Expression={[math]::round($_.Size / 1GB,2)}},
@{Label = 'FreeSpaceInGB';Expression={[math]::round($_.FreeSpace / 1GB,2)}},
@{Label = 'FreeSpaceIn%';Expression={[math]::round($_.FreeSpace / $_.Size * 100,2)}},
@{Label = 'FreeSpaceLessThan10%';Expression={If(($_.FreeSpace / $_.Size * 100) -lt 10){$true}Else{$false}}}|


#Export report to CSV file (Disk Report)

Export-Csv -path "E:\lhaarmann\DiskReport\CheckFreeSpace\Reports\RHB_Windows_DiskReport_$logDate.csv" -NoTypeInformation

#Send-MailMessage
Send-MailMessage -From 'canoc@changehealthcare.com' -To 'lee.haarmann@changehealthcare.com' -Attachment (Get-ChildItem E:\lhaarmann\DiskReport\CheckFreeSpace\Reports\*.* | sort LastWriteTime | select -last 1) -Subject 'Disk Space Check' -SmtpServer 'SJMGMAIL.RHF.AD'

## END OF SCRIPT ###
