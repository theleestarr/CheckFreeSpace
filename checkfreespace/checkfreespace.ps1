# -----------------------------------------------------------------------------############################
# Script Function: This sample Windows powershell script calculates free disk spaces
# in multiple servers and emails copy of  csv report. 
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

$File = Get-Content -Path C:\Scripts\DiskReport\CheckFreeSpace\Servers.txt

#Define admin account variable (Uncommented it and the line in Get-WmiObject
#Line 44 below is commented out because I run this script via task schedule.
#If you wish to run this script manually, please uncomment line 44.


$RunAccount = get-Credential  

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
@{Label = "Total Capacity (GB)";Expression = {"{0:N1}" -f( $_.Size / 1gb)}},
@{Label = "Free Space (GB)";Expression = {"{0:N1}" -f( $_.Freespace / 1gb ) }},
@{Label = 'Free Space (%)'; Expression = {"{0:P0}" -f ($_.freespace/$_.size)}} |


#Export report to CSV file (Disk Report)

Export-Csv -path "C:\Scripts\DiskReport\CheckFreeSpace\Reports\RHB_Windows_DiskReport_$logDate.csv" -NoTypeInformation



#Send disk report using the exchange email module

$fromaddress = "WindowsDriveReport@rhb.ad" 
$toaddress = "lee.haarmann@changehealthcare.com" 
$CCaddress = "" 
$Subject = "Windows Server Free Space Report" 
$body = get-content "C:\Scripts\DiskReport\CheckFreeSpace\content.htm" 
$attachment = (Get-ChildItem C:\Scripts\DiskReport\CheckFreeSpace\Reports\*.* | sort LastWriteTime | select -last 1)
$smtpserver = "sjmgmail.rhb.ad"

#################################### 
 
#$message = new-object System.Net.Mail.MailMessage 
#$message.From = $fromaddress 
#$message.To.Add($toaddress) 
#$message.CC.Add($CCaddress)  
#$message.IsBodyHtml = $True 
#$message.Subject = $Subject 
#$attach = new-object Net.Mail.Attachment($attachment) 
#$message.Attachments.Add($attach) 
#$message.body = $body 
#$smtp = new-object Net.Mail.SmtpClient($smtpserver) 
#$smtp.Send($message) 




## END OF SCRIPT ###
