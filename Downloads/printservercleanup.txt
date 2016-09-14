##Print Queue cleanup by Ryan Nordlocken
##Clear out Print Queues with Old Print jobs

##Items to configure below.  
##$old = How many hours to look back to delete old print Jobs
##$pcname = the servers that you want to scan and remove old print jobs.
##$body = Define each parameter if there is something to send off.
##sendMail function = configure the email relay and addresses.

#Get date and days older than for jobs to be deleted.
$d = Get-Date
$old=$d.AddHours(-1)

##List out the print servers you want to monitor.  Comma seperated list.
$pcname="printserver1","printserver2"
$globallist=@()

##Go through each printer, find old print jobs, delete them and add them to gloabl list to send off in email.
foreach ($i in $pcname){

    $PrintJobs = get-wmiobject -ComputerName $i -class "Win32_PrintJob" -namespace "root\CIMV2" | Where-Object { $_.ConvertToDateTime($_.TimeSubmitted) -lt $old }
    $locallist=@()
    foreach ($job in $PrintJobs) {

       ##Break loop is job is empty
       if ($job -eq $NULL){
            continue
       }

       ##Trim the job name to just the name of the printer
       $pos= $job.name.IndexOf(",")
       $locallist+=$job.Name.substring(0,$pos)+"`n"
       $globallist+=$i+" : "+$job.Name.substring(0,$pos)+"`n"
       $job.Delete()
    }
}
##Select just the unique print queues
$list=$globallist|select -uniq

##get server list into readable format
foreach ($c in $pcname){
    $serverlist+= $c+"`n"
}

##change message of body
if (!$list) {
	#Body of email.
    $body = "It's a good day, nothing is older than 1 hours. `r`n"
}
else {
	#body of email if there is something to send off.
    $body= "Printers with one or more stuck jobs older than 1 hour old.  An attempt to remove these jobs has been completed.  If a printer continues to show up, additional investigation into the printer or print queue is needed.  Can manage the printers directly from: `r`n https://yoururlhere `r`n`n"
}

##Function to send mail
function sendMail ($a, $b){

     #SMTP server name
     $smtpServer = "smtp.relay.org"

     #Creating a Mail object
     $msg = new-object Net.Mail.MailMessage

     #Creating SMTP server object
     $smtp = new-object Net.Mail.SmtpClient($smtpServer)

     #Email structure 
     $msg.From = "from@company.org"
     $msg.ReplyTo = "from@company.org"
     $msg.To.Add("emailaddress1@company.org")
     $msg.To.Add("emailaddress2@company.org")
	 #Can keep adding as many emails as you'd like.
     $msg.subject = $a
     $msg.body = $body+$b + "`n"+"Print servers searched:`n$serverlist" +"`r`n`n The Print Team"
     
     #Sending email 
     $smtp.Send($msg)
  
}

##Calling function to send email.
##sendMail "Print Removal" $list
if (!$list) {
    #Do nothing so no notifications go out.
}
else {
    sendMail "Print Removal" $list
}
