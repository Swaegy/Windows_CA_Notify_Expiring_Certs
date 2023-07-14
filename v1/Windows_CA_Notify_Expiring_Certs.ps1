#------------------------------------#
#--------- ARC-XX  v2304.16 ---------#
#------------------------------------#

# Cert Expiry Export Variables 
$dateformat = 'dd.MM.yyyy'
$dateformatcsv = 'yyyy.MM.dd'
$Today = Get-Date -Format $dateformat
$Todaycsv = Get-Date -Format $dateformatcsv
$30Days = (Get-Date).AddDays(+30)
$ExpireDate = Get-Date -Date $30Days -Format $dateformat

# Export List of expiring Certificates as CSV for  
certutil -view -restrict "Certificate Expiration Date <= $ExpireDate,Certificate Expiration Date > $Today,Disposition = 20" -out "Issued Common Name,Issued Email Address,Certificate Template,Certificate Effective Date,Certificate Expiration Date" csv > .\CertExpiring_$Todaycsv.csv

# Get Certifiactes expiring in 30 days into variable for output in Mailbody
$certexport = Import-CSV .\CertExpiring_$Todaycsv.csv | ConvertTo-Html -Fragment

# Static Mail send Variables
$hostname = $env:COMPUTERNAME
$to = '' # Set your Receivers Mailaddresses here, 'user1@example.com'
$cc = '' # Use as fallback if the main Mailaddress is no longer available 'user1@example.com'
$smtphost = '' # FQDN or IP of your SMTP Server or Relay, 'smtp.example.com'
$maildomain = '' # Add your Maildomain, only needed for the Senderadress which is build from the Systemname and the Maildomain, 'example.com'
$from = "$hostname@$maildomain"
$subject = "SSL Zertifikatsablauf $hostname"
$body = "This is a overview over all Certificates that are going to expire on the CA $hostname in the next 30 Days.<br>
This is an automated E-Mail, do not reply to this Mailaddress!<br>
<br>
<br>
<br>
$certexport"

# SendMail for expiring certs
Send-MailMessage -To $to -Cc $cc -SmtpServer $smtphost -From $from -Subject $subject -Body $body -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
