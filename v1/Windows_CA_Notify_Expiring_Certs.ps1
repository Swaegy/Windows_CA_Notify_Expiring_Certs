<#
	.SYNOPSIS
		PKI_Get_ExipringCerts.ps1

	.DESCRIPTION
		The Script is used to export the soon to be expiring certs from a CA and Send a Mailreport with the Information within the Mailbody.

	.NOTES
		Swaegy
		v2301.16

	.LINK
		https://github.com/Swaegy

#>
# Cert Expiry Export Variables 
$dateformat = 'dd.MM.yyyy'
$dateformatcsv = 'yyyy.MM.dd'
$Today = Get-Date -Format $dateformat
$Todaycsv = Get-Date -Format $dateformatcsv
$30Days = (Get-Date).AddDays(+30)
$ExpireDate = Get-Date -Date $30Days -Format $dateformat

# Export List of expiring Certifiactes as CSV for debugging 
certutil -view -restrict "Certificate Expiration Date <= $ExpireDate,Certificate Expiration Date > $Today,Disposition = 20" -out "Issued Common Name,Issued Email Address,Certificate Template,Certificate Effective Date,Certificate Expiration Date" csv > C:\Report\CertExpiring_$Todaycsv.csv

# Get Certifiactes expiring in 30 days into veriable for output im Mailbody
$certexport = Import-CSV C:\Report\CertExpiring_$Todaycsv.csv | ConvertTo-Html -Fragment

# Mail send Variables
$hostname = $env:COMPUTERNAME
$to = ""
$cc = ""
$smtphost = ""
$from = "$hostname-"
$subject = "SSL Zertifikatsablauf $hostname"
$body = "Dies ist eine Übersicht über alle Zertifikate die von der $hostname ausgestellt wurden und in weniger als 30 Tagen auslaufen werden.<br>
Dies ist eine automatisch generierte E-Mail bitte nicht darauf antworten!<br>
<br>
<br>
<br>
$certexport"

# SendMail with Csv attached srv-infra
Send-MailMessage -To $to -Cc $cc -SmtpServer $smtphost -From $from -Subject $subject -Body $body -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
