<#
	.SYNOPSIS
		Windows_CA_Notify_Expiring_Certs.ps1

	.DESCRIPTION
		The Script is used to export the soon to be expiring certs from a CA.
        Check if these Certificates are already renewed, based on the Common Name of the Cert.
        Afterwards there is a Mail gernerated based on the Value of the Issued Mail Address only containing Certs for that Mail address.

	.NOTES
		Swaegy
		v2501.13

	.LINK
		https://github.com/Swaegy

#>
# Cert Expiry Export Variables 
$dateformat = 'dd.MM.yyyy'
$dateformatcsv = 'yyyy.MM.dd'
$Today = Get-Date -Format $dateformat
$Todaycsv = Get-Date -Format $dateformatcsv
$Add30Days = (Get-Date).AddDays(+30)
$ExpireDate = Get-Date -Date $Add30Days -Format $dateformat
$Remove30Days = (Get-Date).AddDays(-30)
$EffectiveDate = Get-Date -Date $Remove30Days -Format $dateformat
$outformat = "Issued Common Name,SerialNumber,Certificate Effective Date,Certificate Expiration Date,Issued Email Address,Certificate Template"

# Static Mail send Variables
$hostname = $env:COMPUTERNAME # Needed for the Senderadress
$cc = '' # Use as fallback if the Mailaddress is no set for a certificate, 'fallback@example.com'
$to = '' # is filled with each itteration of the foreach
$smtphost = '' # FQDN or IP of your SMTP Server or Relay, 'smtp.example.com'
$maildomain = '' # Add your Maildomain, only needed for the Senderadress which is build from the Systemname and the Maildomain, 'example.com'
$from = "$hostname@$maildomain"
$subject = "SSL Zertifikatsablauf $hostname"
$body = '' # gets filled later in the foreach loop
$htmltableheader = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

# Export List of expiring Certifiactes as CSV for debugging 
certutil -view -restrict "Certificate Expiration Date <= $ExpireDate,Certificate Expiration Date > $Today,Disposition = 20" -out "$outformat" csv > ".\CertExpiring_$Todaycsv.csv"

# Get Certifiactes expiring in 30 days into veriable for output im Mailbody
$ExpiringCerts = Import-Csv -Path ".\CertExpiring_$Todaycsv.csv"

# Get the unique email addresses from the desired row
$ExpiringCertsUniqueMailadresses = $ExpiringCerts.'Issued Email Address' | Sort-Object -Unique

# Get most recently enrolled certificates for matching it with the expiring certs
certutil -view -restrict "Certificate Effective Date <= $Today,Certificate Effective Date > $EffectiveDate,Disposition = 20" -out "$outformat" csv > ".\CertRenewed_$Todaycsv.csv"
$RenewedCerts = Import-Csv -Path ".\CertRenewed_$Todaycsv.csv"
$RenewedCertsFilter = $RenewedCerts.'Issued Common Name' -join '|'

# Loop through each unique email address
foreach ($UniqueMailadresses in $ExpiringCertsUniqueMailadresses) {
    if ($UniqueMailadresses -eq 'EMPTY'){
        # Create a mail that is sent to a specific mail address
        $body = "Dies ist eine Übersicht über alle Zertifikate die von der CA $hostname ausgestellt wurden und in weniger als 30 Tagen auslaufen werden.<br>Dies ist eine automatisch generierte E-Mail bitte nicht darauf antworten!<br>Diese Mail enthält alle Zertifikate die keine Issued Mail Address angegeben haben.<br><br>"

        # Get all certificates, filtering based on the Issued Email Address happens later
        $ExpiringCertsByMailaddress = certutil -view -restrict "Certificate Expiration Date <= $ExpireDate,Certificate Expiration Date > $Today,Disposition = 20" -out "$outformat" csv
        
        # Filter out the Certificates that are already renewed based on the Issued Common Name and only for the certificates that have 'EMPTY' as Issued Email Address
        $ExpiringCertsByMailaddressFiltered = $ExpiringCertsByMailaddress | ConvertFrom-Csv | Where-Object { $_.'Issued Email Address' -eq 'EMPTY' -and $_.'Issued Common Name' -notmatch "$RenewedCertsFilter" } | ConvertTo-Csv

        # Check if there are still expiring certificates that are not already renewed
        if ($ExpiringCertsByMailaddressFiltered){

        # Create a HTML Table to integrate the Infomration into the mailbody
        $body += $ExpiringCertsByMailaddressFiltered | ConvertFrom-Csv | ConvertTo-Html -Head $htmltableheader
        
        # Set the emailadress from the expiring certs that have no issueing Mail Address to the Backup Mailaddress
        $to = $cc

        # Send the email
        Send-MailMessage -SmtpServer $smtphost -From $from -To $to -Subject $subject -Body $body -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
        }
    }
    else {
        # Create a mail that is sent to the unique email address
        $body = "Dies ist eine Übersicht über alle Zertifikate die von der CA $hostname ausgestellt wurden und in weniger als 30 Tagen auslaufen werden.<br>Dies ist eine automatisch generierte E-Mail bitte nicht darauf antworten!<br><br>"

        # Get all the rows that have this email address
        $ExpiringCertsByMailaddress = certutil -view -restrict "Certificate Expiration Date <= $ExpireDate,Certificate Expiration Date > $Today,Disposition = 20,Issued Email Address = $UniqueMailadresses" -out "$outformat" csv
        
        # Filter out the Certificates that are already renewed based on the Issued Common Name
        $ExpiringCertsByMailaddressFiltered = $ExpiringCertsByMailaddress | Where-Object { $_.'Issued Common Name' -notmatch "$RenewedCertsFilter" }

        # Check if there are still expiring certificates that are not already renewed
        if ($ExpiringCertsByMailaddressFiltered){

        # Create a HTML Table to integrate the Infomration into the mailbody
        $body += $ExpiringCertsByMailaddress | ConvertFrom-Csv | ConvertTo-Html -Head $htmltableheader

        # Set the emailadress from the expiring certs as the $to Mailaddress
        $to = $UniqueMailadresses

        # Send the email
        Send-MailMessage -SmtpServer $smtphost -From $from -To $to -Subject $subject -Body $body -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
        }
    }
}
