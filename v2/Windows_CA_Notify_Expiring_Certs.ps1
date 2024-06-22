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

# Static Mail send Variables
$hostname = $env:COMPUTERNAME # Needed for the Senderadress
$cc = '' # Use as fallback if the Mailaddress is no set for a certificate, 'fallback@example.com'
$smtphost = '' # FQDN or IP of your SMTP Server or Relay, 'smtp.example.com'
$maildomain = '' # Add your Maildomain, only needed for the Senderadress which is build from the Systemname and the Maildomain, 'example.com'
$from = "$hostname@$maildomain"
$subject = "SSL Zertifikatsablauf $hostname"
$to = '' # is filled with each itteration of the for loop, do not input anything

# Export List of expiring Certifiactes as CSV for debugging 
certutil -view -restrict "Certificate Expiration Date <= $ExpireDate,Certificate Expiration Date > $Today,Disposition = 20" -out "Issued Common Name,Issued Email Address,Certificate Template,Certificate Effective Date,Certificate Expiration Date" csv > .\CertExpiring_$Todaycsv.csv

# Get Certifiactes expiring in 30 days into veriable for output im Mailbody
$data = Import-Csv -Path ".\CertExpiring_$Todaycsv.csv"

# Get the unique email addresses from the desired row
$emailColumn = 'Issued Email Address' # Replace "Email" with the name of the column that contains the email addresses
$uniqueEmails = $data.$emailColumn | Sort-Object -Unique

# Loop through each unique email address
foreach ($email in $uniqueEmails) {
    # Get all the rows that have this email address
    $rowsWithEmail = $data | Where-Object { $_.$emailColumn -eq $email }

    # Check if the email address is "EMPTY"
    if ($email -eq "EMPTY") {
        # Create a mail that is sent to a specific mail address
        $body = "This is a overview over all Certificates that are going to expire on the CA $hostname in the next 30 Days.<br>This is an automated E-Mail, do not reply!<br>This Mail contains the all Certificates that have not set the issued Mail Address.<br>"

        # Loop through each row and add it to the email body
        foreach ($row in $rowsWithEmail) {
            $rowString = ""

            # Loop through each column in the row and add it to the row string
            foreach ($property in $row.PsObject.Properties) {
                $rowString += "$($property.Name): $($property.Value)<br>"
            }

            # Add the row string to the email body
            $body += "<br>$rowString"
        }

        $to = $cc
    } else {
        # Create a mail that is sent to the unique email address
        $body = "This is a overview over all Certificates that are going to expire on the CA $hostname in the next 30 Days.<br>This is an automated E-Mail, do not reply!<br><br>"

        # Loop through each row and add it to the email body
        foreach ($row in $rowsWithEmail) {
            $rowString = ""

            # Loop through each column in the row and add it to the row string
            foreach ($property in $row.PsObject.Properties) {
                $rowString += "$($property.Name): $($property.Value)<br>"
            }

            # Add the row string to the email body
            $body += "<br>$rowString"
        }

        $to = $email
    }

    # Send the email
    Send-MailMessage -SmtpServer $smtphost -From $from -To $to -Subject $subject -Body $body -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
}
