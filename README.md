# Windows CA Notify about Expiring Certificates
This repository contains a collection of scripts that help keep an eye on the certificate expiration on a Windows certification authority (CA). <br /><br />

## Usage
The scripts can be run manually. <br />
But it is advised to automate the job with the Windows Task Scheduler. <br />
For full coverage and to have time left to renew expiring certs, run the script via the Task Scheduler every 14 Days.<br />
You can use the System Account for running the Script, but i advise you to an account with read right to the certificate Authority.<br /><br />

#### Script v1
This Script creates a csv from all certificates that will expire in the next 30 days. <br />
After that it will create a email with all certificates listed in the mailbody as a table. <br />

You will need to fill in the following Infos into the Script:<br />
`$to` = '' # Set your Receivers Mailaddresses here, 'user1@example.com'<br />
`$cc` = '' # Use as fallback if the main Mailaddress is no longer available 'fallback@example.com'<br />
`$smtphost` = '' # FQDN or IP of your SMTP Server or Relay, 'smtp.example.com'<br />
`$maildomain` = '' # Add your Maildomain, only needed for the Senderadress which is build from the Systemname and the Maildomain, 'example.com'<br /><br />

#### Script v2
This Script creates a csv from all certificates that will expire in the next 30 days. <br />
After that it will create a email for each unique emailaddress specified in the certificate, with all corresponding certificates listed in the mailbody. <br />
Watch out for certificates where no emailaddress is specified, fill the $cc variable for that scenario.<br />

You will need to fill in the following Infos into the Script:<br />
`$cc` = '' # Use as fallback if the Mailaddress is not set for a certificate, 'fallback@example.com'<br />
`$smtphost` = '' # FQDN or IP of your SMTP Server or Relay, 'smtp.example.com'<br />
`$maildomain` = '' # Add your Maildomain, only needed for the Senderadress which is build from the Systemname and the Maildomain, 'example.com'<br /><br />

#### Script v3 (ReleaseDate: To Be Announced)
The Script will read out all CAs from Active Directory.<br />
Script accesses each CA via Remote Powershell and executes Script v2.<br /><br />

## License
This repository is licensed under the GNU General Public License v3.0. <br />
For more information, see the LICENSE file.<br /><br />

## Disclaimer
This repository is for educational and informational purposes only. <br />
The author assumes no liability for any damages that may arise from the use of the contents of this repository.<br /><br />

## Contributions
Contributions are always welcome! If you find an error or would like to suggest an improvement, please create an issue.<br /><br />
