# Windows CA Notify about Expiring Certificates

This repository contains a collection of scripts that help keep an eye on the certificate expiration on a Windows certification authority (CA). <br /> 

## Usage

The scripts can be run manually. <br />
But it is advised to automate the job with the Windows Task Scheduler. <br />
For full coverage and to have time left to renew expiring certs, run the script via the Task Scheduler every 14 Days.<br />
You can use the System Account for running the Script, but i advise you to an account with read right to the certificate Authority.<br />
<br />
#### Script v1

This Script creates a csv from all certificates that will expire in the next 30 days. <br />
After that it will create a email with all certificates listed in the mailbody as a table. <br />
<br />
#### Script v2
This Script creates a csv from all certificates that will expire in the next 30 days. <br />
After that it will create a email for each unique emailaddress specified in the certificate, with all corresponding certificates listed in the mailbody. <br />
Watch out for certificates where no emailaddress is specified, fill the $cc variable for that scenario.<br />
<br />
## License

This repository is licensed under the GNU General Public License v3.0. <br />
For more information, see the LICENSE file.

## Disclaimer

This repository is for educational and informational purposes only. <br />
The author assumes no liability for any damages that may arise from the use of the contents of this repository.

## Contributions

Contributions are always welcome! If you find an error or would like to suggest an improvement, please create an issue.
