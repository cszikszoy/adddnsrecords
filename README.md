 # adddnsrecords
 PowerShell script to add DNS records (A, CNAME, PTR) to a Windows DNS server
 ## Arguments
 ### Common
| Syntax | Description |
| - | - |
| &#x2011;Debug | *[Optional]* Print extra information |
| &#x2011;DryRun | *[Optional]* Only print destructive (add / delete) commands, don't actually perform them |
| &#x2011;DnsServer | **[Required]** Target DNS server to add records to |
| &#x2011;RemoteHost | *[Optional]* Run commands on remote host using [PowerShell Implicit Remoting](https://devblogs.microsoft.com/scripting/remoting-the-implicit-way/) <br> *Required if local machine is Linux or doesn't have PS DnsServer module installed* |
| &#x2011;Auth | *[Optional]* Perform commands as a different user <br> *Required if run from Linux or current user lacks privileges on the DNS Server* |
| &#x2011;AuthPwd | *[Optional]* Password for `-Auth` user <br> *Console will prompt if not supplied <br> Takes priority over `-AuthPwdFile`* |
| &#x2011;AuthPwdFile | *[Optional]* File containing password for `-Auth` user |
### CSV Files
| Syntax | Description |
| - | - |
| &#x2011;CsvFile | **[Required]** Path to CSV file containing records to add
#### CSV File Format:
```
name,ip,fqdn,cname
Host 1,192.168.1.1,test1.te.st,test.te.st
Host 2,192.168.1.2,test2.te.st,
...
```
### Excel Files
| Syntax | Description |
| - | - |
| &#x2011;ExcelSheetName | *[Optional]* Name of sheet containing host named ranges (defaults to "Hosts" if missing) |
| &#x2011;ExcelRangeNames | **[Required]** [Regular expression](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.2) to match named ranges containing hosts to add |
#### Named Ranges:
When processing an Excel file for hosts, the script will search for any [named ranges](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64) matching the regex supplied by `-ExcelNamedRanges` on the sheet named `-ExcelSheetName`. Make sure not to include table headers in named ranges. Table format should be (excluding headers... markdown doesn't support tables without headers):
| Host Name / Description | IP Address | Fully Qualified Domain Name (FQDN) | Alias (CNAME) |
| - | - | -| - |
| Host #1 | 1.1.1.1 | hostnumber1.te.st | host1.te.st |
| Host #2 | 2.2.2.2 | host2.te.st| |

## Linux Notes
### Install PowerShell
Follow [Microsoft's documentation](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux?view=powershell-7.2) to install PowerShell on Linux.
### NTLM Authentication
Most (all?) Linux distributions don't support the NTLM authentication protocol without additional packages. 
#### On Ubuntu
This package is in the default repos and can be installed by running the following command:
```
sudo apt install gss-ntlmssp
```
#### On RHEL / CentOS / RockyLinux / Other "EL"'s
Version 7
```
sudo yum -y install epel-release
sudo yum -y update
sudo yum -y install gssntlmssp
```
Version 8
```
sudo dnf -y install gssntlmssp
```
