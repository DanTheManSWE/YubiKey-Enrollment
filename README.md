# YubiKey-Enrollment

This is a Work In Progress tool, it comes "as-is" with no warranty or support.
I would not recommend using it in a production environment without fully understanding everything the script does. 

**Pre requisites:**
* PowerShell Version 5
* Active Directory PowerShell module
* Yubikey Manager (https://www.yubico.com/products/services-software/download/yubikey-manager/)
* YubiKey minidriver (https://www.yubico.com/products/services-software/download/smart-card-drivers-tools/)
* Microsoft Enterprise PKI
* Sqlserver PowerShell Module (Install-Module -Name SqlServer -RequiredVersion 21.1.18256)
* Microsoft SQL Server


## Quick Start Guide
* Download Zip-file containing script, config and Resources folder.
* Install the required pre requisites.
* Create templates for YubiKey Smart Card certificate and Enrollment Agent.
* Enroll a Certificate Request Agent cert on the user running the script
* Create a Database in SQL
* Edit config.xml


**config.xml explanation**
* Ykman - Path to YubiKey Manager
* SignerCert - Thumbprint of Enrollment Agent certificate
* CertTemplate - YubiKey Smart Card certificate template
* WriteToAD - if true, PIN, PUK, Smartcard serial will be stored in AD on the selected attribute
* ADAttribute - AD Attribute to store PIN, PUK and Smartcard serial
* DeleteOTP - If true, both OTP keys will be deleted on the YubiKey
* DeleteCachedSC - If true, all cached smart card certs will be deleted under the user account running the script
* WriteLog - Writes to logfile
* WriteToDB - Stores pin, puk, yubikey serial etc as an Encrypted string in the specified database
* DBServer - SQL Instance, for example Server01\SQLEXPRESS
* Database - Database name
* DBTable - Database table name
* UseDatabase - Check for YubiKey info in the specified database
* UseAD - Check for YubiKey info in the specified AD attribute



If WriteToAD is set to true, make sure that the user has write-access to the specified AD Attribute. 
**Make sure the selected AD attribute can hold at least 1024 chars, preferrably 2048**

List user attributes and their limit (rangeupper)

`
$schema =[DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()
$schema.FindClass('user').optionalproperties | select name,rangeupper
`

It might also be good to use an attribute where you can set the confidentiality bit, so it can be hidden from normal users. 

The randomized PIN code is padded with two leading zeroes since a lot of users are not comfortable remembering 6-digit pins. 

That can be changed if needed by changing the maximum value to 999999 instead of 9999

`$pin = (Get-Random -Minimum 1000 -Maximum 9999).ToString("000000")`


## Enrolling a new YubiKey
Type the username, press Enroll Smartcard.
When the PIN request is displayed, input the PIN from YubiKey Info

![](https://user-images.githubusercontent.com/16286830/54113995-b7b45300-43e9-11e9-89aa-20d8b90c40bb.png)


If everything is ok you should get a green light on the YubiKey image and all information will be displayed in the YubiKey Info and Log. 
![](https://user-images.githubusercontent.com/16286830/54114018-c3a01500-43e9-11e9-8dad-2427c49da971.png)
