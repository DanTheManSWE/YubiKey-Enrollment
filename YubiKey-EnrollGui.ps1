<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

if (Test-Path "$PSScriptRoot\config.xml"){
	$configFile = "$PSScriptRoot\config.xml"
	$script:config = [xml](Get-Content $configFile)
}else{
	Write-Error "Configfile not found, aborting!"
	return
}

$script:selectedKey = ""
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '570,510'
$Form.MinimumSize                = '570,510'
$Form.text                       = "YubiKey Enrollment"
$Form.TopMost                    = $false
$Form.Icon                       = "$PSScriptRoot\Resources\YubiKey-5-NFC.ico"
$Form.StartPosition              = 'CenterScreen'

$userNameTextBox1                = New-Object system.Windows.Forms.TextBox
$userNameTextBox1.multiline      = $false
$userNameTextBox1.width          = 240
$userNameTextBox1.height         = 20
$userNameTextBox1.location       = New-Object System.Drawing.Point(15,30)
$userNameTextBox1.Font           = 'Microsoft Sans Serif,10'

$enrollButton                    = New-Object system.Windows.Forms.Button
$enrollButton.text               = "Enroll Smartcard"
$enrollButton.width              = 115
$enrollButton.height             = 30
$enrollButton.location           = New-Object System.Drawing.Point(280,25)
$enrollButton.Font               = 'Microsoft Sans Serif,10'

$logButton                       = New-Object system.Windows.Forms.Button
$logButton.text                  = "YubiKey Info"
$logButton.width                 = 125
$logButton.height                = 30
$logButton.location              = New-Object System.Drawing.Point(420,25)
$logButton.Font                  = 'Microsoft Sans Serif,10'

$img = [System.Drawing.Image]::Fromfile("$PSScriptRoot\Resources\YubiKey-5-NFC_Default.png")
$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 125
$PictureBox1.height              = 125
$PictureBox1.location            = New-Object System.Drawing.Point(420,68)
$PictureBox1.Image			     = $img
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom

$Groupbox1                       = New-Object system.Windows.Forms.Groupbox
$Groupbox1.height                = 153
$Groupbox1.width                 = 400
$Groupbox1.text                  = "YubiKey Info"
$Groupbox1.location              = New-Object System.Drawing.Point(15,69)

$userNameLabel                   = New-Object system.Windows.Forms.Label
$userNameLabel.text              = "Username"
$userNameLabel.AutoSize          = $true
$userNameLabel.width             = 25
$userNameLabel.height            = 10
$userNameLabel.location          = New-Object System.Drawing.Point(15,15)
$userNameLabel.Font              = 'Microsoft Sans Serif,10'

$serialLabel                     = New-Object system.Windows.Forms.Label
$serialLabel.text                = "Serial:"
$serialLabel.AutoSize            = $true
$serialLabel.width               = 25
$serialLabel.height              = 10
$serialLabel.location            = New-Object System.Drawing.Point(7,18)
$serialLabel.Font                = 'Microsoft Sans Serif,10'

$yubiKeycomboBox1                = New-Object System.Windows.Forms.ComboBox
$yubiKeycomboBox1.Location       = New-Object System.Drawing.Point(64,14)
$yubiKeycomboBox1.Size           = New-Object System.Drawing.Size(100,20)
$yubiKeycomboBox1.DropDownStyle  = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
$yubiKeycomboBox1.Font           = 'Microsoft Sans Serif,10'
#$yubiKeycomboBox1.Sorted         = $True
$yubiKeycomboBox1.Visible        = $false

$deleteButton                    = New-Object system.Windows.Forms.Button
$deleteButton.text               = "Revoke"
$deleteButton.width              = 115
$deleteButton.height             = 30
$deleteButton.location           = New-Object System.Drawing.Point(265,12)
$deleteButton.Font               = 'Microsoft Sans Serif,10'
$deleteButton.Enabled            = $false

$serialOutputLabel               = New-Object system.Windows.Forms.Label
$serialOutputLabel.AutoSize      = $true
#$serialOutputLabel.Text          = "90123456"
$serialOutputLabel.width         = 25
$serialOutputLabel.height        = 10
$serialOutputLabel.location      = New-Object System.Drawing.Point(64,18)
$serialOutputLabel.Font          = 'Microsoft Sans Serif,10'
$serialOutputLabel.Visible       = $true

$pinLabel                        = New-Object system.Windows.Forms.Label
$pinLabel.text                   = "Pin:"
$pinLabel.AutoSize               = $true
$pinLabel.width                  = 25
$pinLabel.height                 = 10
$pinLabel.location               = New-Object System.Drawing.Point(7,40)
$pinLabel.Font                   = 'Microsoft Sans Serif,10'

$pinOutputLabel                  = New-Object system.Windows.Forms.Label
#$pinOutputLabel.Text             = "001234"
$pinOutputLabel.AutoSize         = $true
$pinOutputLabel.width            = 25
$pinOutputLabel.height           = 10
$pinOutputLabel.location         = New-Object System.Drawing.Point(64,40)
$pinOutputLabel.Font             = 'Microsoft Sans Serif,10'

$pukLabel                        = New-Object system.Windows.Forms.Label
$pukLabel.text                   = "Puk:"
$pukLabel.AutoSize               = $true
$pukLabel.width                  = 25
$pukLabel.height                 = 10
$pukLabel.location               = New-Object System.Drawing.Point(7,62)
$pukLabel.Font                   = 'Microsoft Sans Serif,10'

$pukOutputLabel                  = New-Object system.Windows.Forms.Label
#$pukOutputLabel.Text             = "12345678"
$pukOutputLabel.AutoSize         = $true
$pukOutputLabel.width            = 25
$pukOutputLabel.height           = 10
$pukOutputLabel.location         = New-Object System.Drawing.Point(64,62)
$pukOutputLabel.Font             = 'Microsoft Sans Serif,10'

$certLabel                       = New-Object system.Windows.Forms.Label
$certLabel.text                  = "Cert:"
$certLabel.AutoSize              = $true
$certLabel.width                 = 25
$certLabel.height                = 10
$certLabel.location              = New-Object System.Drawing.Point(7,84)
$certLabel.Font                  = 'Microsoft Sans Serif,10'

$certOutputLabel                  = New-Object system.Windows.Forms.Label
#$certOutputLabel.Text            = "ABCDEFG123456425324325432"
$certOutputLabel.AutoSize         = $true
$certOutputLabel.width            = 25
$certOutputLabel.height           = 10
$certOutputLabel.location         = New-Object System.Drawing.Point(64,84)
$certOutputLabel.Font             = 'Microsoft Sans Serif,10'

$adCertLabel                       = New-Object system.Windows.Forms.Label
$adCertLabel.text                  = "AD Cert:"
$adCertLabel.AutoSize              = $true
$adCertLabel.width                 = 25
$adCertLabel.height                = 10
$adCertLabel.location              = New-Object System.Drawing.Point(7,106)
$adCertLabel.Font                  = 'Microsoft Sans Serif,10'

$adCertOutputLabel                  = New-Object system.Windows.Forms.Label
#$adCertOutputLabel.Text            = "A23144ADSFG123456425324325432"
$adCertOutputLabel.AutoSize         = $true
$adCertOutputLabel.width            = 25
$adCertOutputLabel.height           = 10
$adCertOutputLabel.location         = New-Object System.Drawing.Point(64,106)
$adCertOutputLabel.Font             = 'Microsoft Sans Serif,10'

$dateLabel                       = New-Object system.Windows.Forms.Label
$dateLabel.text                  = "Date:"
$dateLabel.AutoSize              = $true
$dateLabel.width                 = 25
$dateLabel.height                = 10
$dateLabel.location              = New-Object System.Drawing.Point(7,128)
$dateLabel.Font                  = 'Microsoft Sans Serif,10'

$dateOutputLabel                  = New-Object system.Windows.Forms.Label
#$dateOutputLabel.Text            = "den 22 maj 2019 08:05:18"
$dateOutputLabel.AutoSize         = $true
$dateOutputLabel.width            = 25
$dateOutputLabel.height           = 10
$dateOutputLabel.location         = New-Object System.Drawing.Point(64,128)
$dateOutputLabel.Font             = 'Microsoft Sans Serif,10'

$logLabel                        = New-Object system.Windows.Forms.Label
$logLabel.text                   = "Log"
$logLabel.AutoSize               = $true
$logLabel.width                  = 25
$logLabel.height                 = 10
$logLabel.location               = New-Object System.Drawing.Point(15,230)
$logLabel.Font                   = 'Microsoft Sans Serif,10'

$logTextBox                      = New-Object system.Windows.Forms.TextBox
$logTextBox.multiline            = $true
$logTextBox.width                = 530
$logTextBox.height               = 240
$logTextBox.Anchor               = 'top,right,bottom,left'
$logTextBox.location             = New-Object System.Drawing.Point(15,250)
$logTextBox.Font                 = 'Consolas,10'
$logTextBox.ScrollBars           = "Vertical"

$Form.controls.AddRange(@($userNameTextBox1,$enrollButton,$userNameLabel,$logButton,$PictureBox1,$Groupbox1,$logTextBox,$logLabel))
$Groupbox1.controls.AddRange(@($certOutputLabel,$certLabel,$serialLabel,$pinLabel,$pukLabel,$yubiKeycomboBox1,$serialOutputLabel,$pinOutputLabel,$pukOutputLabel,$dateLabel,$dateOutputLabel,$adCertLabel,$adCertOutputLabel,$deleteButton))

$enrollButton.Add_Click({Confirm-ADUser})
$logButton.Add_Click({Get-YubiKeyInfo})
$deleteButton.Add_Click({Remove-SQLData})

$userNameTextBox1.Add_KeyDown({
    if ($_.KeyCode -eq "Enter"){
        Confirm-ADUser
    }
})

    
$yubiKeycomboBox1.Add_SelectedIndexChanged(
    { 
        $x = $yubiKeycomboBox1.SelectedItem
        $serialOutputLabel.Text = $x.Serial
        $certOutputLabel.Text = $x.Cert
        $pinOutputLabel.Text = $x.Pin
        $pukOutputLabel.Text = $x.Puk
        $dateOutputLabel.Text = $x.Date
        $script:selectedKey = $yubiKeycomboBox1.SelectedItem
    }
)

#Write your logic code here

$script:logFile = "$PSScriptRoot\Logs\YubiKey_" + [string](get-date -Format yyyyMMdd) + ".log"

function New-LogEntry($entry, $first){
    [Boolean]$writeLog = [System.Convert]::ToBoolean($config.Configuration.WriteLog)
    if($writeLog){
        $separator = "********************************"
        if ($first){
            $logDate = Get-Date -Format F
            $separator | Out-File $logFile -Append
            $logDate | Out-File $logFile -append
        }
        $entry | Out-File $logFile -Append
    }

}

function Get-SavedCredential(){
    $CurrentUser = [Environment]::UserName
	$script:CredsFile = $PSScriptRoot + "\Keys\" + $CurrentUser + "_Key.txt"
	$FileExists = Test-Path $CredsFile
		if  ($FileExists -eq $false) {
            $EncKey = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a 16-32 character encryption key', 'Enryption key', "")
            if (($EncKey.Length -lt 16) -or ($EncKey.Length -gt 32)){
                Set-YubiImage Red
                $logTextBox.Text = "Invalid encryption key, must be between 16 and 32 characters!"
                return
            }
            $SecureEncKey = ConvertTo-SecureString $EncKey.ToString() -AsPlainText -Force
            $SecureEncKey | ConvertFrom-SecureString | Out-File $CredsFile
            $password = get-content $CredsFile | convertto-securestring
			$Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist "dummy",$password            

		}
		else{
			Write-Host 'Using your stored credential file' -ForegroundColor Green
			$password = get-content $CredsFile | convertto-securestring
			$Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist "dummy",$password
		}
		return $Credential
}


function Set-YubiImage($color)
{
    $imgpath = "$PSScriptRoot\Resources\YubiKey-5-NFC_" + $color + ".png"
    $img = [System.Drawing.Image]::Fromfile($imgpath)
    $PictureBox1.Image = $img
} 

function Set-Key() {
	param([string]$string)
	$length = $string.length
	$pad = 32-$length
	if (($length -lt 16) -or ($length -gt 32)) {Throw "String must be between 16 and 32 characters"}
	$encoding = New-Object System.Text.ASCIIEncoding
	$bytes = $encoding.GetBytes($string + "0" * $pad)
	return $bytes
}


function Set-EncryptedData {
	param($key,[string]$plainText)
	$securestring = new-object System.Security.SecureString
	$chars = $plainText.toCharArray()
	foreach ($char in $chars) {$secureString.AppendChar($char)}
	$encryptedData = ConvertFrom-SecureString -SecureString $secureString -Key $key
	return $encryptedData
}


function Get-EncryptedData {
	param($key,$data)
	$data | ConvertTo-SecureString -key $key |
	ForEach-Object {[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($_))}
}

function Check-DBConnection(){
    $Instance = $config.Configuration.DBServer
    $Database = $config.Configuration.Database
    $db = Get-SqlDatabase -ServerInstance $Instance -Name $Database

    if(!($db)){
        Write-Host "No Connection to Database! Check connection"
        Set-YubiImage Red
        $logTextBox.Text = "No Connection to Database! Check connection"
        $logTextBox.Text += "`n" | Out-String
        return $false
    }

    $logTextBox.Text += "Connection to $Database on Instance: $Instance established."
    $logTextBox.Text += "`n" | Out-String
    ## Create DBTable if it doesn't exist
    Create-DBTable
    return $true
}

function Create-DBTable{
    $Instance = $config.Configuration.DBServer
    $Database = $config.Configuration.Database
    $Table = $config.Configuration.DBTable
    $tableQuery = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'$Table'"
    $sqlTable = Invoke-Sqlcmd -ServerInstance $Instance -Database $Database -Query $tableQuery
    if ($sqlTable.TABLE_NAME -eq $Table){
        Write-Host "Table: $Table already exist"
    }else{
        Write-Host "Creating table: $Table in Database: $Database"
        $logTextBox.Text += "Creating table: $Table in Database: $Database"
        $logTextBox.Text += "`n" | Out-String
        $createtableQuery = "CREATE TABLE $Table (ID INT IDENTITY(1,1) NOT NULL, Username NVARCHAR(256), EncryptedInfo NVARCHAR(MAX), Date DATETIME2(7))"
        $setPrimaryKeyQuery = "ALTER TABLE $Table ADD CONSTRAINT PK_$Table PRIMARY KEY(Id)"
        Invoke-Sqlcmd -ServerInstance $Instance -Database $Database -Query $createtableQuery
        Invoke-Sqlcmd -ServerInstance $Instance -Database $Database -Query $setPrimaryKeyQuery
    }

}

function Clear-YubiInfoBox(){
#    $userOutputLabel.Text = ""
    $serialOutputLabel.Text = ""
    $pinOutputLabel.Text = ""
    $pukOutputLabel.Text = ""
    $certOutputLabel.Text = ""
    $dateOutputLabel.Text = ""
    $adCertOutputLabel.Text = ""
}

function Get-ADCertificate($userName){
    if (!(Get-Module ActiveDirectory)){
        $logTextBox.Text += "`n" | Out-String
        $logTextBox.Text += "Active Directory Module not loaded!" | Out-String
        $adCertOutputLabel.Text = "N/A"
        Set-YubiImage Red
        return
    }else{
        $user = Get-ADUser -Filter {SamAccountName -eq $userName} -Properties displayName, Certificates
    }

    $YubiCerts = @()
    $SortedCerts = @()
    $templateName = $config.Configuration.CertTemplate
    foreach ($usercert in $user.Certificates){
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $usercert
        $temp = $cert.Extensions | Where-Object {$_.Oid.Value -eq "1.3.6.1.4.1.311.20.2"}
        if (!$temp) {
                $temp = $cert.Extensions | Where-Object {$_.Oid.Value -eq "1.3.6.1.4.1.311.21.7"}
        }
        if (!$temp){
            $adCertOutputLabel.Text = "N/A"
            return
        }
        if($temp.Format(1) -match $templateName){
            Write-Host "Found matching cert in AD! $($temp.Format(1))"
            $YubiCerts += $cert
        }
    }
    
    ## Sort Certificates based on SerialNumber and only return latest one
    $SortedCerts = $YubiCerts | Sort-Object -Descending SerialNumber
    if ($SortedCerts){
        $SerialNumber = $SortedCerts[0].SerialNumber | Out-String
        $adCertOutputLabel.Text = $SerialNumber.Trim()
        $logTextBox.Text += "`n" | Out-String
        $logTextBox.Text += "Latest YubiKey cert in AD: " + $SerialNumber
    }else{
        $adCertOutputLabel.Text = "N/A"
    }    
}

function Remove-ADCertificate($userName, $serialNumber){
    if (!(Get-Module ActiveDirectory)){
        $logTextBox.Text += "`n" | Out-String
        $logTextBox.Text += "Active Directory Module not loaded!" | Out-String
        $adCertOutputLabel.Text = "N/A"
        Set-YubiImage Red
        return
    }else{
        $user = Get-ADUser -Filter {SamAccountName -eq $userName} -Properties displayName, Certificates
    }
    
    foreach ($usercert in $user.Certificates){
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $usercert
	    if ($cert.SerialNumber -eq $serialNumber){
            Write-Host "Removing Certificate with serialnumber: $($serialNumber) from AD"
            Set-ADUser -Identity $userName -Certificates @{ Remove = $cert}
        }
    }
        

}

function Get-SQLData{
    $yubiKeycomboBox1.DataSource = $null
    $yubiKeycomboBox1.Visible = $false
    $serialOutputLabel.Visible = $true
    $Instance = $config.Configuration.DBServer
    $Database = $config.Configuration.Database
    $Table = $config.Configuration.DBTable
    $db = Get-SqlDatabase -ServerInstance $Instance -Name $Database

    if(!($db)){
        Write-Host "No Connection to Database! Check connection"
        Set-YubiImage Red
        $logTextBox.Text = "No Connection to Database! Check connection"
        $logTextBox.Text += "`n" | Out-String
        return
    }

    $userName = $userNameTextBox1.Text
    $Query = "select * from dbo.$Table WHERE USERNAME='$userName'"
    $EncryptedKeys = Invoke-Sqlcmd -ServerInstance $Instance -Database $Database -Query $Query
    if(!($EncryptedKeys)){
        $logTextBox.Text = "No YubiKey information found on User: $userName in database: $Database"
        $deleteButton.Enabled = $false
        Set-YubiImage Red
        return
    }
    $DecryptedKeys = New-Object System.Data.Datatable
    $DecryptedKeys.Columns.Add("Serial") | Out-Null
    $DecryptedKeys.Columns.Add("Pin") | Out-Null
    $DecryptedKeys.Columns.Add("Puk") | Out-Null 
    $DecryptedKeys.Columns.Add("Cert") | Out-Null
    $DecryptedKeys.Columns.Add("Date") | Out-Null
    $DecryptedKeys.Columns.Add("Username") | Out-Null
    $DecryptedKeys.Columns.Add("ID") | Out-Null
    
    Foreach ($Encryptedkey in $EncryptedKeys){
        $DecryptedInfo = (Get-EncryptedData -key $key -data $Encryptedkey.EncryptedInfo).Split(";")
        $DecryptedKeys.Rows.Add($DecryptedInfo[0],$DecryptedInfo[1],$DecryptedInfo[2],$DecryptedInfo[3],$EncryptedKey.Date, $EncryptedKey.Username, $EncryptedKey.ID) | Out-Null
    }
    ## Check if Datatable has more than one rows, if so unhide Combobox
    if($DecryptedKeys.Rows.Count -gt 1){
        $yubiKeycomboBox1.Visible = $true
        $serialOutputLabel.Visible = $false
    }

    ## Sort datatable
    $DecryptedKeys.DefaultView.Sort = "Cert DESC"
    $DecryptedKeys = $DecryptedKeys.DefaultView.ToTable()

    $yubiKeycomboBox1.DataSource = $DecryptedKeys
    $yubiKeycomboBox1.DisplayMember = "Serial"
    $yubiKeycomboBox1.ValueMember = "Serial"

    $yubiInfo = $DecryptedKeys.Rows[0].ItemArray

    $logTextBox.Text += "`n" | Out-String
    $logTextBox.Text += "If YubiKey is lost make sure to revoke certificate: $($yubiInfo[3])" | Out-String
    $logTextBox.Text += "`n" | Out-String
    $logTextBox.Text += "If user has renewed the certificate, this may not be the correct serialnumber.`nCheck AD Cert instead." | Out-String

    ## Get Certificate information from AD
    Get-ADCertificate $userName
    $deleteButton.Enabled = $true

    Set-YubiImage Green
}

function Get-YubiKeyInfo(){
    Clear-YubiInfoBox
    $logTextBox.Text = ""
    $UserName = $userNameTextBox1.Text
    [Boolean]$useDatabase = [System.Convert]::ToBoolean($config.Configuration.UseDatabase)
    [Boolean]$useAD = [System.Convert]::ToBoolean($config.Configuration.UseAD)   
    if($UserName){
        if($useDatabase){
            $logTextBox.Text += "Fetching data from Database"
            $sqlData = Get-SQLData
            ## return $sqlData
        }
        if($useAD){
            $adData = Get-ADData
            return $adData
        }
        ## Read info from YubiKey
        $ykman = $config.Configuration.Ykman
        $pivinfo = & $ykman piv info
        if($LASTEXITCODE -ne 0){
	        Write-Host "No YubiKey detected!" -ForegroundColor Red
            $pivinfo = "No YubiKey detected, unable to read data."
        }
        $logTextBox.Text += "`n" | Out-String
        $logTextBox.Text += "`n" | Out-String
        $logTextBox.Text += "Inserted YubiKey Piv Info" | Out-String
        $logTextBox.Text += "`n" | Out-String
        $logTextBox.Text += $pivinfo | Out-String
        $logTextBox.Text += "`n`n" | Out-String



    }else{
        $logTextBox.Text += "Input username!" | Out-String
        Set-YubiImage Red
        return
    }
}

function Get-ADData(){
    $adInfo = @()
    $adAttribute = $config.Configuration.ADAttribute
    if (!(Get-Module ActiveDirectory)){
        $logTextBox.Text += "`n" | Out-String
        $logTextBox.Text += "Active Directory Module not loaded!" | Out-String
        Set-YubiImage Red
        return
    }else{
        $user = Get-ADUser -Filter {SamAccountName -eq $userName} -Properties $adAttribute, displayName, Certificates
    }
    $adInfo += $user | Select-Object -ExpandProperty $adAttribute
    if ($user){
        ## Get certificate serialnumber from AD if it exist on the user
        if($adInfo[0] -match "(YubiKey: )(.+)"){
            $logTextBox.Text = "Fetching encrypted data from AD for user: $($user.DisplayName) ($userName)"
            $encryptedYubi = $Matches[2]
            $decryptedYubi = Get-EncryptedData -data $encryptedYubi -key $key
            if ($decryptedYubi){
                $yubiInfo = $decryptedYubi -split ";"
                #$userOutputLabel.Text = $userName
                $serialOutputLabel.Text = $yubiInfo[0]
                $pinOutputLabel.Text = $yubiInfo[1]
                $pukOutputLabel.Text = $yubiInfo[2]
                $certOutputLabel.Text = $yubiInfo[3]
                $logTextBox.Text += "`n" | Out-String
                $logTextBox.Text += "If YubiKey is lost make sure to revoke certificate: $($yubiInfo[3])" | Out-String
                $logTextBox.Text += "`n" | Out-String
                $logTextBox.Text += "If user has renewed the certificate, this may not be the correct serialnumber.`nCheck ADCert instead." | Out-String
                Get-ADCertificate $userName
                Set-YubiImage Green
                return $yubiInfo
            }
            else{
                $logTextBox.Text = "Invalid encryption key, unable to encrypt data!"
                $logTextBox.Text += "`n" | Out-String
                $logTextBox.Text += "Delete $CredsFile and restart program!"
                Set-YubiImage Red
            }
        }else{
            Clear-YubiInfoBox
            $entry = "No YubiKey information found on User: $($user.DisplayName) ($userName)"
            Write-Warning $entry
            $logTextBox.Text = $entry | Out-String
            Set-YubiImage Red
        }
    }else{
        $entry = "User $UserName does not exist in AD, verify username!"
        $logTextBox.Text += $entry | Out-String
        Write-Warning $entry
        New-LogEntry $entry $true
        Set-YubiImage Red
        return
    }

}


function Confirm-ADUser(){
    $logTextBox.Text = ""
    $userName = $userNameTextBox1.Text
    Clear-YubiInfoBox
    $adAttribute = $config.Configuration.ADAttribute
    if($userName){
        if (!(Get-Module ActiveDirectory)){
            $logTextBox.Text += "Active Directory Module not loaded!" | Out-String
            Set-YubiImage Red
            return
        }else{
            $user = Get-ADUser -Filter {SamAccountName -eq $UserName} -Properties $adAttribute, Certificates
        }
	}else{
        $logTextBox.Text += "Input username!" | Out-String
        Set-YubiImage Red
		return
    }
	

    if($user){
        Write-Host "User: $($user.GivenName) $($user.Surname) verified in Active Directory" -ForegroundColor Green
        $logTextBox.Text += "User: $($user.GivenName) $($user.Surname) verified in Active Directory" | Out-String
        Get-ADCertificate $userName
        Set-YubiImage Default
    }else{
        $entry = "User $UserName does not exist in AD, verify username!"
        $logTextBox.Text += $entry | Out-String
        Write-Error $entry
        New-LogEntry $entry $true
        Set-YubiImage Red
        return
    }
    $adInfo = @()
    $adInfo += $user | Select-Object -ExpandProperty $adAttribute

     if($adInfo[0] -match "(YubiKey: )(.+)"){
        Get-YubiKeyInfo
        $msgBoxInput = [System.Windows.MessageBox]::Show('YubiKey already enrolled for user, do you want to overwrite information?','YubiKey Warning','YesNo','Warning')

        switch ($msgBoxInput){
            'Yes'
            {
                Write-Host "Enrolling User"
                $logTextBox.Text += "`n" | out-string
                $logTextBox.Text += "`n" | out-string
                Write-YubiKey $user $userName
            }
            'No'
            {
                Write-Warning "Certificate not enrolled!"
                $logTextBox.Text += "`n" | out-string
                $logTextBox.Text += "`n" | out-string
                $logTextBox.Text += "Certificate not enrolled!"
                Set-YubiImage Red
            }

        }    
     
     }else{
        Write-YubiKey $user $userName
     }

}

function Write-SQLData($userName, $encryptedInfo){
    $Date = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fffffff")
    $Instance = $config.Configuration.DBServer
    $Database = $config.Configuration.Database
    $Table = $config.Configuration.DBTable

    $insertQuery = "INSERT INTO $Table (Username, EncryptedInfo, Date) VALUES ('$userName','$encryptedInfo','$Date')"
    Invoke-SqlCmd -ServerInstance $Instance -Database $Database -Query $insertQuery

}

function Remove-SQLData(){
    $keyID = $selectedKey.ID
    $certSerial = $selectedKey.Cert
    $Instance = $config.Configuration.DBServer
    $Database = $config.Configuration.Database
    $Table = $config.Configuration.DBTable
    $DeleteQuery = "DELETE FROM dbo.$Table WHERE ID = $keyID"
    $UserName = $userNameTextBox1.Text

    if($keyID){
        $msgBoxInputDelete = [System.Windows.MessageBox]::Show("Are you sure you want to delete database-entry with id: $($keyID) and revoke certificate with serialnumber: $($certSerial)?",'Delete Warning','YesNo','Warning')

        switch ($msgBoxInputDelete){
            'Yes'
            {
                Write-Host "Deleting row with ID: $keyID"
                Invoke-Sqlcmd -ServerInstance $Instance -Database $Database -Query $DeleteQuery
                Revoke-Certificate $certSerial
                Remove-ADCertificate $UserName $certSerial
                Set-YubiImage Default
                Get-YubiKeyInfo

            }
            'No'
            {
                Write-Warning "Delete aborted!"
                Set-YubiImage Red
            }

        }
    }else{
        $msgBoxNoUser = [System.Windows.MessageBox]::Show('No YubiKey selected!','No YubiKey','Ok','Warning')
    }

}

function Revoke-Certificate($serialNumber){
    $CertAdmin = New-Object -ComObject CertificateAuthority.Admin
    $Reason = 0
    $CAConfig = $Config.Configuration.CAServer + "\" + $Config.Configuration.CAName
    $RevokationDate = (Get-Date).ToUniversalTime()
    Write-Host "Revoking $($serialNumber) using $($CAConfig) with reason: $Reason and Date: $RevokationDate"
    $CertAdmin.RevokeCertificate($CAConfig,$serialNumber,$Reason,$RevokationDate)

}

function Write-YubiKey($user, $userName){
    Set-YubiImage Default
    $yubiKeycomboBox1.Visible = $false
    $serialOutputLabel.Visible = $true
    Clear-YubiInfoBox
    $ykman = $config.Configuration.Ykman

    
    $templateName = $config.Configuration.CertTemplate
    $signerCert = $config.Configuration.SignerCert
    

    Write-Host "Resetting YubiKey" -Foregroundcolor Yellow
    $logTextBox.Text += "Resetting YubiKey`n`n" | Out-String

    & $ykman piv reset -f | Out-Null

    ## Delete OTP Passwords
    [Boolean]$deleteOTP = [System.Convert]::ToBoolean($config.Configuration.DeleteOTP)
    if($deleteOTP){
	Write-Host "Deleting OTP!" -Foregroundcolor Yellow
        & $ykman otp delete 1 -f | Out-Null
        & $ykman otp delete 2 -f | Out-Null
    }

    if($LASTEXITCODE -ne 0){
	    Write-Host "Error resetting YubiKey, make sure YubiKey is inserted" -ForegroundColor Red
        $logTextBox.Text += "Error resetting YubiKey, make sure YubiKey is inserted" | Out-String
        Set-YubiImage Red
	    return
    }

    $serial = & $ykman list -s
    $pin = (Get-Random -Minimum 1000 -Maximum 9999).ToString("000000")
    $puk = (Get-Random -Minimum 10000000 -Maximum 99999999).ToString("00000000")

 
    & $ykman piv access change-puk -p 12345678 -n $puk | Out-Null
    if($LASTEXITCODE -ne 0){
	    Write-Host "Error setting new PUK" -ForegroundColor Red
        $logTextBox.Text += "Error setting new PUK!" | Out-String
        Set-YubiImage Red
	    return
    }
    & $ykman piv access change-pin -P 123456 -n $pin | Out-Null
    if($LASTEXITCODE -ne 0){
	    Write-Host "Error setting new PIN" -ForegroundColor Red
        $logTextBox.Text += "Error setting new PIN!" | Out-String
        Set-YubiImage Red
	    return
    }

    Write-Host "`n"

    $pinOutputLabel.Text = $pin
    #$userOutputLabel.Text = $userName
    $serialOutputLabel.Text = $serial
    $pukOutputLabel.Text = $puk

    $pinEntry = "User: $($userName)`nSerial: $($serial)`nPin: $($pin)`nPuk: $($puk)"
    Write-Host $pinEntry -ForegroundColor Green
    Write-Host "`n"


    try{	
        #Remove any existing cached Smart Card cert before enrolling new
	[Boolean]$deleteCachedSC = [System.Convert]::ToBoolean($config.Configuration.DeleteCachedSC)
        if($deleteCachedSC){
		$oldSmartcardcert = Get-ChildItem Cert:\CurrentUser\my | where-object {$_.EnhancedKeyUsageList -like "*Smart Card*"} 
		Write-Host "Removing cached smart card certs!" -Foregroundcolor Yellow
        	$oldSmartcardcert | Remove-Item
	}
        $PKCS10 = New-Object -ComObject X509Enrollment.CX509CertificateRequestPkcs10
        $PKCS10.InitializeFromTemplateName(0x1,$templateName)

        Write-Host "Writing Private Key, if Touch-Policy is enabled tap the button when it starts flashing slowly!`n" -ForegroundColor Yellow
        $logTextBox.Text += "`n" | Out-String
        $logTextBox.Text += "Writing Private Key, if Touch-Policy is enabled tap the button when it starts flashing slowly!`n" | Out-String

        $PKCS10.Encode()
        $pkcs7 = New-Object -ComObject X509enrollment.CX509CertificateRequestPkcs7
        $pkcs7.InitializeFromInnerRequest($pkcs10)

        $pkcs7.RequesterName = $userName

        $signer = New-Object -ComObject X509Enrollment.CSignerCertificate
        $signer.Initialize(0,0,0xc,$signerCert)

        $pkcs7.SignerCertificate = $signer

        $Request = New-Object -ComObject X509Enrollment.CX509Enrollment
        $Request.InitializeFromRequest($pkcs7)
        $Request.Enroll()
        $smartCardCerts = @()
        $smartCardCerts = Get-ChildItem Cert:\CurrentUser\my | where-object {$_.EnhancedKeyUsageList -like "*Smart Card*"} 
        $smartCardCerts = $smartCardCerts | Sort-Object NotBefore -Descending
        ##$smartCardInfo = "YubiKey: $serial" + ", " + "Cert thumbprint: $($smartcardcert.thumbprint)"

        ## Return latest smart card serialnumber
        $smartCardCert = $smartCardCerts[0]
        $certSerial = $smartCardCert.SerialNumber

        $certOutputLabel.Text = $certSerial | out-string

        $plainText = $serial + ";" + $pin + ";" + $puk + ";" + $certSerial

        $encryptedText = Set-EncryptedData -key $key -plainText $plainText
        $stringToAD = "YubiKey: " + $encryptedText

        $adAttribute = $config.Configuration.ADAttribute
		[Boolean]$adWrite = [System.Convert]::ToBoolean($config.Configuration.WriteToAD)
        [Boolean]$dbWrite = [System.Convert]::ToBoolean($config.Configuration.WriteToDB)

        if ($adWrite){
			Write-Host "Writing Encrypted string to AD-attribute: $adAttribute" -foregroundcolor Yellow
			$logTextBox.Text += "`n" | Out-String
			$logTextBox.Text += "Writing Encrypted string to AD-attribute: $adAttribute`n" | Out-String
            $user | Set-ADUser -Replace @{$adAttribute = $stringToAD}        
        }else{
			Write-Host "Yubikey info not written to AD, make sure to record information somewhere safe!" -foregroundcolor Yellow
			$logTextBox.Text += "`n" | Out-String
			$logTextBox.Text += "Yubikey info not written to AD, make sure to record information somewhere safe!`n" | Out-String
        }
        
        if($dbWrite){
            $dbConnection = Check-DBConnection
            if($dbConnection){
            $Database = $config.Configuration.Database
                Write-Host "Writing Encrypted string to SQL DB." -foregroundcolor Yellow
                $logTextBox.Text += "`n" | Out-String
                $logTextBox.Text += "Writing Encrypted string to SQL DB: $Database`n" | Out-String
                Write-SQLData $userName $encryptedText
            }else{
                $logTextBox.Text += "Unable to Write to DB, no connection!" | Out-String
                Set-YubiImage Red
            }

        }
	
	#Remove cached smart card cert after enrollment
        $smartcardcert | Remove-Item
    }
    catch [Exception]
    {
        $entry = "***Error*** " + $_.Exception.Message
        New-LogEntry $entry $true
        Write-Error $entry
        $logTextBox.Text += $entry | Out-String
        Set-YubiImage Red
        return
    }
    Write-Host "Certificate enrolled`n`n" -ForegroundColor Green
    $logTextBox.Text += "`n" | Out-String
    $logTextBox.Text += "Certificate enrolled`n`n" | Out-String

    
    $pivinfo = & $ykman piv info
    if($LASTEXITCODE -ne 0){
	    Write-Host "Error getting info from YubiKey!" -ForegroundColor Red
        Set-YubiImage Red
	    return
    }
    Set-YubiImage Green
    $logTextBox.Text += "`n" | Out-String
    $logTextBox.Text += $pivinfo | Out-String
    $logTextBox.Text += "`n`n" | Out-String

    

    New-LogEntry "User: $($UserName)" $true
    New-LogEntry "Serial: $($serial)" $false
    New-LogEntry "Pin: $($pin)" $false
    New-LogEntry "Puk: $($puk)" $false
    New-LogEntry $adInfo $false
}

##Loading Active Directory Module
if (!(Get-Module ActiveDirectory)){
    if (Get-Module -ListAvailable -Name ActiveDirectory){
        $logTextBox.Text = "Loading Active Directory Module" | Out-String
        Import-Module ActiveDirectory
    }else{
        $logTextBox.Text += "Active Directory Module not found!" 
        $logTextBox.Text += "`n" | Out-String 
    }
    if (!(Get-Module ActiveDirectory)){
        $logTextBox.Text += "Active Directory Module not loaded!" | Out-String
        Set-YubiImage Red
    }else{
        $logTextBox.Text += "Active Directory Module loaded" | Out-String
    }
}



##Create Log-folder if needed
if(!(Test-Path $PSScriptRoot\Logs)){
	Write-Host "Creating Log dir"
	New-Item $PSScriptRoot\Logs -Type Directory | Out-Null
}

##Create Keys-folder if needed
if(!(Test-Path $PSScriptRoot\Keys)){
	Write-Host "Creating Keys dir"
	New-Item $PSScriptRoot\Keys -Type Directory | Out-Null
}

$Credentials = Get-SavedCredential
if($Credentials){
    $key = Set-Key ($Credentials.GetNetworkCredential().Password).ToString()
}

[Boolean]$useDatabase = [System.Convert]::ToBoolean($config.Configuration.UseDatabase)

if($useDatabase){
    $dbConnection = Check-DBConnection
    if(!($dbConnection)){
        Set-YubiImage Red
    }else{

    }
}



[void]$Form.ShowDialog()