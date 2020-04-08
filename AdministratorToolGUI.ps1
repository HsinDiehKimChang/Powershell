#######################################################################################
###
### Created by Kim Chang 
### Last update @2020/04/01
### This is for creating a new account or reset password in myDomain
### Graphic interface will let help desk/Junior System administrator easier
###
######################################################################################


function checkADM {

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole] "Administrator")) {
Write-Warning "Hey noob, you forgot to run as administrator!"
Write-Warning "Please open the Exchange Shell console as an administrator and run this script again."
}
else {
Write-Host "Code is running as administrator — go on executing the script..." -ForegroundColor Green
}

}

function duplicateNameCheck{

 Param
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string] $Alias
    )

    $dup=Get-ADUser -Filter "UserPrincipalName -like '$Alias*'"

    if ($dup) { [System.Windows.MessageBox]::Show('Alias has been used! Please enter another Alias.') }
    Else {
            
            creatNewUser
            }
}

function creatNewUsertxt{

$Alias=@()
$FirstName=@()
$LastName=@()
$count='y'


 if (!$InputFirstName.Text) { [System.Windows.MessageBox]::Show('Please enter First name!') }
    Elseif 
        (!$InputLastName.Text) { [System.Windows.MessageBox]::Show('Please enter Last name!') }
    Elseif 
        (!$InputAlias.Text) { [System.Windows.MessageBox]::Show('Please enter Alias!') }
    Else
    {
    #Worklog Entry
    $WorkLog.text += "`r`n Checking duplicate for username..."
    duplicateNameCheck $InputAlias.Text
    }    
    
}

function changePW{

    $defaultPW='BlackBerry123'

    if ($InputFullName.Text) { 
            $fullName=$InputFullName.Text
            $userInfo=Get-ADUser -Filter "Name -like '$fullName'"
            $Ind=$userInfo.SamAccountName

            $WorkLog.text += "`r`n Reset Username:'$Ind' password..."
            Set-ADAccountPassword -Identity $Ind -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $defaultPW -Force)
            $WorkLog.text += "`r`n '$Ind' password has been changed to '$defaultPW'" 
             }
    Elseif 
        ($InputAlias2.Text) {
        $Ind=$InputAlias2.Text
        $WorkLog.text += "`r`n Reset Username:'$Ind' password..."
        Set-ADAccountPassword -Identity $Ind -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $defaultPW -Force)
        $WorkLog.text += "`r`n '$Ind' password has been changed to '$defaultPW'" 
         }
    Else { [System.Windows.MessageBox]::Show('Please enter Full name or Alias!') }
}



function Import-MailboxContacts{
     Param
        (    
        [Parameter(Mandatory=$true,Position=0)]
         [string]$CSVFileName,
        [Parameter(Mandatory=$true,Position=1)]
         [string]$EmailAddress,
        [Parameter(Mandatory=$true,Position=2)]
         [string]$Username,
        [Parameter(Mandatory=$true,Position=3)]
         [string]$Password,
        [Parameter(Mandatory=$true,Position=4)]
         [string]$Domain,
        [Parameter(Mandatory=$true,Position=5)]
         [bool]$Impersonate,
        [Parameter(Mandatory=$false,Position=6)]
         [string]$EwsUrl,
        [Parameter(Mandatory=$false,Position=7)]
         [string]$EWSManagedApiDLLFilePath,
        [Parameter(Mandatory=$false,Position=8)]
         [bool]$Exchange2007
        )

$ContactMapping=@{
    "First Name" = "GivenName";
    "Middle Name" = "MiddleName";
    "Last Name" = "Surname";
    "Company" = "CompanyName";
    "Department" = "Department";
    "Job Title" = "JobTitle";
    "Business Street" = "Address:Business:Street";
    "Business City" = "Address:Business:City";
    "Business State" = "Address:Business:State";
    "Business Postal Code" = "Address:Business:PostalCode";
    "Business Country/Region" = "Address:Business:CountryOrRegion";
    "Home Street" = "Address:Home:Street";
    "Home City" = "Address:Home:City";
    "Home State" = "Address:Home:State";
    "Home Postal Code" = "Other:Home:PostalCode";
    "Home Country/Region" = "Address:Home:CountryOrRegion";
    "Other Street" = "Address:Other:Street";
    "Other City" = "Address:Other:City";
    "Other State" = "Address:Other:State";
    "Other Postal Code" = "Address:Other:PostalCode";
    "Other Country/Region" = "Address:Other:CountryOrRegion";
    "Assistant's Phone" = "Phone:AssistantPhone";
    "Business Fax" = "Phone:BusinessFax";
    "Business Phone" = "Phone:BusinessPhone";
    "Business Phone 2" = "Phone:BusinessPhone2";
    "Callback" = "Phone:CallBack";
    "Car Phone" = "Phone:CarPhone";
    "Company Main Phone" = "Phone:CompanyMainPhone";
    "Home Fax" = "Phone:HomeFax";
    "Home Phone" = "Phone:HomePhone";
    "Home Phone 2" = "Phone:HomePhone2";
    "ISDN" = "Phone:ISDN";
    "Mobile Phone" = "Phone:MobilePhone";
    "Other Fax" = "Phone:OtherFax";
    "Other Phone" = "Phone:OtherTelephone";
    "Pager" = "Phone:Pager";
    "Primary Phone" = "Phone:PrimaryPhone";
    "Radio Phone" = "Phone:RadioPhone";
    "TTY/TDD Phone" = "Phone:TtyTddPhone";
    "Telex" = "Phone:Telex";
    "Anniversary" = "WeddingAnniversary";
    "Birthday" = "Birthday";
    "E-mail Address" = "Email:EmailAddress1";
    "E-mail 2 Address" = "Email:EmailAddress2";
    "E-mail 3 Address" = "Email:EmailAddress3";
    "Initials" = "Initials";
    "Office Location" = "OfficeLocation";
    "Manager's Name" = "Manager";
    "Mileage" = "Mileage";
    "Notes" = "Body";
    "Profession" = "Profession";
    "Spouse" = "SpouseName";
    "Web Page" = "BusinessHomePage";
    "Contact Picture File" = "Method:SetContactPicture"
}

# CSV File Checks
# Check filename is specified
if (!$CSVFileName)
{
    throw "Parameter CSVFileName must be specified";
}

# Check file exists
if (!(Get-Item -Path $CSVFileName -ErrorAction SilentlyContinue))
{
    throw "Please provide a valid filename for parameter CSVFileName";
}

# Check file has required fields and check if is a single row, or multiple rows
$SingleItem = $false;
$CSVFile = Import-Csv -Path $CSVFileName;
if ($CSVFile."First Name")
{
    $SingleItem = $true;
} else {
    if (!$CSVFile[0]."First Name")
    {
        throw "File $($CSVFileName) must specify at least the field 'First Name'";
    }
}

# Check email address
if (!$EmailAddress)
{
    throw "Parameter EmailAddress must be specified";
}
if (!$EmailAddress.Contains("@"))
{
    throw "Parameter EmailAddress does not appear valid";
}

# Check EWS Managed API available
if (!$EWSManagedApiDLLFilePath)
{
    $EWSManagedApiDLLFilePath = "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
}
if (!(Get-Item -Path $EWSManagedApiDLLFilePath -ErrorAction SilentlyContinue))
{
    throw "EWS Managed API not found at $($EWSManagedApiDLLFilePath). Download from http://www.microsoft.com/download/en/details.aspx?id=28952";
}

# Load EWS Managed API
[void][Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll");

# Create Service Object
if ($Exchange2007)
{
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
} else {
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)
}
# Set credentials if specified, or use logged on user.
if ($Username -and $Password)
{
    if ($Domain)
    {
        $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain);
    } else {
        $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password);
    }
    
} else {
    $service.UseDefaultCredentials = $true;
}


# Set EWS URL if specified, or use autodiscover if no URL specified.
if ($EwsUrl)
{
    $service.URL = New-Object Uri($EwsUrl);
} else {
    try {
        $service.AutodiscoverUrl($EmailAddress);
    } catch {
        throw;
    }
}

# Perform a test - try and get the default, well known contacts folder.

if ($Impersonate)
{
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
}
try {
    $ContactsFolder = [Microsoft.Exchange.WebServices.Data.ContactsFolder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts);
} catch {
    throw;
}

# Add contacts
foreach ($ContactItem in $CSVFile)
{
    # If impersonate is specified, do so.
    if ($Impersonate)
    {
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
    }

    $ExchangeContact = New-Object Microsoft.Exchange.WebServices.Data.Contact($service);
    if ($ContactItem."First Name" -and $ContactItem."Last Name")
    {
        $ExchangeContact.NickName = $ContactItem."First Name" + " " + $ContactItem."Last Name";
    }
    elseif ($ContactItem."First Name" -and !$ContactItem."Last Name")
    {
        $ExchangeContact.NickName = $ContactItem."First Name";
    }
    elseif (!$ContactItem."First Name" -and $ContactItem."Last Name")
    {
        $ExchangeContact.NickName = $ContactItem."Last Name";
    }
    $ExchangeContact.DisplayName = $ExchangeContact.NickName;
    $ExchangeContact.FileAs = $ExchangeContact.NickName;
    
    $BusinessPhysicalAddressEntry = New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry;
    $HomePhysicalAddressEntry = New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry;
    $OtherPhysicalAddressEntry = New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry;
    
    # This uses the Contact Mapping above to save coding each and every field, one by one. Instead we look for a mapping and perform an action on
    # what maps across. As some methods need more "code" a fake multi-dimensional array (seperated by :'s) is used where needed.
    foreach ($Key in $ContactMapping.Keys)
    {
        # Only do something if the key exists
        if ($ContactItem.$Key)
        {
            # Will this call a more complicated mapping?
            if ($ContactMapping[$Key] -like "*:*")
            {
                # Make an array using the : to split items.
                $MappingArray = $ContactMapping[$Key].Split(":")
                # Do action
                switch ($MappingArray[0])
                {
                    "Email"
                    {
                        $ExchangeContact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::($MappingArray[1])] = $ContactItem.$Key;
                    }
                    "Phone"
                    {
                        $ExchangeContact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::($MappingArray[1])] = $ContactItem.$Key;
                    }
                    "Address"
                    {
                        switch ($MappingArray[1])
                        {
                            "Business"
                            {
                                $BusinessPhysicalAddressEntry.($MappingArray[2]) = $ContactItem.$Key;
                                $ExchangeContact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])] = $BusinessPhysicalAddressEntry;
                            }
                            "Home"
                            {
                                $HomePhysicalAddressEntry.($MappingArray[2]) = $ContactItem.$Key;
                                $ExchangeContact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])] = $HomePhysicalAddressEntry;
                            }
                            "Other"
                            {
                                $OtherPhysicalAddressEntry.($MappingArray[2]) = $ContactItem.$Key;
                                $ExchangeContact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])] = $OtherPhysicalAddressEntry;
                            }
                        }
                    }
                    "Method"
                    {
                        switch ($MappingArray[1])
                        {
                            "SetContactPicture" 
                            {
                                if (!$Exchange2007)
                                {
                                    if (!(Get-Item -Path $ContactItem.$Key -ErrorAction SilentlyContinue))
                                    {
                                        throw "Contact Picture File not found at $($ContactItem.$Key)";
                                    }
                                    $ExchangeContact.SetContactPicture($ContactItem.$Key);
                                }
                            }
                        }
                    }
                
                }                
            } else {
                # It's a direct mapping - simple!
                if ($ContactMapping[$Key] -eq "Birthday" -or $ContactMapping[$Key] -eq "WeddingAnniversary")
                {
                    [System.DateTime]$ContactItem.$Key = Get-Date($ContactItem.$Key);
                }
                $ExchangeContact.($ContactMapping[$Key]) = $ContactItem.$Key;            
            }
            
        }    
    }
    # Save the contact    
    $ExchangeContact.Save();
    
    # Provide output that can be used on the pipeline
    $Output_Object = New-Object Object;
    $Output_Object | Add-Member NoteProperty FileAs $ExchangeContact.FileAs;
    $Output_Object | Add-Member NoteProperty GivenName $ExchangeContact.GivenName;
    $Output_Object | Add-Member NoteProperty Surname $ExchangeContact.Surname;
    $Output_Object | Add-Member NoteProperty EmailAddress1 $ExchangeContact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1]
    $Output_Object;
}
}


function creatNewUser{

$strMailboxDatabase = "Mailbox Database 1188336507"
$strOU = "OU=NewUsers,DC=myDomain,DC=COM"
$Domain = "myDomain.com"
$strMX = "@" + $Domain
$outFile = "c:\admin\created.txt"
$Registrar = "pool.myDomain.com"

$defaultPW='BlackBerry123'
#$Password	= Read-Host 'Password' -AsSecureString
$Password = ConvertTo-SecureString $defaultPW -AsPlainText -Force

$objUsersNewOU = New-Object DirectoryServices.DirectoryEntry("LDAP://" + $strOU)
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objUsersNewOU

$WorkLog.text += "`r`n Checking Permission..."
checkADM

$WorkLog.text += "`r`n creating "+$InputAlias.Text+"@myDomain.com..."

add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
$colEmail = Import-csv c:\admin\NewUsers.txt
$RequestorText = $Requestor.test


	New-Mailbox -Alias $InputAlias.Text `
                  -UserPrincipalName ($InputAlias.Text + $strMX) `
                  -FirstName $InputFirstName.Text `
                  -LastName $InputLastName.Text `
                  -Name ($InputFirstName.Text + " " +$InputLastName.Text) `
                  -DisplayName ($InputFirstName.Text + " " +$InputLastName.Text) `
                  -Database $strMailboxDatabase `
                  -OrganizationalUnit $strOU `
                  -Password $Password `
                  -ResetPasswordOnNextLogon $False

    $WorkLog.text += "`r`n Enabling " + $InputFirstName.Text + " " + $InputLastName.Text + " for Outlook"

    Set-ADUser -Identity $InputAlias.Text -Description $RequestorTest

    $WorkLog.text += "`r`n Enabling " + $InputFirstName.Text + " " + $InputLastName.Text + " for Lync"

    Enable-csuser -identity ($InputAlias.Text) -registrarpool $Registrar -sipaddresstype EmailAddress

    $BesPassword	= 'password'

    $FullAdd = $InputAlias.Text + $strMX

    $WorkLog.text +=  "`r`n Importing fake contacts in " + $FullAdd

    Import-MailboxContacts -CSVFileName .\60fakeuserscontactexport.CSV -EmailAddress $FullAdd -Username besadmin -Password $BesPassword -Domain myDomain -Impersonate $true 

    $WorkLog.text += "`r`n"+ $FullAdd +" has been created!!!"
}

function groupAdded {

 $Group=$GroupList.text
 $FullAdd = $InputAlias.Text + $strMX

 $WorkLog.text += "`r`n Add " +$InputAlias.Text + "@myDomain.com to '$Group' group."

 Add-ADGroupMember -Identity $group -Members  $FullAdd

}


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()




#region begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '720,460'
$Form.text                       = "myDomain tools"
$Form.TopMost                    = $false

######
# Create User
#####

$FirstTitle                  = New-Object system.Windows.Forms.Label
$FirstTitle.text             = "===================Create a New User==================="
$FirstTitle.AutoSize         = $true
$FirstTitle.width            = 25
$FirstTitle.height           = 15
$FirstTitle.location         = New-Object System.Drawing.Point(35,30)
$FirstTitle.Font             = 'Microsoft Sans Serif,10'


$FirstNameTitle                  = New-Object system.Windows.Forms.Label
$FirstNameTitle.text             = "First Name:"
$FirstNameTitle.AutoSize         = $true
$FirstNameTitle.width            = 25
$FirstNameTitle.height           = 15
$FirstNameTitle.location         = New-Object System.Drawing.Point(35,65)
$FirstNameTitle.Font             = 'Microsoft Sans Serif,10'

$InputFirstName                       = New-Object system.Windows.Forms.TextBox
$InputFirstName.multiline             = $false
$InputFirstName.width                 = 80
$InputFirstName.height                = 25
$InputFirstName.location              = New-Object System.Drawing.Point(120,65)
$InputFirstName.Font                  = 'Microsoft Sans Serif,10'

$LastNameTitle                  = New-Object system.Windows.Forms.Label
$LastNameTitle.text             = "Last Name:"
$LastNameTitle.AutoSize         = $true
$LastNameTitle.width            = 25
$LastNameTitle.height           = 15
$LastNameTitle.location         = New-Object System.Drawing.Point(215,65)
$LastNameTitle.Font             = 'Microsoft Sans Serif,10'

$InputLastName                       = New-Object system.Windows.Forms.TextBox
$InputLastName.multiline             = $false
$InputLastName.width                 = 80
$InputLastName.height                = 25
$InputLastName.location              = New-Object System.Drawing.Point(300,65)
$InputLastName.Font                  = 'Microsoft Sans Serif,10'

$AliasTitle                  = New-Object system.Windows.Forms.Label
$AliasTitle.text             = "Alias:"
$AliasTitle.AutoSize         = $true
$AliasTitle.width            = 25
$AliasTitle.height           = 15
$AliasTitle.location         = New-Object System.Drawing.Point(35,100)
$AliasTitle.Font             = 'Microsoft Sans Serif,10'

$InputAlias                       = New-Object system.Windows.Forms.TextBox
$InputAlias.multiline             = $false
$InputAlias.width                 = 120
$InputAlias.height                = 25
$InputAlias.location              = New-Object System.Drawing.Point(85,100)
$InputAlias.Font                  = 'Microsoft Sans Serif,10'

$domainName                  = New-Object system.Windows.Forms.Label
$domainName.text             = "@myDomain.com"
$domainName.AutoSize         = $true
$domainName.width            = 120
$domainName.height           = 15
$domainName.location         = New-Object System.Drawing.Point(215,100)
$domainName.Font             = 'Microsoft Sans Serif,10'


$RequestorLabel                  = New-Object system.Windows.Forms.Label
$RequestorLabel.text             = "Requestor(description):"
$RequestorLabel.AutoSize         = $true
$RequestorLabel.width            = 25
$RequestorLabel.height           = 15
$RequestorLabel.location         = New-Object System.Drawing.Point(35,135)
$RequestorLabel.Font             = 'Microsoft Sans Serif,10'

$Requestor                       = New-Object system.Windows.Forms.TextBox
$Requestor.multiline             = $false
$Requestor.width                 = 250
$Requestor.height                = 25
$Requestor.location              = New-Object System.Drawing.Point(185,135)
$Requestor.Font                  = 'Microsoft Sans Serif,10'


$CreateUser                 = New-Object system.Windows.Forms.Button
$CreateUser.text             = "Create User"
$CreateUser.width            = 180
$CreateUser.height           = 30
$CreateUser.location         = New-Object System.Drawing.Point(100,170)
$CreateUser.Font             = 'Microsoft Sans Serif,10'

$GroupTitle                  = New-Object system.Windows.Forms.Label
$GroupTitle.text             = "Group:"
$GroupTitle.AutoSize         = $true
$GroupTitle.width            = 25
$GroupTitle.height           = 15
$GroupTitle.location         = New-Object System.Drawing.Point(35,220)
$GroupTitle.Font             = 'Microsoft Sans Serif,10'

$GroupList                       = New-Object  system.Windows.Forms.ComboBox
$GroupList.autosize            = $true
$GroupList.width                 = 170
$GroupList.height                = 25
# Add the items in the dropdown list
@('Global TSM','Sales Global','Customers') | ForEach-Object {[void] $GroupList.Items.Add($_)}
$GroupList.location              = New-Object System.Drawing.Point(85,220)
$GroupList.Font                  = 'Microsoft Sans Serif,10'

$AddGroup                 = New-Object system.Windows.Forms.Button
$AddGroup.text             = "Add Group"
$AddGroup.width            = 150
$AddGroup.height           = 25
$AddGroup.location         = New-Object System.Drawing.Point(285,220)
$AddGroup.Font             = 'Microsoft Sans Serif,10'

$GroupDescript                  = New-Object system.Windows.Forms.Label
$GroupDescript.text               = "Directly enter the group name if not in the list"
$GroupDescript.AutoSize         = $true
$GroupDescript.width            = 25
$GroupDescript.height           = 15
$GroupDescript.location         = New-Object System.Drawing.Point(35,245)
$GroupDescript.Font             = 'Microsoft Sans Serif,10'

#$GroupList.name | ForEach-Object {[void] $GroupList.Items.Add($_)}


######
# Reset Password
#####

$SecondTitle                  = New-Object system.Windows.Forms.Label
$SecondTitle.text             = "=================Reser User's Password================="
$SecondTitle.AutoSize         = $true
$SecondTitle.width            = 25
$SecondTitle.height           = 15
$SecondTitle.location         = New-Object System.Drawing.Point(35,275)
$SecondTitle.Font             = 'Microsoft Sans Serif,10'

$SecondDescription                  = New-Object system.Windows.Forms.Label
$SecondDescription.text             = "Either full name or alias"
$SecondDescription.AutoSize         = $true
$SecondDescription.width            = 25
$SecondDescription.height           = 15
$SecondDescription.location         = New-Object System.Drawing.Point(35,300)
$SecondDescription.Font             = 'Microsoft Sans Serif,10'

$FullNameTitle                  = New-Object system.Windows.Forms.Label
$FullNameTitle.text             = "First Name:"
$FullNameTitle.AutoSize         = $true
$FullNameTitle.width            = 25
$FullNameTitle.height           = 15
$FullNameTitle.location         = New-Object System.Drawing.Point(35,330)
$FullNameTitle.Font             = 'Microsoft Sans Serif,10'

$InputFullName                       = New-Object system.Windows.Forms.TextBox
$InputFullName.multiline             = $false
$InputFullName.width                 = 200
$InputFullName.height                = 25
$InputFullName.location              = New-Object System.Drawing.Point(120,330)
$InputFullName.Font                  = 'Microsoft Sans Serif,10'

$domainName2                  = New-Object system.Windows.Forms.Label
$domainName2.text             = "@myDomain.com"
$domainName2.AutoSize         = $true
$domainName2.width            = 120
$domainName2.height           = 15
$domainName2.location         = New-Object System.Drawing.Point(215,365)
$domainName2.Font             = 'Microsoft Sans Serif,10'

$AliasTitle2                  = New-Object system.Windows.Forms.Label
$AliasTitle2.text             = "Alias:"
$AliasTitle2.AutoSize         = $true
$AliasTitle2.width            = 25
$AliasTitle2.height           = 15
$AliasTitle2.location         = New-Object System.Drawing.Point(35,365)
$AliasTitle2.Font             = 'Microsoft Sans Serif,10'

$InputAlias2                       = New-Object system.Windows.Forms.TextBox
$InputAlias2.multiline             = $false
$InputAlias2.width                 = 120
$InputAlias2.height                = 25
$InputAlias2.location              = New-Object System.Drawing.Point(85,365)
$InputAlias2.Font                  = 'Microsoft Sans Serif,10'


$ResetPW                 = New-Object system.Windows.Forms.Button
$ResetPW.text             = "Reset Password"
$ResetPW.width            = 180
$ResetPW.height           = 30
$ResetPW.location         = New-Object System.Drawing.Point(100,410)
$ResetPW.Font             = 'Microsoft Sans Serif,10'



$ClearWorkLog                    = New-Object system.Windows.Forms.Button
$ClearWorkLog.text               = "Clear Worklog"
$ClearWorkLog.width              = 200
$ClearWorkLog.height             = 40
$ClearWorkLog.location           = New-Object System.Drawing.Point(510,325)
$ClearWorkLog.Font               = 'Microsoft Sans Serif,10'


$WorkLog                         = New-Object system.Windows.Forms.RichTextBox
$WorkLog.multiline               = $True
$WorkLog.width                   = 250
$WorkLog.height                  = 270
$WorkLog.location                = New-Object System.Drawing.Point(460,35)
$WorkLog.Font                    = 'Microsoft Sans Serif,10'
$WorkLog.Scrollbars              = "Vertical" 
$WorkLog.SelectionStart          = $WorkLog.Text.Length

$Close                           = New-Object system.Windows.Forms.Button
$Close.text                      = "Close Application"
$Close.width                     = 100
$Close.height                    = 50
$Close.location                  = New-Object System.Drawing.Point(610,390)
$Close.Font                      = 'Microsoft Sans Serif,10'

$Author                          = New-Object system.Windows.Forms.Label
$Author.text                     = "Created by Kim Chang"
$Author.AutoSize                 = $true
$Author.width                    = 25
$Author.height                   = 10
$Author.location                 = New-Object System.Drawing.Point(10,10)
$Author.Font                     = 'Microsoft Sans Serif,8'



$Form.controls.AddRange(@($FirstTitle,$FirstNameTitle,$InputFirstName,$LastNameTitle,$InputLastName,$AliasTitle,$InputAlias,$domainName,$CreateUser,$RequestorLabel,$Requestor,$GroupTitle,$AddGroup,$GroupList,$GroupDescript,
$SecondTitle,$SecondDescription,$FullNameTitle,$InputFullName,$AliasTitle2,$InputAlias2,$domainName2,$ResetPW,
$Close,$WorkLog,$ClearWorkLog,$Author))

#region gui events {
$CreateUser.Add_Click({ creatNewUsertxt }) 
$close.Add_Click({ closeForm })
$ClearWorkLog.Add_Click({ ClearWorkLog })
$ResetPW.Add_Click({ changePW })
$AddGroup.Add_Click({ groupAdded })
#endregion events }


#endregion GUI }

function closeForm(){$Form.close()}

function ClearWorkLog(){$WorkLog.text = ''}



#Show the form
[void]$Form.ShowDialog() 

