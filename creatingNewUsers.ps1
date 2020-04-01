#######################################################################################
###
### Created by Kim Chang 
### Last update @2020/03/27
### This is for creating a new account in myDomain
### Steps for this script==============================
### 1. Creating a new user list file
### 2. Checking if Alias is already exist
### 3. Creating AD new account one by one
### 4. Creating new account mailbox one by one
### 5. Enable Lync account one by one
### 6. Import fake contacts to each new users (This usually takes longer time)
### 7. Display user creation result 
### 8. Ask which group should add the user to
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
    while ($dup)  {
    $dupName=$dup.UserPrincipalName
    Write-Host "OOps your alias has been used`n Duplicate username: '$dupName'"
    $Alias=Read-Host -Prompt 'Please enter another Alias'
    $dup=Get-ADUser -Filter "UserPrincipalName -like '$Alias*'"
            } 
    return $Alias

}

function creatNewUsertxt{

$Alias=@()
$FirstName=@()
$LastName=@()
$count='y'

"Alias,FirstName,LastName" >  NewUsers.txt


$i=1

    do {
		#write-host "`n"
        Write-Host "`n=======Please enter the user information: User$i =========" 
        #Write-Host "=======If you want to quit the script, press 'q' anytime.  =========" 
        $Alias1=Read-Host -Prompt 'Please enter the Alias(ex:kchang)'
            #if (($Alias -eq 'q')){exit}
        $Alias =  duplicateNameCheck $Alias1
        $FirstName=Read-Host -Prompt 'Please enter the First Name(ex:kim)'
            #if (($FirstName -eq 'q')){exit}
        $LastName=Read-Host -Prompt 'Please enter the Last Name(ex:chang)'
            #if (($LastName -eq 'q')){exit}

        if (([string]::IsNullOrWhitespace($Alias)) -or ([string]::IsNullOrWhitespace($FirstName)) -or ([string]::IsNullOrWhitespace($LastName))){
        Read-Host "Oops!Something is wrong! Please enter information again..."
        continue
        }
        "$Alias,$FirstName,$LastName" >>  NewUsers.txt
        $count=Read-Host -Prompt "Do you want to input another user? (y/n)"
        
        $i=$i+1

    }until (($count -eq 'n') -or ($count -eq 'q'))
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

function addGroup {

    Param
    (
        [Parameter(Mandatory=$true,
                    ValueFromPipelineByPropertyName=$true,
                    Position=0)]
        $group
    )

$Content=Get-Content NewUsers.txt | select-object -skip 1

$Content | Foreach {
    $col=$_ -split ','
    Write-Host $col[0]
    Add-ADGroupMember -Identity $group -Members $col[0]
}
}

function groupMenu
{
      Write-Host "=======Select group ========="      
      Write-Host "Press '1' for Global TSM"
      Write-Host "Press '2' for Sales Global"
      Write-Host "Press '3' for Customers"
      Write-Host "Press '4' for Others"
      
 }


###Parameter

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


###Main
Write-Host "Checking for elevated permissions..."
checkADM

add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010

Write-Host "=================Creating new user(s)============================="
Write-Host "Do you want to input user manually? If you have NewUser.txt file," 
$input=Read-Host -Prompt "please enter 'n' (y/n)"
if ($input -eq 'y'){creatNewUsertxt}

$colEmail = Import-csv c:\admin\NewUsers.txt
$Requestor = Read-Host 'Requestor (request information)'

if (Test-Path $outFile){remove-item $outFile}


ForEach ($row in $colEmail) {
	($row.alias + $strMX) | Out-File -filepath $outFile -append
      
	New-Mailbox -Alias $row.Alias `
                  -UserPrincipalName ($row.alias + $strMX) `
                  -FirstName $row.FirstName `
                  -LastName $row.LastName `
                  -Name ($row.FirstName + " " +$row.LastName) `
                  -DisplayName ($row.FirstName + " " +$row.LastName) `
                  -Database $strMailboxDatabase `
                  -OrganizationalUnit $strOU `
                  -Password $Password `
                  -ResetPasswordOnNextLogon $False
	
}  #End Email Loop


# create new user
ForEach ($row in $colEmail) {
$row.descrition
Set-ADUser -Identity $row.alias -Description $Requestor
}

# enable lync (Skype of Business)
ForEach ($row in $colEmail) {
# Enable for lync and configure settings
"Enabling " + $row.FirstName + " " + $row.LastName + " for Lync"

Enable-csuser -identity ($row.Alias) -registrarpool $Registrar -sipaddresstype EmailAddress
}


#import fake contacts
$NewUsers = Import-CSV C:\admin\Created.txt -Header @("Email")
$BesPassword	= Read-Host 'BESADMIN Password'

foreach ($row in $NewUsers) {
Import-MailboxContacts -CSVFileName .\60fakeuserscontactexport.CSV -EmailAddress $row.Email -Username besadmin -Password $BesPassword -Domain myDomain -Impersonate $true 
}


#output results
write-host "=======Created User Summary======="

ForEach ($row in $colEmail) {

    $Name=($row.FirstName + " " +$row.LastName) 
	$UName=$row.Alias 
    $Email=($row.Alias +"@myDomain.com")
    
    Write-Host "$Name account has been created."
    Write-Host "Username:$UName"
    Write-Host "Email:$Email"
    Write-Host "Password: BlackBerry123 `n"
} 




do {
    $input2=Read-Host -Prompt "Do you want to add user(s) to a group? (y/n)"
    if ($input2 -eq 'n') {break}
    groupMenu
    $defaultGroup=Read-Host -Prompt "Please enter your selection"

    switch($defaultGroup){
     '1' {
            addGroup "Global TSM"
    }'2' {
            addGroup "Sales Global"
    }'3' {
            addGroup "Customers"
    }'4' {
            $groupName=Read-Host -Prompt "Please type the group name"
            addGroup $groupName
    }
    }
}until ($input2 -eq 'n')





