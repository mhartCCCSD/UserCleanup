param(
    [Parameter(Mandatory = $True)]
    [string]$username
)
$VerbosePreference = "Continue"
######################################################################################################
### Import all necessary modules
### Declare all variables
### Prompt for both sets of credentials
### Create session variables for on-prem and cloud sessions
Import-Module Microsoft.Online.Sharepoint.PowerShell
Import-Module MSOnline
Import-Module ActiveDirectory
$emailbody = ""
$count = 0 ### Used to count the number of groups a user was removed from
$upnusername = $username.split("@")[0]
$cccsdupnusername = "CCCSD\" + $upnusername
$targetOU = ""
$aduserDN = ""
$domainstring = "DC=cccsd,DC=centralsan,DC=dst,DC=ca,DC=us"
$usermanager = $null
$localadmin = Get-Credential -message "Please enter your local Exchange administrator credentials (not Office365 credentials)"
$cloudadmin = Get-Credential -message "Please enter your Office365 administrator credentials"
$localsession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mtzcas01/powershell/ -Authentication Kerberos -Credential $localadmin
$cloudsession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell/" -Credential $cloudadmin -Authentication Basic -AllowRedirection
$securitysession = new-pssession -configurationname Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -credential $cloudadmin -authentication Basic -AllowRedirection
######################################################################################################
### Connect to security and compliance center to start the export to .pst
Import-PSSession $securitysession -AllowClobber
### Create a new search based on the username
$compliancename = "export "+$username + " to pst"
New-Compliancesearch -name $compliancename -exchangelocation $username
Start-ComplianceSearch -identity $compliancename
### Have the script pause for 300 seconds to allow the search time to finish
Clear-Host
for ($a=120; $a -gt 1; $a--){
    Write-Progress -Activity "Searching Office365 mailbox..."`
    -SecondsRemaining $a `
    -Status "Please wait for the search to be completed"
    Start-Sleep 1
}
### Now start the export so all we have to do is go and save it
New-ComplianceSearchAction -searchname $compliancename -Export
$emailbody = "<html><body>Beginning user removal for " + $username + "<br>"
$emailbody = $emailbody + "New search started. <a href=https://protection.office.com/#/contentsearch>Login here</a> to start the download for the search called: " + $compliancename + "<br>"
### Disconnect from security and compliance
remove-pssession $securitysession
### Connect to the cloud instance to query group membership
Import-PSSession $cloudsession -AllowClobber
Connect-MsolService -Credential $cloudadmin
### Return mailbox details for the user
$mailbox=get-mailbox $username
### Now get all groups to step through and search for the user
$groups= Get-DistributionGroup
### Remove the license from the user
Set-MsolUserLicense -userprincipalname $username -removelicenses "centralsan:ENTERPRISEPACK_GOV"
### Disconnect from cloud instance
remove-PSsession $cloudsession
### Connect to local instance to do the actual removal (necessary for hybrid scenario)
Import-PSSession $localsession -AllowClobber
foreach($dg in $groups){
    $DGMs = Get-DistributionGroupMember -identity $dg.Identity
    foreach ($dgm in $DGMs){
        if ($dgm.name -eq $mailbox.name){
            $count = $count + 1
            $emailbody = $emailbody + "User removed from group: " + $dg.identity + "<br>"
              Remove-DistributionGroupMember $dg.Name -Member $username -confirm:$false
        }
    }
}
$emailbody = $emailbody + "User was found and removed from " + [string]$count + " group(s)<br>"
################################################################################################################
################################### Need to add:
################################### Change PC description to OU/IP-VLAN/Model#

### Disconnect from the local session
Remove-PSSession $localsession
### Begin AD cleanup (disable user for now)
### Connect to AD
Enter-PSSession -ComputerName mtzpdc01 -credential $localadmin
### Get the user's manager (to email at a later date)
$aduser = get-aduser -identity $upnusername -properties OfficePhone, manager
$usermanager = (get-aduser (get-aduser $upnusername -properties manager).manager).samaccountname
if (!$usermanager){
    $emailbody = $emailbody + $upnusername + " did not have a manager listed in AD<br>"
}
Else {
    $emailbody = $emailbody + $upnusername + "'s manager is " + $usermanager + "<br>"
}

$phoneemailbody = "Hello, the user $upnusername has been marked for termination."
if ($aduser.officephone -eq $null){
    $phoneemailbody = $phoneemailbody + " Their account did not have a phone number associated with it in Active Directory."
    if (!$usermanager){
        $phoneemailbody = $phoneemailbody + "There is no manager listed in AD. Happy hunting!"
    }
    else {
        $phoneemailbody = $phoneemailbody + "You can try contacting their manager, $usermanager, to see if they can assist."
    }
    $emailbody = $emailbody + "Their account did not have a phone number associated with it<br>"
}
else {
    $phoneemailbody = $phoneemailbody + "Their account was associated with the phone number " + $aduser.officephone
    if (!$usermanager){
        $phoneemailbody = $phoneemailbody + + ". The user has no manager listed in AD. Happy hunting!"
    }
    else {
        $phoneemailbody = $phoneemailbody  + ". If you have any questions, their manager is/was $usermanager"
    }
    $emailbody = $emailbody + "Their account was associated with the phone number " + $aduser.officephone + "<br>"
}
### Move the user to the disabled OU
$position = $aduser.distinguishedname.IndexOf(",")
$aduserDN = $aduser.distinguishedname.SubString($position+1)
$aduserDNminusdomain = $aduserDN.Replace($domainstring,"")
$targetOU = $aduserDNminusdomain+"OU=Disabled Items,"+$domainstring
### Test if the target OU exists
if ([adsi]::Exists("LDAP://$targetOU")){
    $emailbody = $emailbody + "Moved "+$username+" from:<br>"+$aduserDN+"<br>to:<br>"+$targetOU+"<br>"
    Move-ADObject $aduser -TargetPath $targetOU
}
else {
    $emailbody = $emailbody + "Attempting to move user to: "+$targetOU+"<br>FAILED!<br>Please create the appropriate container and then manually move the user<br>"
}
### Disable the account
Disable-ADAccount $upnusername
$emailbody = $emailbody + "Disabled " + $upnusername + " in Active Directory<br>Office365 License removed<br>Cleanup completed<br>To finalize you must download the exported search results from the mailbox export</body></html>"
### Setup the splat to send email
$MailMessage = @{
    To="mhart@centralsan.org"
    From="MailUserMaint@centralsan.org"
    Subject="Maintenance on user " + $username
    Body=$emailbody
    SMTPServer="distprint"
}
### Setup the splat to mail the phone admin
$MailPhoneMessage = @{
    To="jvega@centralsan.org"
    From="PhoneUserMaint@centralsan.org"
    Subject="Phone Maintenance on user " + $username
    Body=$Phoneemailbody
    SMTPServer="distprint"
}
$mhartMailPhoneMessage = @{
    To="mhart@centralsan.org"
    From="PhoneUserMaint@centralsan.org"
    Subject="Phone Maintenance on user " + $username
    Body=$Phoneemailbody
    SMTPServer="distprint"
}
### Send the message
Send-MailMessage @MailMessage -BodyAsHtml
Send-MailMessage @MailPhoneMessage
Send-MailMessage @mhartMailPhoneMessage
### Cleanup
Get-PSSession | Remove-PSSession
Exit-PSSession
### End of script