Connect-ExchangeOnline

#Check if the user is in the group
function Check-Membership($group, $user) {
    $members = Get-DistributionGroupMember -Identity $group
    $isMember = $members | Where-Object { $_.PrimarySmtpAddress -eq $user }
    if ($isMember)
    {
        return $true
    }
    else 
    {
        return $false
    }
}

function Check-UserExists($user) {
    $recipient = Get-EXORecipient -Filter "EmailAddresses -eq 'SMTP:$user'"
    if ($recipient) 
    {
        Write-Host "$user exists in Exchange."
        return $true
    } 
    else 
    {
        Write-Host "The email address $user does not exist in Exchange."
        return $false
    }
}

#Adds user to AD if they exist and they're not in the group already
function Add-DLUser($group, $user){
    if (Check-UserExists $user -and Check-Membership $group $user)
    {
        Add-DistributionGroupMember -Identity $group -Member $user
    } 
}

# Check if the ExchangeOnlineManagement module is installed
$moduleName = "ExchangeOnlineManagement"
$module = Get-Module -ListAvailable -Name $moduleName

if ($module) {
    Write-Host "The $moduleName module is installed."
} else {
    Install-Module -Name ExchangeOnlineManagement -Force
}