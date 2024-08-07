function Check-Membership($group, $user) {
    $members = Get-DistributionGroupMember -Identity $group
    $isMember = $members | Where-Object { $_.PrimarySmtpAddress -eq $user }
    return $isMember
}

function Add-DLUser($group, $user){
    if (Check-Membership $group $user)
    {
        Write-Host "User already exists in this group"
    } 
    else 
    {
        Add-DistributionGroupMember -Identity $group -Member $user
    }
}