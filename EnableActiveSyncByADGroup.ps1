# Assign ALL USERS to a dynamic array 
$allUsers = get-Mailbox -ResultSize:unlimited 
 
# Assign all members of the ALLOWED GROUP to a dynamic array 
$groupUsers = Get-ADGroupMember -Identity '00SU-Exchange ActiveSync Allowed' 
 
# Loop through array of all users 
foreach ($member in $allUsers)  
{ 
    $str = "" 
     
    #get CAS attributes for current user 
    $mailbox = Get-CasMailbox -resultsize unlimited -identity $member.Name 
     
    #determine if current user is member of allowed group 
    if(($groupUsers | where-object{$_.Name -eq $member.Name})) 
    { 
        #if user already has ActiveSync enabled, do nothing 
        if ($mailbox.ActiveSyncEnabled -eq "true") 
        { 
        } 
        #if user does not have ActiveSync enabled, enable it 
        else 
        { 
            $member | Set-CASMailbox –ActiveSyncEnabled $true -confirm:$false
            $str += "Włączono ActiveSync - " + $mailbox.Name + "`n" 
	    (get-date).Tostring() + ‘ ‘ + [string] $str | Out-file C:\BIN\ActiveSync_LOG.txt -append         } 
    } 
    #if user is not member of allowed group, disable ActiveSync 
    else 
    { 
        if ($mailbox.ActiveSyncEnabled -eq "true") 
        { 
            $member | Set-CASMailbox –ActiveSyncEnabled $false -confirm:$false
            $str = "Wyłączono ActiveSync - " + $mailbox.Name + "`n" 
            (get-date).Tostring() + ‘ ‘ + [string] $str | Out-file C:\BIN\ActiveSync_LOG.txt -append 
        } 
    } 
}
