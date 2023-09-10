#----------------------------------------------------------------------
#Author    : MJBCT / Marcin Jędorowicz
#Data      : 2018-09-05
#Version   : 1.0.0
#CopyRight : MJBCT / Marcin Jędorowicz
#----------------------------------------------------------------------
<#
.SYNOPSIS
Script to check Active Sync using a group from AD.
    .
.DESCRIPTION
The script is used to leave the Active Sync functionality enabled only for people who are members of a specific group in AD.

.PARAMETER
-ADGroup - specifying a group from AD for Activesync verification.

-ExchangeFQDN - specifying the name of the Exchange server using FQDN

-logFolder - location of saving the login file by default C:\Windows\System32\LogFiles
-senderEmail - EmaillAddress for sendig information
-recipientEmail - EmaillAddress to recipient sendig information

.EXAMPLE
    ./EnabledActiveSyncByADGroup.ps1 Such a call will trigger information that there is no data regarding the AD group and/or the Exchange server name

    ./EnabledActiveSyncByADGroup.ps1 -ADGroup 'Exchange ActiveSync Allowed' -ExchangeFQDN Exchange.onmicrosoft.com -logFolder "C:\" 
    ./EnabledActiveSyncByADGroup.ps1 -ADGroup 'Exchange ActiveSync' -ExchangeFQDN Exchange.onmicrosoft.com -senderEmail PS_JOB@<domain> recipientEmail Recipient@<domain>

.INPUTS

.OUTPUTS
The result file is a log of the performed operations. The log is saved by default in C:\Windows\System32\LogFiles\
    C:\Windows\System32\LogFiles\PS_ActiveSync"+ $date +".log"

.NOTES
    .
.LINK
    .
#>

Param(
[Parameter (Mandatory=$True)] [string] $ADGroup, #group name in AD, Parameter (Mandatory=True) - forcing the parameter to be provided
[string] $exchangeFQDN, #FQDN address of the Exchange server - forcing the parameter to be entered
[string] $logFolder ="C:\Windows\System32\LogFiles", #saving the login file
[string] $senderEmail, #EmaillAddress for sendig information
[string] $recipientEmail #EmaillAddress to recipient information
)

$date = get-date -UFormat "%Y-%m-%d"
$logfile = $logFolder +"\PS_ActiveSync"+ $date +".log"

#Clear Error Cash
$Error.Clear()


#write starting Script
(get-date).Tostring() + ‘ <=====----- START -----=====>‘| Out-file $logfile -append 

(get-date).Tostring() + ‘ Using parameters :'| Out-file $logfile -append 
(get-date).Tostring() + ‘ ADGroup ' + $ADGroup| Out-file $logfile -append 
(get-date).Tostring() + ‘ ExchangeFQDN ' + $ExchangeFQDN| Out-file $logfile -append 
(get-date).Tostring() + ‘ logfolder ' + $logFolder| Out-file $logfile -append 
(get-date).Tostring() + ‘ senderEmail ' + $senderEmail| Out-file $logfile -append 
(get-date).Tostring() + ‘ recipientEmail ' + $senderEmail| Out-file $logfile -append 

#Transaction for connection to Exchange server
try{
$uri="http://"+$ExchangeFQDN+"/PowerShell/"
$ExSession= New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos 
Import-PSSession $ExSession -AllowClobber
}
Catch{
    $str = "issue establishing a session with the Exchange server" + $ExchangeFQDN
    (get-date).Tostring() + ‘ ‘ + [string] $str | Out-file $logfile -append 
        
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    (get-date).Tostring() + ‘ ‘ + $errorMessage | Out-file $logfile -append 
    (get-date).Tostring() + ‘ ‘ + $FailedItem | Out-file $logfile -append 
}
Finally{
if($ExSession){
    #Clear Error Cash
    $Error.Clear()
    $allUsers = get-Mailbox -ResultSize:unlimited
    $groupUsers = Get-ADGroupMember -Identity $ADGroup
    $Changes=0

    # Action for every user 
    foreach ($member in $allUsers)  
    { 
        $str = "" 
     
        #Get settings for specific user
        $mailbox = Get-CasMailbox -resultsize unlimited -identity $member.Name 
     
        #verification whether the user is a member of the group 
        if(($groupUsers | where-object{$_.Name -eq $member.Name})) 
        { 
            #If ActiveSync is Enabled - nothing to do 
            if ($mailbox.ActiveSyncEnabled -eq "true") 
            { 
            } 
            #If ActiveSync is Disabled - Enabled it 
            else 
            { 
                $member | Set-CASMailbox –ActiveSyncEnabled:$true -confirm:$false
                $str = "Enabled ActiveSync - " + $mailbox.Name + "`n" 
                (get-date).Tostring() + ‘ ‘ + [string] $str | Out-file $logfile -append 
                $Changes+=1
            } 
            #If ExchangeWebServis is Enabled - nothing to do
            if ($mailbox.EwsEnabled -eq "true") 
            { 
            } 
            #If ExchangeWebServis is Disabled - Enabled it 
            else 
            { 
                $member | Set-CASMailbox –EwsEnabled:$true -confirm:$false
                $str = "Enabled EWS - " + $mailbox.Name + "`n" 
                (get-date).Tostring() + ‘ ‘ + [string] $str | Out-file $logfile -append 
                $Changes+=1
            } 
        } 
        #If the user is not a member of the group, disable ActiveSync
        else 
        { 
            if ($mailbox.ActiveSyncEnabled -eq "true") 
            { 
                $member | Set-CASMailbox –ActiveSyncEnabled $false -confirm:$false
                $str = "Disabled ActiveSync - " + $mailbox.Name + "`n" 
                (get-date).Tostring() + ‘ ‘ + [string] $str | Out-file $logfile -append 
                $Changes+=1
            } 
         } 
    }

    Remove-PSSession $ExSession
}
}

#checking if an error occurred
if ($Error -ne 0){
    #Inform of ending script and error occured
    Write-Host "Finishes script processing. problems occurred. details in the file " + $logfile -ForegroundColor red
    $Body = "A review of enabled ActiveSync features was performed. Problems occurred during operation, information is in the file "+$LogFile
    Send-MailMessage -SmtpServer $ExchangeFQDN -From $senderEmail -To $recipientEmail -Subject "[Error] Review of ActiveSync" -Body $Body -Encoding UTF8
}
else{
    #Inform of ending script
    Write-Host "Finishes script processing. problems occurred. details in the file " + $logfile -ForegroundColor green
    $Body = "A review of enabled ActiveSync features was performed. For " + $Changes + " accounts have been changed. Details are in the file "+$LogFile
    Send-MailMessage -SmtpServer $ExchangeFQDN -From $senderEmail -To $recipientEmail -Subject "Review of ActiveSync" -Body $Body -Encoding UTF8
}

#write end script
(get-date).Tostring() + ‘ <=====----- STOP -----=====>‘| Out-file $logfile -append 
