#----------------------------------------------------------------------
#Author    : MJBCT IT / Marcin Jędorowicz
#Data      : 2016-09-23
#Version   : 1.0.1
#CopyRight : MJBCT IT / Marcin Jędorowicz
#----------------------------------------------------------------------
<#
.SYNOPSIS
Skrypt do wysyłania wiadomości mailowych
    .
.DESCRIPTION
Skrypt służy do wysyłania jednej wiadomości mailowej do masowej liczby użytkowników.

.PARAMETER
[string] $attachment,

-usersListCSV - PARAMETR WYMAGANY lista użytkowników do których nalezy wysłac wiadomość

-sender - PARAMETR WYMAGANY nadawca wiadomości

-subject - temat wiadomości domyślnie Departament Informatyki NFZ Centrala - ważna informacja

-casServer - nazwa serwera z którego będa wysyłane wiadomości , domyślnie 00sexc01.health.local

-messageBodyFile - PARAMETR WYMAGANY wskazanie pliku zawierającego wiadomość

-attachment - wskazanie pliku jako załącznika

.EXAMPLE

.INPUTS

.OUTPUTS
Plikiem wynikowym jest log z wykonywanych operacji. Log domyślnie jest zapisywany w C:\Windows\System32\LogFiles\
Pliki generowane są w:
    C:\Windows\System32\LogFiles\PS_SendMessage"+ $date +".log"

.NOTES
    .
.LINK
    .
#>

Param(
[Parameter (Mandatory=$True)] [string]$usersListCSV,
[Parameter (Mandatory=$True)] [string]$sender,
[string] $casServer = "FQDN ClientAccessServer",
[string] $subject = "Tytuł wiadomości - ważna informacja",
[Parameter (Mandatory=$True)] [string] $messageBodyFile,
[string] $attachment,
[string] $logFolder ="C:\Windows\System32\LogFiles" #folder zapisywania logu
)
#parametr zapisywania logowania
$date = get-date -UFormat "%Y-%m-%d"
$logfile = $logFolder +"\PS_SendMessage"+ $date +".log"

#Czyszczenie Cash dla błędów
$Error.Clear()

#Wpisanie w logu rozpoczecia działania skrypu
(get-date).Tostring() + ‘ <=====----- START -----=====>‘| Out-file $logfile -append 

(get-date).Tostring() + ‘ Użyte parametry :'| Out-file $logfile -append 
(get-date).Tostring() + ‘ usersListCSV ' + $usersListCSV| Out-file $logfile -append 
(get-date).Tostring() + ‘ sender ' + $sender| Out-file $logfile -append 
(get-date).Tostring() + ‘ casServer ' + $casServe| Out-file $casServer -append 
(get-date).Tostring() + ‘ subject ' + $subject| Out-file $logfile -append 
(get-date).Tostring() + ‘ messageBodyFile ' + $messageBodyFile| Out-file $logfile -append 
(get-date).Tostring() + ‘ attachment ' + $attachment| Out-file $logfile -append 
(get-date).Tostring() + ‘ logfolder ' + $logFolder| Out-file $logfile -append 

if (!(Test-Path $usersListCSV))
{
    Write-Host "Problem z przekazanymi informacjami odnośnie listy mailowej" -ForegroundColor red
    Break
}
elseif(($messageBodyFile -eq "") -and !(Test-Patch $messageBodyFile))
{
    Write-Host "Problem z przekazanymi informacjami odnosnie pliku zawierającego wiadomość" -ForegroundColor red
    Break
}

try{
    if ($messageBodyFile -ne ""){
        $message = (Get-Content $messageBodyFile) | out-String
    }

    (get-date).Tostring() + ‘ message ' + $lmessage| Out-file $logfile -append 

    $recipientCount = (Get-Content -Path $usersListCSV | Measure-Object -Line).Lines -1

    $counter = 1
    $users = $users = Import-CSV $usersListCSV
    foreach ($user in $users) {
        Write-Progress -Id 0 -Activity ("Wysyłam maila na adres: " + ($user.user | out-String)) -Status "$counter z $recipientCount" -PercentComplete ($counter / $recipientCount*100) 
        
        try{
            if ($attachment -ne ""){
                Send-MailMessage -To $user.user -From $sender -SMTPServer $casServer -Subject $subject -Attachments $attachment -Body $message -Encoding UTF8 -BodyAsHtml 
            }
            else{
                Send-MailMessage -To $user.user -From $sender -SMTPServer $casServer -Subject $subject -Body $message -Encoding UTF8 -BodyAsHtml 
            }
        
            $str = "Wysłano wiadomośc do użytkownika do " + $user.user + "`n"
            (get-date).Tostring() + ‘ ‘ + $str | Out-file $logfile -append 
        }
        catch{
            $str = "problem z wysłaniem wiadomości do " + $user.user + "`n"
            (get-date).Tostring() + ‘ ‘ + [string] $str | Out-file $logfile -append 
        
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            (get-date).Tostring() + ‘ ‘ + $errorMessage | Out-file $logfile -append 
            (get-date).Tostring() + ‘ ‘ + $FailedItem | Out-file $logfile -append         
        }

        Start-Sleep -Milliseconds 200 
        if ($counter%100 -eq 0) {Start-Sleep 10}
        $counter++
    } 

}
catch{
    $str = "problem z wysłaniem wiadomości " + "`n"
    (get-date).Tostring() + ‘ ‘ + [string] $str | Out-file $logfile -append 
        
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    (get-date).Tostring() + ‘ ‘ + $errorMessage | Out-file $logfile -append 
    (get-date).Tostring() + ‘ ‘ + $FailedItem | Out-file $logfile -append 
}

#sprawdzanie czy wystąpił błąd
if ($Error -ne 0){
    #poinformowanie o zakończeniu pobierania danych oraz że wystąpiły błedy
    Write-Host "Zakończenie przetwarzania skryptu. Wystąpił problem, szczegóły w pliku " + $logfile -ForegroundColor red
}
else{
    #poinformowanie o zakończeniu pobierania
    Write-Host "Zakończenie przetwarzania skryptu. Informacje szczegółowe " $logfile -ForegroundColor green
}

#wpisanie zakończenia działania w pliku LOG
(get-date).Tostring() + ‘ <=====----- STOP -----=====>‘| Out-file $logfile -append 
