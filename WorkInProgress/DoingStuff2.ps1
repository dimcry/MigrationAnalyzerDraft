cd "C:\Users\cristid\AppData\Local\Temp\MigrationAnalyzer\03042019_173206\SavedData"
dir

$EXO_Recipient_Dtest2 = Import-Clixml .\EXO_Recipient_Dtest2@dimcry.ro.xml
$EXO_Recipient_Dtest5 = Import-Clixml .\EXO_Recipient_Dtest5@dimcry.ro.xml
$EXO_Recipient_Dtest6 = Import-Clixml .\EXO_Recipient_Dtest6@dimcry.ro.xml

$EXO_Recipient_Dtest2 | ft -AutoSize *Move*
$EXO_Recipient_Dtest5 | ft -AutoSize *Move*
$EXO_Recipient_Dtest6 | ft -AutoSize *Move*

cd "C:\Temp\MSSupport"
dir
$Recipient_Hybrid11 = Import-Clixml .\Recipient_Hybrid11@dimcry.ro.xml
$Recipient_Hybrid11 | ft -AutoSize *Move*
$Recipient_Hybrid4 = Import-Clixml .\Recipient_Hybrid4@dimcry.ro.xml
$Recipient_Hybrid4 | ft -AutoSize *Move*
$Recipient_T1 = Import-Clixml .\Recipient_T1@dimcry.ro.xml
$Recipient_T1 | ft -AutoSize *Move*



if (($($Recipient_Hybrid11.MailboxMoveStatus) -ne "None") -and ($($Recipient_Hybrid11.MailboxMoveFlags) -ne "None") -and ($($Recipient_Hybrid11.MailboxMoveTargetMDB) -or ($($Recipient_Hybrid11.MailboxMoveSourceMDB)))) {
    Write-host "Move in progress" -ForegroundColor Green
}
else {
    Write-host "No move in progress" -ForegroundColor Red
}

if (($($EXO_Recipient_Dtest2.MailboxMoveStatus) -ne "None") -and ($($EXO_Recipient_Dtest2.MailboxMoveFlags) -ne "None") -and ($($EXO_Recipient_Dtest2.MailboxMoveTargetMDB) -or ($($EXO_Recipient_Dtest2.MailboxMoveSourceMDB)))) {
    Write-host "Move in progress" -ForegroundColor Green
}
else {
    Write-host "No move in progress" -ForegroundColor Red
}


if ($($EXO_Recipient_Dtest2.MailboxMoveTargetMDB) -or ($($EXO_Recipient_Dtest2.MailboxMoveSourceMDB))) {
    Write-host "Move in progress" -ForegroundColor Green
}
else {
    Write-host "No move in progress" -ForegroundColor Red
}