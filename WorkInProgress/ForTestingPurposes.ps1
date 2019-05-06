function Collect-MoveRequestStatistics {
    [CmdletBinding()]
    param (
        [string]
        $AffectedUser,
        [int]
        $NumberOfChecks
    )

    try {
        Get-EXOMoveRequest $AffectedUser -ErrorAction Stop
    }
    catch {
        Write-Log "[WARNING] || We were unable to find a MoveRequest in place for `"$AffectedUser`". We will check if MigrationUser was created for it."
        Collect-MigrationUserStatistics -AffectedUser $AffectedUser -MigrationType Hybrid -ErrorAction SilentlyContinue
        Break
    }

    if ($NumberOfChecks -le 5) {
        try {
            $script:MoveRequestStatistics = Get-EXOMoveRequestStatistics $AffectedUser -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" -ErrorAction Stop
        }
        catch {
            Write-Log "[WARNING] || ($NumberOfChecks / 5) We were unable to collect MoveRequestStatistics"
            $NumberOfChecks++
            Collect-MoveRequestStatistics -AffectedUser $AffectedUser -NumberOfChecks $NumberOfChecks
        }
    }
    else {
        Write-Log "[WARNING] || We were unable to find a MoveRequestStatistics in place for `"$AffectedUser`". We will check if MigrationUser was created for it."
        Collect-MigrationUserStatistics -AffectedUser $AffectedUser -MigrationType Hybrid -ErrorAction SilentlyContinue
    }
}

function Collect-SyncRequestStatistics {
    [CmdletBinding()]
    param (
        [string]
        $AffectedUser
    )

    if (Get-EXOSyncRequest -Mailbox $AffectedUser -ErrorAction SilentlyContinue)  {
        $script:SyncRequestStatistics = Get-EXOSyncRequestStatistics $AffectedUser -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose"
    }
    else {
        Write-Log ("[WARNING] || We were unable to find a SyncRequest in place for `"$AffectedUser`". We will check if MigrationUser was created for it.")
        Collect-MigrationUserStatistics -AffectedUser $AffectedUser -MigrationType IMAP
    }
}

function Collect-MigrationUserStatistics {
    [CmdletBinding()]
    param (
        [string]
        $AffectedUser,
        [string]
        $MigrationType
    )

    try {
        Get-EXOMigrationUser $AffectedUser -ErrorAction Stop
    }
    catch {
        Write-Log "[WARNING] || We were unable to find a MigrationUser in place for `"$AffectedUser`". This might be because you haven't provided the PrimarySMTPAddress of the user, or, no MigrationUser is present."
        Write-Log "[INFO] || We are trying to identify the `"$AffectedUser`" recipient in your organization."
        try {
            $script:Recipient = Get-EXORecipient $AffectedUser -ErrorAction Stop
        }
        catch {
            Write-Log "[WARNING] || We were unable to find, in your organization, a recipient for `"$AffectedUser`"."
            Write-Log "[ERROR] || The script will exit now. Please re-run it, using the PrimarySMTPAddress of the affected user."
            Exit
        }
        $AffectedUser = $($script:Recipient.PrimarySMTPAddress)
        Collect-MigrationBatch -AffectedUser $AffectedUser -MigrationType Hybrid -ErrorAction SilentlyContinue
    }
<#
    if ($NumberOfChecks -le 5) {
        try {
            $script:MoveRequestStatistics = Get-EXOMoveRequestStatistics $AffectedUser -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" -ErrorAction Stop
        }
        catch {
            $NumberOfChecks++
            Write-Log "[WARNING] || ($NumberOfChecks / 5) We were unable to collect MoveRequestStatistics"
        }
    }
    else {
        Write-Log "[WARNING] || We were unable to find a MoveRequestStatistics in place for `"$AffectedUser`". We will check if MigrationUser was created for it."
        Collect-MigrationUserStatistics -AffectedUser $AffectedUser -MigrationType Hybrid
    }



    if (Get-EXOMigrationUser $AffectedUser -ErrorAction SilentlyContinue){
        $script:MigrationUserStatistics = Get-EXOMigrationUserStatistics $AffectedUser -IncludeSkippedItems -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose"
    }
#>
}

function Collect-MigrationBatch {
    [CmdletBinding()]
    param (
        [string]
        $AffectedUser,
        [string]
        $MigrationType
    )

    Write-Host "MigrationBatch"
    ### $script:MigrationBatch = Get-EXOMigrationBatch $AffectedUser -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose"
}


$AffectedUser = "NewRemoteMBX11"
Collect-MoveRequestStatistics -AffectedUser $AffectedUser -NumberOfChecks 1