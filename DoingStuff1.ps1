#region Helper Functions

# Generates enums used in the script if they are not defined
Function Ensure-EnumTypes
{
    try
    {
        # If the type is already loaded, return
        [AnalyzeMoveRequest] | Out-Null
        return;
    }
    catch { <# If they aren't loaded, we load them below #>  }

    # Add the types
    Add-Type -TypeDefinition @"
        namespace AnalyzeMoveRequest
        {
            public enum Severity
            {
                Info,
                Warning,
                Error
            }
        }
"@
}

# Gets the number of bytes in a ByteQuantifiedSize
Function Get-Bytes
{
    param ($datasize)

    try
    {
        $datasize.tobytes()
    }
    catch [Exception]
    {
        Parse-ByteQuantifiedSize $datasize
    }
}

# Parses a serialized ByteQuantifiedSize
Function Parse-ByteQuantifiedSize
{
    param ([Parameter(Mandatory = $true)][string]$SerializedSize)

    $result =  [regex]::Match($SerializedSize, '[^\(]+\((([0-9]+),?)+ bytes\)', [Text.RegularExpressions.RegexOptions]::Compiled)
    if ($result.Success)
    {
        [string]$extractedSize = ""
        $result.Groups[2].Captures | %{ $extractedSize += $_.Value }
        return [long]$extractedSize
    }

    return [long]0
}

# Returns totalseconds or 0 for durations
# This operates against: System.TimeSpan, M.E.Data.EnhancedTimeSpan and Deserialized.M.E.Data.EnhancedTimeSpan
Function DurationtoSeconds
{
    Param(
        [Parameter(Mandatory = $false)]
        $time = $null
    )

    if ($time -eq $null) { 0 }
    else { $time.TotalSeconds }
}

# Writes an empty line
Function Write-EmptyLine
{
    Write-Host "`t"
}

# Takes a string or multiple strings and prints every
# line in a seperate write-host call
Function Write-HostMultiline {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline=$true)]
        [string[]]$Content
    )

    PROCESS
    {
        foreach ($multilineString in $Content)
        {
            $multilineString.Split("`r`n", [StringSplitOptions]::RemoveEmptyEntries) | % { Write-Host $_ }
        }
    }
}

# Evaluates an expression and returns the result
# if an exception is thrown, returns a default value
Function Eval-Safe
{
    param(
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$Expression,

        [Parameter(Mandatory=$false)]
        $DefaultValue = $null
    )

    try
    {
        return (Invoke-Command -ScriptBlock $Expression)
    }
    catch
    {
        Write-Warning ("Eval-Safe: Error: '{0}'; returning default value: {1}" -f $_,$DefaultValue)
        return $DefaultValue
    }
}

#endregion

function Create-MoveObject {
    param (
        $MigrationLogs,
        
        [ValidateSet("Exchange Online", "Exchange OnPremises")]
        [string]$TheEnvironment,
        
        [ValidateSet("FromFile", "FromExchangeOnline", "FromExchangeOnPremises")]
        [string]$LogFrom,
        
        [ValidateSet("MoveRequestStatistics", "MigrationUserStatistics", "MigrationBatch", "MailboxStatistics")]
        [string]$LogType,

        [ValidateSet("RemoteMove", "IMAP", "Cutover", "Staged")]
        [string]$MigrationType
    )

    # List of fields to output
    [Array]$OrderedFields = "BasicInformation","PerformanceStatistics","FailureSummary","FailureStatistics","LargeItemSummary","BadItemSummary","MailboxVerification"

    # Create the Result object that will be used to store all results
    $MoveAnalysis = New-Object PSObject
    $OrderedFields | foreach { $MoveAnalysis | Add-Member -Name $_ -Value $null -MemberType NoteProperty  }

    
    # Pull everything that we need that is common to all status types
    $MoveAnalysis.BasicInformation        = New-BasicInformation -RequestStats $MigrationLogs
    $MoveAnalysis.PerformanceStatistics   = New-PerformanceStatistics -RequestStats $MigrationLogs
    $MoveAnalysis.FailureSummary          = New-FailureSummary -RequestStats $MigrationLogs
    $MoveAnalysis.FailureStatistics       = New-FailureStatistics -FailureSummaries $MoveAnalysis.FailureSummary
    $MoveAnalysis.LargeItemSummary        = New-LargeItemSummary -RequestStats $MigrationLogs
    $MoveAnalysis.BadItemSummary          = New-BadItemSummary -RequestStats $MigrationLogs

    # Add fields that are not printed in the analysis
    $MoveAnalysis | Add-Member -NotePropertyName Report -NotePropertyValue $MigrationLogs.Report
    $MoveAnalysis | Add-Member -NotePropertyName DiagnosticInfo -NotePropertyValue $MigrationLogs.DiagnosticInfo

    $DetailsAboutTheMove = New-Object PSObject
    $DetailsAboutTheMove | Add-Member -NotePropertyName Environment -NotePropertyValue $TheEnvironment
    $DetailsAboutTheMove | Add-Member -NotePropertyName LogFrom -NotePropertyValue $LogFrom
    $DetailsAboutTheMove | Add-Member -NotePropertyName LogType -NotePropertyValue $LogType
    $DetailsAboutTheMove | Add-Member -NotePropertyName MigrationType -NotePropertyValue $MigrationType

    $MoveAnalysis | Add-Member -NotePropertyName DetailsAboutTheMove -NotePropertyValue $DetailsAboutTheMove
    
    return $MoveAnalysis

}

# Create the Basic Information object and populate it with the baseline values
Function New-BasicInformation
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    # Build all properties to be added to the oubject
    New-Object PSObject -Property ([ordered]@{
        Alias           = $RequestStats.Alias 
        ExchangeGuid    = $RequestStats.ExchangeGuid
        Created         = $RequestStats.QueuedTimestamp
        Completed       = $RequestStats.CompletionTimeStamp
        Failed          = $RequestStats.FailureTimeStamp
        Direction       = ([String]$RequestStats.Direction)
        Status          = ([String]$RequestStats.Status)
        StatusDetail    = ([String]$RequestStats.StatusDetail)
        Workload        = ([String]$RequestStats.Workloadtype)
        Flags           = ([String]$RequestStats.Flags)
        SourceServer    = $RequestStats.SourceServer
        SourceVersion   = $RequestStats.SourceVersion
        SourceDatabase  = $RequestStats.SourceDatabase
        TargetServer    = $RequestStats.TargetServer
        TargetVersion   = $RequestStats.TargetVersion
        TargetDatabase  = $RequestStats.TargetDatabase
        MRSProxy        = $RequestStats.RemoteHostName
        RemoteDatabase  = $RequestStats.RemoteDatabaseName
        Protected       = $RequestStats.Protect
        BadItemLimit    = ([int][String]$RequestStats.BadItemLimit)
        LargeItemLimit  = ([int][String]$RequestStats.LargeItemLimit)
        Failures        = $RequestStats.Report.Failures
        BadItems        = $RequestStats.Report.BadItems
        LargeItems      = $RequestStats.Report.LargeItems
    })
}

# Add values to the object relevant to failed moves
Function Add-BasicInformationFailed
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [array]$BasicInformation
    )

    BEGIN
    {
        # Build all properties to be added to the oubject
        $Properties = [ordered]@{
            FailureTimestamp    = $RequestStats.FailureTimestamp
            FailureType         = $RequestStats.FailureType
            FailureSide         = ([String]$RequestStats.FailureSide)
        }
    }

    PROCESS
    {
        foreach ($info in $BasicInformation)
        {
            # Add them to the object
            $info | Add-Member -NotePropertyMembers $Properties
        }
    }
}

# Add values to the object relevant to completed moves
Function Add-BasicInformationComplete
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [array]$BasicInformation
    )

    BEGIN
    {
        $SourceMailboxSizeBytes = $null
        $RequestStats.report.sourcemailboxsize | Select-Object *size | Get-Member -MemberType noteproperty | foreach { $SourceMailboxSizeBytes = $SourceMailboxSizeBytes + $RequestStats.Report.SourceMailboxsize.($_.name) }

        $TargetMailboxSizeBytes = $null
        $RequestStats.report.targetmailboxsize | Select-Object *size | Get-Member -MemberType noteproperty | foreach { $TargetMailboxSizeBytes = $TargetMailboxSizeBytes + $RequestStats.Report.Targetmailboxsize.($_.name) }

        $Properties = [ordered]@{
            SourceMailboxSizeGB = $SourceMailboxSizeBytes / 1GB
            TargetMailboxSizeGB = $TargetMailboxSizeBytes / 1GB
            PercentMailboxBloat = (($TargetMailboxSizeBytes - $SourceMailboxSizeBytes) / $SourceMailboxSizeBytes) * 100
        }
    }

    PROCESS
    {
        foreach ($info in $BasicInformation)
        {
            # Add the properties to the object
            $info | Add-Member -NotePropertyMembers $Properties
        }
    }
}

# Build information for mailbox verification
Function New-MailboxVerification
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    # Pull the mailbox verification off of the report
    [Array]$Folders = $RequestStats.Report.MailboxVerification

    $Object = New-Object PSObject
    $BadFolderFound = $false

    # Loop thru and push just what we want to report into the object
    Foreach ($folder in $Folders)
    {
        # Set the folder name (need to do a bit of formating here)
        if ([String]::IsNullOrEmpty($folder.folderpath))
        {
            $Foldername = "Root"
        }
        else
        {
            $Foldername = $folder.folderpath.replace("/Top of Information Store/","")
        }

        if ($folder.source.count -gt $folder.target.count) 
        {
            [String]$Message = "Missing Items - Corrupt:" + $folder.corrupt.count + " Large:" + $folder.large.count + " Skipped:" + $folder.skipped.count + " Missing:" + $folder.missing.count; $BadFolderFound = $true
        }
        elseif ($folder.source.count -eq $folder.target.count)
        {
            [String]$Message = "Valid"
        }
        else
        {
            [String]$Message = "Target > source??"
        }

        # Push the Hash in to the object
        if ($Message -ne "Valid")
        {
            $Object | Add-Member -NotePropertyName $FolderName -NotePropertyValue $Message -Force
        }
    }

    if ($BadFolderFound -eq $false)
    {
        $Object | Add-Member -Name "No issues found" -Value "All Folders Verified" -Type NoteProperty
    }

    Return $Object
}

# Build the stats for performance
Function New-PerformanceStatistics
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    New-Object PSObject -Property ([ordered]@{
        MigrationDuration         = $RequestStats.TotalInProgressDuration
        TotalGBTransferred        = (Get-Bytes $RequestStats.BytesTransferred) / 1GB
        PercentComplete           = $RequestStats.PercentComplete
        DataTransferRateMBPerHour = Eval-Safe { (((Get-Bytes $RequestStats.BytesTransferred) / 1MB) / (DurationtoSeconds $RequestStats.TotalInProgressDuration)) * 3600 }
        DataTransferRateGBPerHour = Eval-Safe { (((Get-Bytes $RequestStats.BytesTransferred) / 1GB) / (DurationtoSeconds $RequestStats.TotalInProgressDuration)) * 3600 }
        AverageSourceLatency      = Eval-Safe { $RequestStats.report.sessionstatistics.sourcelatencyinfo.average }
        AverageDestinationLatency = Eval-Safe { $RequestStats.report.sessionstatistics.destinationlatencyinfo.average }
        SourceSideDuration        = Eval-Safe { $RequestStats.Report.SessionStatistics.SourceProviderInfo.TotalDuration }
        DestinationSideDuration   = Eval-Safe { $RequestStats.Report.SessionStatistics.DestinationProviderInfo.TotalDuration }
        PercentDurationIdle       = Eval-Safe { ((DurationToSeconds $RequestStats.TotalIdleDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationSuspended  = Eval-Safe { ((DurationToSeconds $RequestStats.TotalSuspendedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationFailed     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalFailedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationQueued     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalQueuedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationLocked     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalStalledDueToMailboxLockedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationTransient  = Eval-Safe { ((DurationToSeconds $RequestStats.TotalTransientFailureDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
    })
}

Function Add-PerformanceStatisticsComplete
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [array]$PerformanceStatistics
    )

    BEGIN
    {
            $TargetMailboxSizeBytes = $null
            $RequestStats.report.targetmailboxsize | Select-Object *size | Get-Member -MemberType noteproperty | foreach { $TargetMailboxSizeBytes = $TargetMailboxSizeBytes + $RequestStats.Report.Targetmailboxsize.($_.name) }

            $Properties = [ordered]@{
                TransferOverHeadPercent = Eval-Safe { (((Get-Bytes $RequestStats.BytesTransferred) - $TargetMailboxSizeBytes) / (Get-Bytes $RequestStats.BytesTransferred) ) * 100 } -DefaultValue 0
            }
    }

    PROCESS
    {
        foreach ($perfStat in $PerformanceStatistics)
        {
            # Add them to the object
            $perfStat | Add-Member -NotePropertyMembers $Properties
        }
    }
}

# Creates an object with just the failure message and timestamp
Function New-FailureSummary
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    # Create the object
    $compactFailures = @()

    # If we have no failures make sure we write something
    if ($RequestStats.report.failures -eq $null)
    {
        $compactFailures += New-Object PSObject -Property @{
            TimeStamp = [DateTime]::MinValue
            FailureType = "No Failures Found"
        }
    }
    # Pull out just what we want in the compact report
    else
    {
        $compactFailures += $RequestStats.report.failures | Select-object -Property TimeStamp,Failuretype,Message

        # Pull in the entries that indicate us starting a mailbox move
        $compactFailures += ($RequestStats.report.entries | where { $_.message -like "*examining the request*" } | 
        select-Object @{
            Name = "TimeStamp"; 
            Expression = { $_.CreationTime }
        },
        @{
            Name = "FailureType";
            Expression = { "-->MRSPickingUpMove" }
        },
        Message)
    }
    
    $compactFailures = $compactFailures | sort-Object -Property timestamp

    Return $compactFailures
}

# Creates a summary of what failures we saw and what count of that failure
# Need the object from New-FailureSummary as input
Function New-FailureStatistics
{
    Param(
        [parameter(Mandatory = $true)]
        $FailureSummaries
    )

    $FailureStats = New-Object PSObject

    $FailureSummaries | Group-Object FailureType | Sort-Object Count -Descending | foreach {

        # Skip the fake events we've inserted
        # when generating failure summaries
        if ($_.Name.StartsWith('-->'))
        {
            # In a pipeline Foreach-Object, you have to use
            # return instead of continue...
            return;
        }

        $FailureStats | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Count
    }

    return $FailureStats
}

Function New-LargeItemSummary
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    $entries = @()

    foreach ($li in $RequestStats.Report.LargeItems)
    {
        # Extract the current size and limitation from the exception message if we can
        $limitMatch = [Regex]::Match($li.Failure.Message, 'dwParam:\s+(?<limit>0x[0-9A-F]+)\s+Msg:\s+Limitation')
        if ($limitMatch.Success)
        {
            $limitInBytes = [Convert]::ToUInt64($limitMatch.Groups["limit"].Value, 16)
        }

        $sizeMatch = [Regex]::Match($li.Failure.Message, 'dwParam:\s+(?<size>0x[0-9A-F]+)\s+Msg:\s+CurrentSize')
        if ($sizeMatch.Success)
        {
            $sizeInBytes = [Convert]::ToUInt64($sizeMatch.Groups["size"].Value, 16)
        }

        $entries += New-Object PSObject -Property ([ordered]@{
            ItemSize = ("{0:F1} MB" -f ($sizeInBytes / 1024 / 1024))
            SizeLimit = ("{0:F1} MB" -f ($limitInBytes / 1024 / 1024))
            FolderName = $li.FolderName
            Subject = $li.Subject
        })
    }

    return @(,$entries)
}

Function New-BadItemSummary
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    $entries = @()
    foreach ($bi in $RequestStats.Report.BadItems)
    {
        $entries += New-Object PSObject -Property ([ordered]@{
            FailureType = $bi.Failure.FailureType
            Category = $bi.Category
            FolderName = $bi.FolderName
            Subject = $bi.Subject
        })
    }

    return @(,$entries)
}


#utility function used to check if a specific command exists
Function Test-CommandExists
{
    Param ($command)

    [bool]$result = $false 

    try
    {
        Get-Command $command -EA:Stop
        $result = $true
    }
    Catch
    {
        Write-Error "$command does not exist"
    }

    $result
} #end function test-CommandExists


$Data = Import-Clixml C:\1.xml
$TheInfo = Create-MoveObject -MigrationLogs $Data -TheEnvironment 'Exchange Online' -LogFrom FromFile -LogType MoveRequestStatistics -MigrationType RemoteMove

$MigrationUserStatistics = Import-Clixml C:\Temp\MigrationUserStatistics.xml
$TheInfo = Create-MoveObject -MigrationLogs $MigrationUserStatistics -TheEnvironment 'Exchange Online' -LogFrom FromFile -LogType MigrationUserStatistics -MigrationType RemoteMove

$TheInfo.GetType()
$TheInfo | Get-Member


$TheInfo.BasicInformation | select * -ExcludeProperty Failures
$TheInfo.BadItemSummary
$TheInfo.FailureStatistics
$TheInfo.FailureSummary
$TheInfo.LargeItemSummary
$TheInfo.MailboxVerification
$TheInfo.PerformanceStatistics
$TheInfo.DetailsAboutTheMove
# $TheInfo.Report
# $TheInfo.DiagnosticInfo