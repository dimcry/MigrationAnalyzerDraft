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
        
        [ValidateSet("MoveRequestStatistics", "MoveRequest", "MigrationUserStatistics", "MigrationUser", "MigrationBatch", "SyncRequestStatistics", "SyncRequest", "MailboxStatistics", "FromFile")]
        [string]$LogType,

        [ValidateSet("Hybrid", "IMAP", "Cutover", "Staged")]
        [string]$MigrationType
    )

    # List of fields to output
    [Array]$OrderedFields = "BasicInformation","PerformanceStatistics","FailureSummary","FailureStatistics","LargeItemSummary","BadItemSummary","MailboxVerification"

    # Create the Result object that will be used to store all results
    $MoveAnalysis = New-Object PSObject
    $OrderedFields | foreach { $MoveAnalysis | Add-Member -Name $_ -Value $null -MemberType NoteProperty  }

    
    # Pull everything that we need that is common to all status types
    $MoveAnalysis.BasicInformation        = New-BasicInformation -RequestStats $($MigrationLogs.Logs)
    $MoveAnalysis.PerformanceStatistics   = New-PerformanceStatistics -RequestStats $($MigrationLogs.Logs)
    $MoveAnalysis.FailureSummary          = New-FailureSummary -RequestStats $($MigrationLogs.Logs)
    $MoveAnalysis.FailureStatistics       = New-FailureStatistics -FailureSummaries $MoveAnalysis.FailureSummary
    $MoveAnalysis.LargeItemSummary        = New-LargeItemSummary -RequestStats $($MigrationLogs.Logs)
    $MoveAnalysis.BadItemSummary          = New-BadItemSummary -RequestStats $($MigrationLogs.Logs)

    # Add fields that are not printed in the analysis
    $MoveAnalysis | Add-Member -NotePropertyName Report -NotePropertyValue $($MigrationLogs.Logs.Report)
    $MoveAnalysis | Add-Member -NotePropertyName DiagnosticInfo -NotePropertyValue $($MigrationLogs.Logs.DiagnosticInfo)

    $timelineMonth = Build-TimeTrackerTable -MrsJob $($MigrationLogs.Logs) -Aggregation Month
    $timelineDay = Build-TimeTrackerTable -MrsJob $($MigrationLogs.Logs) -Aggregation Day
    $timelineHour = Build-TimeTrackerTable -MrsJob $($MigrationLogs.Logs) -Aggregation Hour
    $timelineMinute = Build-TimeTrackerTable -MrsJob $($MigrationLogs.Logs) -Aggregation Minute

    $Timeline = New-Object PSObject
    $Timeline | Add-Member -NotePropertyName timelineMonth -NotePropertyValue $timelineMonth
    $Timeline | Add-Member -NotePropertyName timelineDay -NotePropertyValue $timelineDay
    $Timeline | Add-Member -NotePropertyName timelineHour -NotePropertyValue $timelineHour
    $Timeline | Add-Member -NotePropertyName timelineMinute -NotePropertyValue $timelineMinute

    $MoveAnalysis | Add-Member -NotePropertyName Timeline -NotePropertyValue $Timeline

    $DetailsAboutTheMove = New-Object PSObject
    $DetailsAboutTheMove | Add-Member -NotePropertyName Environment -NotePropertyValue $TheEnvironment
    $DetailsAboutTheMove | Add-Member -NotePropertyName LogFrom -NotePropertyValue $LogFrom
    $DetailsAboutTheMove | Add-Member -NotePropertyName LogType -NotePropertyValue $LogType
    $DetailsAboutTheMove | Add-Member -NotePropertyName MigrationType -NotePropertyValue $MigrationType
    $DetailsAboutTheMove | Add-Member -NotePropertyName PrimarySMTPAddress -NotePropertyValue $($MigrationLogs.PrimarySMTPAddress)

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

function Build-TimeTrackerTable
{
    <#
        .Synopsis
        Retrieves the set of MRS indexes in AD for jobs matching the given query.

        .Parameter MrsJob
        An object returned by Get-*RequestStatistics with detailed time-tracker data.
		These data are obtained by passing these arguments to the cmdlet: -Diagnostic -DiagnosticArgument "showtimeline,verbose"

        .Parameter Aggregation
        Build the table from the ByMinute, ByHour, ByDay or ByMonth XML aggregations.
    #>
    param(
        [Parameter(Mandatory=$true)]
        $MrsJob,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Minute', 'Hour', 'Day', 'Month')]
        [string]
        $Aggregation = 'Hour'
    )

    $diagnosticInfo = [xml]$MrsJob.DiagnosticInfo
    if ($diagnosticInfo -eq $null)
    {
        return
    }

	$seriesName = 'By{0}' -f $Aggregation
    $seriesData = $diagnosticInfo.Job.TimeTracker.Timeline.$seriesName
    if ($seriesData -eq $null -or $seriesData.$Aggregation.Count -eq 0)
    {
        return
    }

    $seriesSize = $seriesData.$Aggregation.D.Count
    $series = [System.Collections.Generic.List[object]]::new($seriesSize)
    foreach ($hour in $seriesData.$Aggregation)
    {
        $startTime = $hour.StartTime -as [DateTime]
        foreach ($entry in $hour.D)
        {
            $state = $entry.State
            $duration = $entry.Duration -as [TimeSpan]
			$msecs = $entry.MSecs -as [long]
            $row = [PSCustomObject][Ordered]@{
                'StartTime' = $startTime
                'State' = $state
                'Duration' = $duration
				'Milliseconds' = $msecs
				'CumulativeDuration' = $null
				'CumulativeMilliseconds' = $null
            }

            $series.Add($row)
        }
    }

	$series = $series | sort StartTime, State

	$gSeries = $series | group -NoElement State
	$accumulations = @{}
	$gSeries.Name | %{ $accumulations[$_] = [TimeSpan]::Zero }

    $series | %{
		$state = $_.State
		$accumulation = $accumulations[$state]
		$accumulation += $_.Duration
		$accumulations[$state] = $accumulation
		$_.CumulativeDuration = $accumulation
		$_.CumulativeMilliseconds = $accumulation.TotalMilliseconds
	}

	$series
}


$Data = Import-Clixml C:\1.xml
$TheInfo = Create-MoveObject -MigrationLogs $LogEntry -TheEnvironment 'Exchange Online' -LogFrom FromFile -LogType FromFile -MigrationType Hybrid

$MigrationUserStatistics = Import-Clixml C:\Temp\MigrationUserStatistics.xml

$LogEntry = New-Object PSObject
$LogEntry | Add-Member -NotePropertyName PrimarySMTPAddress -NotePropertyValue "FromFile"
$LogEntry | Add-Member -NotePropertyName Logs -NotePropertyValue $Data
$Entry = Create-MoveObject -MigrationLogs $LogEntry -TheEnvironment 'Exchange Online' -LogFrom FromFile -LogType FromFile -MigrationType Hybrid

[int]$TheDurationMilliseconds = 0
foreach ($TimelineEntry in $($Entry.Timeline.timelineMonth)) {
    $TheDurationMilliseconds = $TheDurationMilliseconds + $($TimelineEntry.Milliseconds)
}

$timelineMonthSorted = $($Entry.Timeline.timelineMonth) | sort Milliseconds -Descending | select -First 3
$timelineMonthSorted | ft -AutoSize
Write-Host "The job was impacted mostly, by the following:" -ForegroundColor Green
foreach ($timelineMonthSortedEntry in $timelineMonthSorted) {
    Write-Host "`t$($timelineMonthSortedEntry.State): " -ForegroundColor Cyan -NoNewline
    Write-Host "`t$($timelineMonthSortedEntry.Milliseconds)" -ForegroundColor White -NoNewline
    $ThePercent = (([int]$($timelineMonthSortedEntry.Milliseconds)/[int]$TheDurationMilliseconds)*100).ToString("#.##")
    Write-Host " ($ThePercent `%)"
}






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