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


    # List of fields to output
    [Array]$OrderedFields = "BasicInformation","PerformanceStatistics","FailureSummary","FailureStatistics","LargeItemSummary","BadItemSummary","MailboxVerification"

    # Create the Result object that will be used to store all results
    $MoveAnalysis = New-Object PSObject
    $OrderedFields | foreach { $MoveAnalysis | Add-Member -Name $_ -Value $null -MemberType NoteProperty  }

    # Pull everything that we need that is common to all status types
    $MoveAnalysis.BasicInformation        = New-BasicInformation -RequestStats $MigrationLogs

    $MigrationLogs = Import-Clixml C:\1.xml

    Add-BasicInformationFailed -RequestStats $MigrationLogs -BasicInformation $MoveAnalysis.BasicInformation