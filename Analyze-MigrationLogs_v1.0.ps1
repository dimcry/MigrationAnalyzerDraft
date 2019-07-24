# MIT License
# 
# Copyright (c) 2019 Cristian Dimofte
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# This script is analyzing the migration reports

### Versions:
### - May 6, 2019 - v1.0 --- Analyzing output of Get-MoveRequestStatistics, from an existing .xml file


########################################
# Common space for script's parameters #
########################################
#region Parameters

Param(
    [Parameter(Position=0, ParameterSetName = "FilePath", Mandatory = $false)]
    [String]$FilePath = $null
)

#endregion Parameters


#####################################
# Common space for global variables #
#####################################
#region Global variables

### ValidationForWorkingDirectory (Scope: Script) variable is used to validate if the Working Directory was created, or not
[bool]$script:ValidationForWorkingDirectory = $true
### TheWorkingDirectory (Scope: Script) variable is used to list the exact value for the Working Directory
$script:TheWorkingDirectory = $null
### TheWorkingDirectorySavedData (Scope: Script) variable is used to list the exact value for the location in which we will save the needed outputs
[string]$script:TheWorkingDirectorySavedData = $null
### LogsToAnalyze (Scope: Script) variable will contain mailbox migration logs for all affected users
[System.Collections.ArrayList]$script:LogsToAnalyze = @()
### ParsedLogs (Scope: Script) variable will contain parsed mailbox migration logs for all affected users
[System.Collections.ArrayList]$script:ParsedLogs = @()

#endregion Global variables


##############################
# Common space for functions #
##############################
#region Functions

### <summary>
### Show-Header function is adding a header on the screen at the time the script starts
### </summary>
function Show-Header {

    $menuprompt = $null

    Clear-Host
    $title = "=== Mailbox migration analyzer ==="
    if (!($menuprompt)) 
    {
        $menuprompt+="="*$title.Length
    }
    Write-Host $menuprompt
    Write-Host $title
    Write-Host $menuprompt

}


### <summary>
### Create-WorkingDirectory function will create the Working Directory (desired location "%temp%\MigrationAnalyzer\<MMddyyyy_HHmmss>").
### In case we will not be able to create it as mentioned above, the Working Directory will be created on a path inserted from keyboard.
### The script will continue to work on the Working Directory location (Set-Location -Path $WorkingDirectoryToUse)
### </summary>
### <param name="NumberOfChecks">NumberOfChecks is used in case we will not be able to create the Working directory, and we will retry for 5 times.</param>
function Create-WorkingDirectory {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [int]
        $NumberOfChecks
    )
    
    Write-Host "We are creating the working folder with " -ForegroundColor Cyan -NoNewline
    Write-Host "`"<MMddyyyy_HHmmss>`"" -ForegroundColor White -NoNewline
    Write-Host " format, under " -ForegroundColor Cyan -NoNewline
    Write-Host "`"%temp%\MigrationAnalyzer\`"" -ForegroundColor White

    ### TheDateToUse variable is used to collect the current date&time in the "MMddyyyy_HHmmss" format.
    $TheDateToUse = (Get-Date).ToString("MMddyyyy_HHmmss")

    ### WorkingDirectory variable is initialized to "%temp%\MigrationAnalyzer\<MMddyyyy_HHmmss>".
    $WorkingDirectory = "$env:temp\MigrationAnalyzer\$TheDateToUse"

    ### Creating the Working directory in the desired format.
    if (-not (Test-Path $WorkingDirectory)) {
        try {
            $null = New-Item -ItemType Directory -Force -Path $WorkingDirectory -ErrorAction Stop            
            [string]$script:TheWorkingDirectory = $WorkingDirectory
        }
        catch {
            ### In case of error, we will retry to create the Working directory, for maximum 5 times.
            if ($NumberOfChecks -le 5) {
                if (Test-Path $WorkingDirectory){
                    [string]$script:TheWorkingDirectory = $WorkingDirectory
                }
                else {
                    $NumberOfChecks++
                    Create-WorkingDirectory -NumberOfChecks $NumberOfChecks    
                }
            }
            ### In case we will not be able to create the Working directory even after 5 times, the script will ask to
            ### insert, from the keyboard, the path where it can be created.
            else {
                Write-Host
                Write-Host "We were unable to create the working directory with " -ForegroundColor Red -NoNewline
                Write-Host "`"<MMddyyyy_HHmmss>`"" -ForegroundColor White -NoNewline
                Write-Host " format, under " -ForegroundColor Red -NoNewline
                Write-Host "`"%temp%\MigrationAnalyzer\`"" -ForegroundColor White
                Write-Host
                Write-Host "Please provide a location on which you have permissions to create folders/files." -ForegroundColor Cyan
                Write-Host "In it, we will log the actions the script will take." -ForegroundColor Cyan
                Write-Host "`tPath: " -ForegroundColor Cyan -NoNewline
                $WorkingDirectoryToUse = Read-Host
            
                ### If entered value will be empty, the script will exit.
                if (-not ($WorkingDirectoryToUse)) {
                    throw "No valid path was provided."
                }
                else {
                    ### Doing 1-time effort to create the Working Directory in the location inserted from keyboard
                    try {
                        $null = New-Item -ItemType Directory -Force -Path $WorkingDirectoryToUse -ErrorAction Stop
                        [string]$script:TheWorkingDirectory = $WorkingDirectoryToUse
                    }
                    catch {
                        ### In case of error, we will exit the script.
                        [bool]$script:ValidationForWorkingDirectory = $false
                        throw "We were unable to create the Working Directory: $WorkingDirectoryToUse"
                    }
                }
            }
        }
    }

    ### We successfully created a Working Directory. We will set it as current path (Set-Location -Path $WorkingDirectoryToUse)
    if ($script:TheWorkingDirectory) {
        Write-Host
        Write-Host "We successfully created the following working directory:" -ForegroundColor Green
        Write-Host "`tFull path: " -ForegroundColor Cyan -NoNewline
        Write-Host "`t$script:TheWorkingDirectory" -ForegroundColor White
        Write-Host "`tShort path:" -ForegroundColor Cyan -NoNewline
        $TheShortPath = ($script:TheWorkingDirectory -split "MigrationAnalyzer")[1]
        Write-Host "`t`%temp`%\MigrationAnalyzer$TheShortPath" -ForegroundColor White

        # Keep track of the old location so we can restore it at the end
        $script:OriginalLocation = Get-Location
        Set-Location -Path $script:TheWorkingDirectory
        Create-LogFile -WorkingDirectory $script:TheWorkingDirectory
    }

    ### Doing 1-time effort to create the SavedData folder under the Working Directory
    [string]$script:TheWorkingDirectorySavedData = $script:TheWorkingDirectory + "\SavedData"
    try {
        $null = New-Item -ItemType Directory -Force -Path $script:TheWorkingDirectorySavedData -ErrorAction Stop    
        Write-Log ("[INFO] || We successfully created the SavedData folder: $script:TheWorkingDirectorySavedData")
    }
    catch {
        ### In case of error, we will exit the script.
        throw "We were unable to create the Working Directory under: $script:TheWorkingDirectorySavedData"
    }
}

### <summary>
### Restores the original state of the console, like current directory, etc.
### </summary>
Function Restore-OriginalState {
    if ($script:OriginalLocation) {
        Set-Location -Path $script:OriginalLocation
    }
}

### <summary>
### Create-LogFile function creates the LogFile, in the Working Directory.
### </summary>
### <param name="WorkingDirectory">WorkingDirectory parameter is used get the location on which the LogFile will be created.</param>
function Create-LogFile {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $WorkingDirectory
    )

    ### LogFile variable (Scope: Script) is initialized to "$WorkingDirectory\MigrationAnalyzer.log".
    $script:LogFile = "$WorkingDirectory\MigrationAnalyzer.log"

    try {
        ### Creating the LogFile.
        $null = New-Item -ItemType "file" -Path "$script:LogFile" -Force -ErrorAction Stop
    }
    catch {
        ### In case of error, the script will exit.
        throw "You do not have permissions to create files under: $WorkingDirectory"
    }

    ### In case of success, we will log the first entry in the LogFile.
    Write-Log "[INFO] || Logfile successfully created. Its location is $script:LogFile"
}

### <summary>
### Write-Log function will add the entries in the LogFile, saved on the Working Directory. Also, it may display a log entry on the screen.
### </summary>
### <param name="string">string parameter is used to get the string that will be listed in the log file, and/or on the screen.</param>
### <param name="NonInteractive">if NonInteractive parameter is set to True, the information will be saved just on the LogFile. Else, it will be displayed
### on the screen, too.</param>
Function Write-Log {
    [CmdletBinding()]
    Param (
        [parameter(Position=0)]
        [string]
        $string,
        [parameter(Position=1)]
        [bool]
        $NonInteractive,
        [parameter(Position=2)]
        [ConsoleColor]
        $ForegroundColor = "White"
    )

    ### Collecting the current date
    [string]$date = Get-Date -Format G
        
    ### Write everything to LogFile

    if ($script:LogFile) {
        ( "[" + $date + "] || " + $string) | Out-File -FilePath $script:LogFile -Append
    }
    
    ### In case NonInteractive is not True, write on display, too
    if (!($NonInteractive)){
        Write-Host
        ("[" + $date + "] || " + $string) | Write-Host -ForegroundColor $ForegroundColor
    }
}

### <summary>
### Check-Parameters function is checking if the script was started with specific parameters.
### </summary>
function Check-Parameters {
    ### If FilePath parameter of the script was used, we will continue on this path.
    if ($FilePath){
        Write-Log ("[INFO] || The script was started with the FilePath parameter: `"-FilePath $FilePath`"")
        Selected-FileOption -FilePath $FilePath
    }
    ### If the script was started without any parameters, we will provide a menu in order to continue
    else {
        Show-Menu
    }
}


### <summary>
### Show-Menu function is used if the script is started without any parameters
### </summary>
### <param name="WorkingDirectory">WorkingDirectory parameter is used get the location on which the LogFile will be created.</param>
function Show-Menu {

    $menu=@"

1 => If you have the migration logs in an .xml file
Q => Quit

Select a task by number, or, Q to quit: 
"@

    Write-Log "[INFO] || Loading the menu..." -NonInteractive $true
    $null = Show-Header

    Write-Host $menu -ForegroundColor Cyan -NoNewline
    $SwitchFromKeyboard = Read-Host

    ### Providing a list of options
    Switch ($SwitchFromKeyboard) {

        ### If "1" is selected, the script will assume you have the mailbox migration logs in an .xml file
        "1" {
            Write-Log "[INFO] || You selected to provide an .xml to be analyzed."
            Selected-FileOption
        }

        ### If "Q" is selected, the script will exit
        "Q" {
            throw "You selected to quit the menu"
         }
 
        ### If you selected anything different than "1", "2", "3" or "Q", the Menu will reload
        default {
            Write-Log "[WARNING] || You selected an option that is not present in the menu (Value inserted from keyboard: `"$SwitchFromKeyboard`" / Expected values are `"1`" and `"Q`")" -ForegroundColor Yellow
            Write-Log "[INFO] || Press any key to re-load the menu"
            Read-Host
            Show-Menu
        }
    } 
}


### <summary>
### Selected-FileOption function is used when the information is already saved on a .xml file.
### </summary>
### <param name="FilePath">FilePath parameter is used when the script is started with the FilePath parameter.</param>
function Selected-FileOption {
    [CmdletBinding()]
    Param
    (        
        [string]$FilePath
    )

    [int]$TheNumberOfChecks = 1
    ### If FilePath was provided, the script will use it in order to validate if the information from this variable is a correct
    ### full path of an .xml file.
    if ($FilePath){
        try {
            ### The script validates that the path provided is of a valid .xml file.
            Write-Log "[INFO] || We are validating if `"$FilePath`" is the full path of a .xml file"
            [string]$PathOfXMLFile = Validate-XMLPath -XMLFilePath $FilePath
        }
        catch {
            ### In case of error, the script will ask to provide again the full path of the .xml file
            [string]$PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks
        }
    }
    ### If no FilePath was provided, the script will ask to provide the full path of the .xml file
    else{
        [string]$PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks
    }

    ### If PathOfXMLFile variable will match "NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed", we will continue the data collection
    ### using other methods.
    if ($PathOfXMLFile -match "NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed") {
        throw "The script will end, because the .xml file provided is not valid from PowerShell's perspective"
    }
    else {
        ### TheMigrationLogs variable will represent MigrationLogs collected using the Collect-MigrationLogs function.
        Write-Log "[INFO] || Trying to collect the Migration Logs using Selected-FileOption -> Collect-MigrationLogs function..."
        Collect-MigrationLogs -XMLFile $PathOfXMLFile
    }

}


### <summary>
### Validate-XMLPath function is used to check if the path provided is a valid .xml file.
### </summary>
### <param name="XMLFilePath">XMLFilePath parameter represents the path the script has to check if it is a valid .xml file.</param>
function Validate-XMLPath {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_})]
        [string]
        $XMLFilePath
    )

    ### Validating if the path has a length greater than 4, and if it is of an .xml file
    Write-Log "[INFO] || Checking if the FilePath is a valid .xml file, from PowerShell's perspective"
    if (($XMLFilePath.Length -gt 4) -and ($XMLFilePath -like "*.xml")) {
        ### Validating if the .xml file was created by PowerShell
        $fileToCheck = new-object System.IO.StreamReader($XMLFilePath)
        if ($fileToCheck.ReadLine() -like "*http://schemas.microsoft.com/powershell*") {
            Write-Host
            Write-Host $XMLFilePath -ForegroundColor Cyan -NoNewline
            Write-Host " seems to be a valid .xml file. We will use it to continue the investigation." -ForegroundColor Green
            Write-Log ("[INFO] || $XMLFilePath seems to be a valid .xml file. We will use it to continue the investigation.") -NonInteractive $true
        }
        ### If not, the script will set the XMLFilePath to NotAValidXMLFile. This will help in next checks, in order to start collecting the mailbox 
        ### migration logs using other methods
        else {
            Write-Log ("$XMLFilePath is not a valid .xml file. We will set: XMLFilePath = `"NotAValidXMLFile`"") -ForegroundColor Yellow
            $XMLFilePath = "NotAValidXMLFile"
        }
        
        $fileToCheck.Close()

    }
    ### If the path's length is not greater than 4 characters and the file is not an .xml file the script will set XMLFilePath to NotAValidPath.
    ### This will help in next checks, in order to start collecting the mailbox migration logs using other methods
    else {
        Write-Log ("[WARNING] || $XMLFilePath is not a valid .xml file. We will set: XMLFilePath = `"NotAValidPath`"") -ForegroundColor Yellow
        $XMLFilePath = "NotAValidPath"
    }
    
    ### The script returns the value of XMLFilePath 
    return $XMLFilePath
}

### <summary>
### Ask-ForXMLPath function is used to ask for the full path of a .xml file.
### </summary>
### <param name="NumberOfChecks">NumberOfChecks is used in order to do an 1-time effort to provide another path of the .xml file,
### in case first time when it was entered, there was a typo </param>
function Ask-ForXMLPath {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [int]$NumberOfChecks
    )

    [string]$PathOfXMLFile = ""
    if ($NumberOfChecks -eq "1") {
        ### Asking to provide the full path of the .xml file for the first time
        Write-Host
        Write-Log "[INFO] || We are asking to provide the path of the .xml file" -NonInteractive $true
        Write-Host "Please provide the path of the .xml file: " -ForegroundColor Cyan
        Write-Host "`t" -NoNewline
        try {
            ### PathOfXMLFile variable will contain the full path of the .xml file, if it will be validated (it will be inserted from the keyboard)
            $PathOfXMLFile = Validate-XMLPath -XMLFilePath (Read-Host)
        }
        catch {
            ### If error, the script is doing the 1-time effort to collect again the full path of the .xml file
            $NumberOfChecks++
            $PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $NumberOfChecks
        }
    }
    else {
        ### The script is doing the 1-time effort to collect again the full path of the .xml file
        Write-Host
        Write-Log "[INFO] || Asking to provide the path of the .xml file again" -NonInteractive $true
        Write-Host "Would you like to provide the path of the .xml file again?" -ForegroundColor Cyan
        Write-Host "`t[Y] Yes`t`t[N] No`t`t(default is `"N`"): " -NoNewline -ForegroundColor White
        $ReadFromKeyboard = Read-Host

        ### Checking if the path will be provided again, or no. If no, we will continue to collect the mailbox migration logs, using other methods.
        [bool]$TheKey = $false
        Switch ($ReadFromKeyboard) 
        { 
          Y {$TheKey=$true} 
          N {$TheKey=$false} 
          Default {$TheKey=$false} 
        }

        if ($TheKey) {
            ### If YES was selected, we are asking to provide the path of the .xml file again
            Write-Host
            Write-Host "Please provide again the path of the .xml file: " -ForegroundColor Cyan
            Write-Host "`t" -NoNewline
            try {
                ### Validating the path of the .xml file
                $PathOfXMLFile = Validate-XMLPath -XMLFilePath (Read-Host)
            }
            catch {
                ### If error, the script will set PathOfXMLFile to ValidationOfFileFailed, which will be used to collect the logs using other methods
                $PathOfXMLFile = "ValidationOfFileFailed"
            }
        }
        else {
            ### If NO was selected, the script will set PathOfXMLFile to ValidationOfFileFailed, which will be used to collect the logs using other methods
            $PathOfXMLFile = "ValidationOfFileFailed"
        }
    }
    
    ### The function returns the full path of the .xml file, or ValidationOfFileFailed
    return $PathOfXMLFile
}


### <summary>
### Collect-MigrationLogs function is used to collect the mailbox migration logs
### </summary>
### <param name="XMLFile">XMLFile represents the .xml file from which we want to import the mailbox migration logs </param>
### <param name="ConnectToExchangeOnline">ConnectToExchangeOnline parameter will be used to connect to Exchange Online, and collect the 
### needed mailbox migration logs, based on the migration type used </param>
### <param name="ConnectToExchangeOnPremises">ConnectToExchangeOnPremises parameter will be used to connect to Exchange On-Premises, and collect the 
### the output of Get-MailboxStatistics (the MoveHistory part), for the affected user </param>
function Collect-MigrationLogs {
    [CmdletBinding()]
    Param (
        [parameter(Mandatory=$true,
        ParameterSetName="XMLFile")]
        [string]$XMLFile
    )
    
    if ($XMLFile) {
        ### Importing data in the LogsToAnalyze (Scope: Script) variable
        Write-Log ("[INFO] || Importing data from `"$XMLFile`" file, in the LogsToAnalyze variable")
        $TheMigrationLogs = Import-Clixml $XMLFile
        foreach ($Log in $TheMigrationLogs) {
            $LogEntry = New-Object PSObject
            $LogEntry | Add-Member -NotePropertyName GUID -NotePropertyValue $($Log.MailboxIdentity.ObjectGuid.ToString())
            $LogEntry | Add-Member -NotePropertyName Name -NotePropertyValue $($Log.MailboxIdentity.Name.ToString())
            $LogEntry | Add-Member -NotePropertyName DistinguishedName -NotePropertyValue $($Log.MailboxIdentity.DistinguishedName.ToString())
            $LogEntry | Add-Member -NotePropertyName SID -NotePropertyValue $($Log.MailboxIdentity.SecurityIdentifierString.ToString())
            $LogEntry | Add-Member -NotePropertyName Logs -NotePropertyValue $Log

            $null = $script:LogsToAnalyze.Add($LogEntry)
        }
    }

    if ($script:LogsToAnalyze) {
        foreach ($LogEntry in $script:LogsToAnalyze) {
            $TheInfo = Create-MoveObject -MigrationLogs $LogEntry -TheEnvironment FromFile -LogFrom FromFile -LogType FromFile -MigrationType FromFile
            $null = $script:ParsedLogs.Add($TheInfo)
        }
    }
}



function Create-MoveObject {
    param (
        $MigrationLogs,
        
        [ValidateSet("Exchange Online", "Exchange OnPremises", "FromFile")]
        [string]$TheEnvironment,
        
        [ValidateSet("FromFile", "FromExchangeOnline", "FromExchangeOnPremises")]
        [string]$LogFrom,
        
        [ValidateSet("MoveRequestStatistics", "MoveRequest", "MigrationUserStatistics", "MigrationUser", "MigrationBatch", "SyncRequestStatistics", "SyncRequest", "MailboxStatistics", "FromFile")]
        [string]$LogType,

        [ValidateSet("Hybrid", "IMAP", "Cutover", "Staged", "FromFile")]
        [string]$MigrationType
    )

    # List of fields to output
    [Array]$OrderedFields = "BasicInformation","PerformanceStatistics","FailureSummary","FailureStatistics","LargeItemSummary","BadItemSummary","MailboxVerification"

    # Create the Result object that will be used to store all results
    $MoveAnalysis = New-Object PSObject
    $OrderedFields | foreach {$MoveAnalysis | Add-Member -Name $_ -Value $null -MemberType NoteProperty}
    
    # Pull everything that we need, that is common to all status types
    $MoveAnalysis.BasicInformation        = New-BasicInformation -RequestStats $($MigrationLogs.Logs)
    $MoveAnalysis.PerformanceStatistics   = New-PerformanceStatistics -RequestStats $($MigrationLogs.Logs)
    $MoveAnalysis.FailureSummary          = New-FailureSummary -RequestStats $($MigrationLogs.Logs)
<#    
    ##$MoveAnalysis.FailureStatistics       = New-FailureStatistics -FailureSummaries $MoveAnalysis.FailureSummary
    ##$MoveAnalysis.LargeItemSummary        = New-LargeItemSummary -RequestStats $($MigrationLogs.Logs)
    ##$MoveAnalysis.BadItemSummary          = New-BadItemSummary -RequestStats $($MigrationLogs.Logs)

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
#>
    $DetailsAboutTheMove = New-Object PSObject
    $DetailsAboutTheMove | Add-Member -NotePropertyName Environment -NotePropertyValue $TheEnvironment
    $DetailsAboutTheMove | Add-Member -NotePropertyName LogFrom -NotePropertyValue $LogFrom
    $DetailsAboutTheMove | Add-Member -NotePropertyName LogType -NotePropertyValue $LogType
    $DetailsAboutTheMove | Add-Member -NotePropertyName MigrationType -NotePropertyValue $MigrationType
    $DetailsAboutTheMove | Add-Member -NotePropertyName PrimarySMTPAddress -NotePropertyValue $($MigrationLogs.Name)

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
        Alias                           = ([String]$RequestStats.Alias)
        BadItemLimit                    = ([int][String]$RequestStats.BadItemLimit)
        BadItemsEncountered             = ([int][String]$RequestStats.BadItemsEncountered)
        Direction                       = ([String]$RequestStats.Direction)
        ExchangeGuid                    = ([String]$RequestStats.ExchangeGuid)
        Flags                           = ([String]$RequestStats.Flags)
        LargeItemLimit                  = ([int][String]$RequestStats.LargeItemLimit)
        LargeItemsEncountered           = ([int][String]$RequestStats.LargeItemsEncountered)
        OverallDuration                 = ([String]$RequestStats.OverallDuration)
        PercentComplete                 = ([int][string]$RequestStats.PercentComplete)
        Protected                       = ([String]$RequestStats.Protect)
        RemoteHostName                  = ([String]$RequestStats.RemoteHostName)
        RemoteDatabase                  = ([String]$RequestStats.RemoteDatabaseName)
        SourceArchiveDatabase           = ([String]$RequestStats.SourceArchiveDatabase)
        SourceArchiveServer             = ([String]$RequestStats.SourceArchiveServer)
        SourceArchiveVersion            = ([String]$RequestStats.SourceArchiveVersion)
        SourceDatabase                  = ([String]$RequestStats.SourceDatabase)
        SourceServer                    = ([String]$RequestStats.SourceServer)
        SourceVersion                   = ([String]$RequestStats.SourceVersion)
        Status                          = ([String]$RequestStats.Status)
        StatusDetail                    = ([String]$RequestStats.StatusDetail)
        TargetArchiveDatabase           = ([String]$RequestStats.TargetArchiveDatabase)
        TargetArchiveServer             = ([String]$RequestStats.TargetArchiveServer)
        TargetArchiveVersion            = ([String]$RequestStats.TargetArchiveVersion)
        TargetDatabase                  = ([String]$RequestStats.TargetDatabase)
        TargetServer                    = ([String]$RequestStats.TargetServer)
        TargetVersion                   = ([String]$RequestStats.TargetVersion)
        TotalArchiveItemCount           = ([UInt64][String]$RequestStats.TotalArchiveItemCount)
        TotalArchiveSize                = ([String]$RequestStats.TotalArchiveSize)
        TotalFailedDuration             = ([String]$RequestStats.TotalFailedDuration)
        TotalInProgressDuration         = ([String]$RequestStats.TotalInProgressDuration)
        TotalMailboxItemCount           = ([UInt64][String]$RequestStats.TotalMailboxItemCount)
        TotalMailboxSize                = ([String]$RequestStats.TotalMailboxSize)
        TotalQueuedDuration             = ([String]$RequestStats.TotalQueuedDuration)
        TotalSuspendedDuration          = ([String]$RequestStats.TotalSuspendedDuration)
        TotalTransientFailureDuration   = ([String]$RequestStats.TotalTransientFailureDuration)
        TotalPrimaryItemCount           = ([UInt64][String]$RequestStats.TotalPrimaryItemCount)
        TotalPrimarySize                = ([String]$RequestStats.TotalPrimarySize)
        WhenCreated                     = $RequestStats.QueuedTimeStamp
        WhenCompleted                   = $RequestStats.CompletionTimeStamp
        WhenFailed                      = $RequestStats.FailureTimeStamp
        Workload                        = ([String]$RequestStats.Workloadtype)
    })
}


# Build the stats for performance
Function New-PerformanceStatistics
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    New-Object PSObject -Property ([ordered]@{
        AverageSourceLatency      = Eval-Safe { $RequestStats.report.sessionstatistics.sourcelatencyinfo.average }
        AverageDestinationLatency = Eval-Safe { $RequestStats.report.sessionstatistics.destinationlatencyinfo.average }

        DataTransferRateMBPerHour = Eval-Safe { (((Get-Bytes $RequestStats.BytesTransferred) / 1MB) / (DurationtoSeconds $RequestStats.TotalInProgressDuration)) * 3600 }
        DataTransferRateGBPerHour = Eval-Safe { (((Get-Bytes $RequestStats.BytesTransferred) / 1GB) / (DurationtoSeconds $RequestStats.TotalInProgressDuration)) * 3600 }
        DestinationSideDuration   = Eval-Safe { $RequestStats.Report.SessionStatistics.DestinationProviderInfo.TotalDuration }
        MigrationDuration         = $RequestStats.TotalInProgressDuration
        PercentDurationIdle       = Eval-Safe { ((DurationToSeconds $RequestStats.TotalIdleDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationSuspended  = Eval-Safe { ((DurationToSeconds $RequestStats.TotalSuspendedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationFailed     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalFailedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationQueued     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalQueuedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationLocked     = Eval-Safe { ((DurationToSeconds $RequestStats.TotalStalledDueToMailboxLockedDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        PercentDurationTransient  = Eval-Safe { ((DurationToSeconds $RequestStats.TotalTransientFailureDuration) / (DurationtoSeconds $RequestStats.OverallDuration)) * 100 } -DefaultValue 0
        SourceSideDuration        = Eval-Safe { $RequestStats.Report.SessionStatistics.SourceProviderInfo.TotalDuration }
        TotalGBTransferred        = (Get-Bytes $RequestStats.BytesTransferred) / 1GB

    })
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


# Creates an object with just the failure message and timestamp
Function New-FailureSummary
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    # Build all properties to be added to the oubject
    if (([int]$RequestStats.Report.Failures.Count) -gt 0) {
        New-Object PSObject -Property ([ordered]@{
            FailuresCount           = ([int][string]$RequestStats.Report.Failures.Count)
            LatestTimeStamp         = $RequestStats.Report.Failures[-1].TimeStamp
            LatestFailureType       = ([string]$RequestStats.Report.Failures[-1].FailureType)
            LatestMessage           = (([string]$RequestStats.Report.Failures[-1].Message -split "`n")[0])
        })
    }
    else {
        New-Object PSObject -Property ([ordered]@{
            FailuresCount           = ([int][string]$RequestStats.Report.Failures.Count)
        })
    }
}

Function Old_New-FailureSummary
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
        $compactFailures += $RequestStats.report.failures | Select-object -Property TimeStamp, Failuretype, @{n='Message'; e={($($_.Message) -Split "`n")[0]}}
<#
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
#>
    }
    
    $compactFailures = $compactFailures | sort-Object -Property timestamp

    Return $compactFailures
}


function Get-RelevantFailures {
    param (
        $MigrationLogs
    )
    
    $LastPercentComplete = $MigrationLogs.Report.Entries | where {$_.Message -like "*Percent complete*"} | select -Last 1
    $RelevantFailures = $MigrationLogs.Report.Entries | where {($_.CreationTime -ge $LastPercentComplete.CreationTime) -and ($_.Type -eq "Error")}
}


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



#endregion Functions




###############
# Main script #
###############
#region Main script

try {
    Clear-Host

    $null = Show-Header
    Create-WorkingDirectory -NumberOfChecks 1
    Check-Parameters


    #region ForTestPurposes - This will be removed

    Write-Host
    Write-Host "Details from the mailbox migration log:" -ForegroundColor Green

    foreach ($Entry in $script:ParsedLogs) {
        if ($($Entry.DetailsAboutTheMove.PrimarySMTPAddress) -eq "FromFile")
        {
            Write-Log "[INFO] || Providing details of the migration from the file you provided"
        }
        else {
            Write-Log "[INFO] || Providing details of the migration for $($Entry.DetailsAboutTheMove.PrimarySMTPAddress)"
        }

        if ($($Entry.DetailsAboutTheMove.PrimarySMTPAddress)) {
            Write-Host "PrimarySMTPAddress: " -ForegroundColor Cyan
            Write-Host "$($Entry.DetailsAboutTheMove.PrimarySMTPAddress)" -ForegroundColor White
            Write-Log ("PrimarySMTPAddress: `n$($Entry.DetailsAboutTheMove.PrimarySMTPAddress)") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.BasicInformation)) {
            Write-Host "Basic Information: " -ForegroundColor Cyan
            $($Entry.BasicInformation)
            Write-Log ("Basic Information: `n$($Entry.BasicInformation)") -NonInteractive $true
            Write-Host
            $($Entry.BasicInformation) | ConvertTo-Html -As List | Out-File $script:TheWorkingDirectory\BasicInformation.html
        }

        if ($($Entry.PerformanceStatistics)) {
            Write-Host "Performance Statistics: " -ForegroundColor Cyan
            ($Entry.PerformanceStatistics)
            Write-Log ("Performance Statistics: `n$($Entry.PerformanceStatistics)") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.FailureSummary)) {
            Write-Host "Failure Summary: " -ForegroundColor Cyan
            $($Entry.FailureSummary)
            Write-Log ("Failure Summary: `n$($Entry.FailureSummary)") -NonInteractive $true
            Write-Host
        }

<# Just testing...
$RequestStats = Import-Clixml C:\1.xml
$FailuresCount1 = ([int][string]$RequestStats.Report.Failures.Count)
$FailuresCount1.GetType()
$FailuresCount1

$FailuresLatestTimeStamp         = ($RequestStats.Report.Failures[-1].TimeStamp)
$FailuresLatestTimeStamp.GetType()

$FailuresLatestFailureType       = ([string]$RequestStats.Report.Failures[-1].FailureType)
$FailuresLatestFailureType.GetType()

$FailuresLatestMessage           = (([string]$RequestStats.Report.Failures[-1].Message -split "`n")[0])
$FailuresLatestMessage.GetType()

$PercentComplete                 = $RequestStats.PercentComplete
$PercentComplete.GetType()


Function New-FailureSummary
{
    Param(
        [Parameter(Mandatory = $true)]
        $RequestStats
    )

    # Build all properties to be added to the oubject
    New-Object PSObject -Property ([ordered]@{
        FailuresCount                   = ([int][string]$RequestStats.Report.Failures.Count)
        FailuresLatestTimeStamp         = $RequestStats.Report.Failures[-1].TimeStamp
        FailuresLatestFailureType       = ([string]$RequestStats.Report.Failures[-1].FailureType)
        FailuresLatestMessage           = (([string]$RequestStats.Report.Failures[-1].Message -split "`n")[0])
    })
}

### <summary>
### Write-Log function will add the entries in the LogFile, saved on the Working Directory. Also, it may display a log entry on the screen.
### </summary>
### <param name="string">string parameter is used to get the string that will be listed in the log file, and/or on the screen.</param>
### <param name="NonInteractive">if NonInteractive parameter is set to True, the information will be saved just on the LogFile. Else, it will be displayed
### on the screen, too.</param>
Function Write-Log {
    [CmdletBinding()]
    Param (
        [parameter(Position=0)]
        [string]
        $string,
        [parameter(Position=1)]
        [bool]
        $NonInteractive,
        [parameter(Position=2)]
        [ConsoleColor]
        $ForegroundColor = "White"
    )

    ### Collecting the current date
    [string]$date = Get-Date -Format G
        
    ### Write everything to LogFile

    if ($script:LogFile) {
        ( "[" + $date + "] || " + $string) | Out-File -FilePath $script:LogFile -Append
    }
    
    ### In case NonInteractive is not True, write on display, too
    if (!($NonInteractive)){
        Write-Host
        ("[" + $date + "] || " + $string) | Write-Host -ForegroundColor $ForegroundColor
    }
}


[Array]$OrderedFields = "BasicInformation","PerformanceStatistics","FailureSummary","FailureStatistics","LargeItemSummary","BadItemSummary","MailboxVerification"

$MoveAnalysis = New-Object PSObject
$OrderedFields | foreach { $MoveAnalysis | Add-Member -Name $_ -Value $null -MemberType NoteProperty  }
$MoveAnalysis.FailureSummary          = New-FailureSummary -RequestStats $RequestStats

$Entry = $MoveAnalysis[0]
$Entry.FailquestStatistureSummary

        if ($($Entry.FailureSummary)) {
            Write-Host "Failure Summary: " -ForegroundColor Cyan
            $($Entry.FailureSummary) | fl
            Write-Log ("Failure Summary: `n$($Entry.FailureSummary)") -NonInteractive $true
            Write-Host
        }

$Entry.FailureSummary | Get-Member -MemberType NoteProperty
foreach ($NewEntry in ($($Entry.FailureSummary) | Get-Member -MemberType NoteProperty).Name) {
    Write-Host $NewEntry":`t" -NoNewline
    $($Entry.FailureSummary.$NewEntry)
}



[Array]$TheObject = (((Get-Content .\Log.log) -split "`{")[2] -split "; ")

$TheObject.GetType()





    $MoveRequestStatistics = Import-Clixml C:\1.xml
    $MoveRequestStatistics = $RequestStats

    $MoveRequestStatistics.Report.Entries | ft -AutoSize CreationTime, ServerName, Type, Flags, Message
    $ConfigObjectSourceBefore = ($MoveRequestStatistics.Report.Entries | where {($_.Flags -like "*Before*") -and ($_.Flags -like "*Source*")}).ConfigObject

    $ConfigObjectSourceBefore.Props | select PropertyName, @{n='Values'; e={(($($_.Values) -replace "{","") -replace "}","")}}

    $MoveRequestStatistics.Report.Entries | fl CreationTime, ServerName, Type, Flags, Message, Failure, BadItem, MailboxSize, SessionStatistics, ArchiveSessionStatistics, MailboxVerificationResults, DebugData, Connectivity, SourceThrottleDurations, TargetThrottleDurations

    $MoveRequestStatistics.Report.Entries | select -Last 11 -Skip 10 | fl
$SessionStatistics = $MoveRequestStatistics.Report.Entries | where {$_.Flags -like "*SessionStatistics*"}
$SessionStatistics.Count
$SessionStatistics | fl LocalizedString
$SessionStatistics[-6]


$LastPercentComplete = $RequestStats.Report.Entries | where {$_.Message -like "*Percent complete*"} | select -Last 1
$LastPercentComplete

$MigrationLogs = $RequestStats
$LastPercentComplete = $MigrationLogs.Report.Entries | where {$_.Message -like "*Percent complete*"} | select -Last 1
    $RelevantFailures = $MigrationLogs.Report.Entries | where {($_.CreationTime -ge $LastPercentComplete.CreationTime) -and ($_.Type -eq "Error")}
    $RelevantFailures.count
#>

        if ($($Entry.Timeline.timelineMinute)) {
            Write-Host "Details about timeline, by minute: " -ForegroundColor Cyan
            $TheEntriesToDisplay = $($Entry.Timeline.timelineMinute)
            $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 5 | ft -AutoSize
            $TheSortedEntriesToDisplay = $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 50
            Write-Log ("Details about timeline, by month: `n$TheSortedEntriesToDisplay") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.Timeline.timelineHour)) {
            Write-Host "Details about timeline, by hour: " -ForegroundColor Cyan
            $TheEntriesToDisplay = $($Entry.Timeline.timelineHour)
            $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 5 | ft -AutoSize
            $TheSortedEntriesToDisplay = $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 50
            Write-Log ("Details about timeline, by month: `n$TheSortedEntriesToDisplay") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.Timeline.timelineDay)) {
            Write-Host "Details about timeline, by Day: " -ForegroundColor Cyan
            $TheEntriesToDisplay = $($Entry.Timeline.timelineDay)
            $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 5 | ft -AutoSize
            $TheSortedEntriesToDisplay = $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 50
            Write-Log ("Details about timeline, by month: `n$TheSortedEntriesToDisplay") -NonInteractive $true
            Write-Host
        }

        if ($($Entry.Timeline.timelineMonth)) {
            Write-Host "Details about timeline, by month: " -ForegroundColor Cyan
            $TheEntriesToDisplay = $($Entry.Timeline.timelineMonth)
            $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 5 | ft -AutoSize
            $TheSortedEntriesToDisplay = $TheEntriesToDisplay | sort Milliseconds -Descending | select -First 50
            Write-Log ("Details about timeline, by month: `n$TheSortedEntriesToDisplay") -NonInteractive $true
            Write-Host
        }

        Write-Host
    }

    if ($script:TheWorkingDirectory) {
        Write-Host
        Write-Host "Useful logs can be found on the following directory:" -ForegroundColor Green
        Write-Host "`tFull path: " -ForegroundColor Cyan -NoNewline
        Write-Host "`t$script:TheWorkingDirectory" -ForegroundColor White
        Write-Host "`tShort path:" -ForegroundColor Cyan -NoNewline
        $TheShortPath = ($script:TheWorkingDirectory -split "MigrationAnalyzer")[1]
        Write-Host "`t`%temp`%\MigrationAnalyzer$TheShortPath" -ForegroundColor White
    }

    #endregion ForTestPurposes - This will be removed
    
} 
catch {
    Write-Log "[ERROR] || $_" -ForegroundColor Red
    Write-Log "[ERROR] || Script will now exit" -ForegroundColor Red
}
finally {
    Restore-OriginalState
    Get-PSSession | Remove-PSSession
}

#endregion Main script

############################
#####################################
# Create / update .xml / .json file #
#####################################


##########################
# Run the script section #
##########################
<#
### Analyze migration logs from .xml file:
.\Analyze-MigrationLogs.ps1 -FilePath C:\1.xml

### Analyze MoveRequestStatistics, from Exchange Online:
.\Analyze-MigrationLogs.ps1 -ConnectToExchangeOnline
.\Analyze-MigrationLogs.ps1 -ConnectToExchangeOnline -EXOAdminAccount Administrator@dimcryro.onmicrosoft.com
.\Analyze-MigrationLogs.ps1 -ConnectToExchangeOnline -AffectedUsers dtest5, dtest6, dtest2 -EXOAdminAccount Administrator@dimcryro.onmicrosoft.com
.\Analyze-MigrationLogs.ps1 -ConnectToExchangeOnline -AffectedUsers dtest5, dtest6, dtest2 -EXOAdminAccount Administrator@dimcryro.onmicrosoft.com -MigrationType Hybrid

### Analyze MailboxStatistics, from Exchange OnPremises, from a machine from external network:
.\Analyze-MigrationLogs.ps1 -ConnectToExchangeOnPremises -ExchangeURL https://owa.dimcry.ro/PowerShell -AffectedUsers dtest5, dtest6, dtest2 -OnPremAdminAccount dimcry\dimcry

### Analyze MailboxStatistics, from Exchange OnPremises, directly from a machine from internal network:
.\Analyze-MigrationLogs.ps1 -ConnectToExchangeOnPremises -AffectedUsers dtest5, dtest6, dtest2
#>