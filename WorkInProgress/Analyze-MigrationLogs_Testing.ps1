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

##############################################
# Comments during the time script is in BETA #
##############################################
#region The logic

<#
The logic:
==========

0. BEGIN
1. Do we have the migration logs in .xml format?
    1.1. Yes - Go to 2.
    1.2. No - Go to 4.

2. Import the .xml file into a PowerShell variable. Is this information an output of <Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose">?
    2.1. Yes - Go to 3.
    2.2. No - Go to 4.

3. Mark into the logs that the outputs were correctly collected, and we can start to analyze them. Go to 10.
4. Ask to provide user for which to analyze the migration logs. Go to 5.
5. Are we connected to Exchange Online?
    5.1. Yes - Go to 7.
    5.2. No - Go to 6.

6. Connect to Exchange Online using a Global Administrator. Go to 7.
7. Is the move request still present for this user?
    7.1. Yes - Go to 8.
    7.2. No - Go to 9.

8. Import into a PowerShell variable the output of <Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose">. Go to 3.
9. Is the object a Mailbox in Exchange Online?
    9.1. Yes - Import the correct move request from the MoveHistory. Mark into the logs that the output was collected from MoveHistory. Go to 10.
    9.2. No - Are we connected to Exchange On-Premises?
        9.2.1. Yes - Go to 9.1.
        9.2.2. No - Inform that the user should have a Mailbox in On-Premises. Ask to run the same script in Exchange Management Shell, in On-Premises. Go to END.

10. Download the .xml/.json file containing the pairs, from GitHub. Go to 11.
11. Analyze the logs and provide the results. Go to END.
12. END
#>

<#
A.    Necessary modules:
    1.    Collect the migration logs (related to one or multiple affected mailboxes):
        a.    From an existing .xml file;
        b.    From Exchange Online, by using the correct command:
            i.      For Hybrid;
            ii.     For IMAP;
            iii.    For Cutover / Staged.

        c.    From Get-MailboxStatistics output, if we speak about Remote moves (Hybrid), in case customer already removed the MoveRequest:
            i.    From Exchange Online, if we speak about an Onboarding;
            ii.    From Exchange On-Premises, if we speak about an Offboarding.

    2.    Download the JSON file from GitHub, and, based on the error received, based on the Migration type, we will provide recommendation about the actions that they can take to solve the issue.

B.    Good to have modules:
    1.    Performance analyzer. Similar to what Karahan provided in his script;
    2.    DiagnosticInfo analyzer.
        a.    Using Build-TimeTrackerTable function from Angus’s module, I’ll parse the DiagnosticInfo details, and provide some information to customer.
        b.    Using the idea described here, I’ll create a function that will provide a Column/Bar Chart similar to (this is screen shot provided by Angus long time ago, from a Pivot Table created in Excel, based on some information created with the above mentioned function):

            EURPRD10> $timeline = Build-TimeTrackerTable -MrsJob $stat
            EURPRD10> $timeline | Export-Csv 'tmp.csv'


C.    Priority of modules:
    Should be present in Version 1:     A.1., A.2., B.2.a.
    Can be introduced in Version 2.:    B.1., B.2.b.


D.    Resource estimates:

    From time perspective:
        Task name   Working hours   Expected completion time
        A.1.        24              31.01.2019
        A.2.        24              15.02.2019
        B.1.        112             30.04.2019
        B.2.a.      8               28.02.2019
        B.2.b.      112             31.05.2019


    From people perspective:
        For the moment I’ll do everything on my own.
        I asked pieces of advice from Brad Hughes, who guided me on the right direction (what type of the file should we use, how to download the JSON file from GitHub).
        If you can find any other resource that can help on any of the mentioned modules, I’ll be more than happy to add them into the “team” 😊
#>

#endregion The logic

########################################
# Common space for script's parameters #
########################################
#region Parameters
Param(
    [Parameter(ParameterSetName = "FilePath", Mandatory = $false)]
    [String]$FilePath = $null,

    [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $true)]
    [switch]$ConnectToExchangeOnline,

    [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $true)]
    [switch]$ConnectToExchangeOnPremises,

    [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
    [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
    [string[]]$AffectedUsers,
    
    [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
    [ValidateSet("Hybrid", "IMAP", "Cutover", "Staged")]
    [string]$MigrationType,

    [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
    [string]$EXOAdminAccount,

    [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
    [string]$OnPremAdminAccount,

    [Parameter(Mandatory=$false,
        ParameterSetName='ConnectToExchangeOnPremises')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({$_ -match "[htp]{4}"})] 
    [string]$ExchangeURL,

    [Parameter(Mandatory=$false,
        ParameterSetName='ConnectToExchangeOnPremises')]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Basic","Digest","Negotiate","Kerberos")]
    [string]$AuthenticationType,

    [Parameter(Mandatory=$false,
        ParameterSetName='ConnectToExchangeOnPremises')]
    [ValidateNotNullOrEmpty()]
    [string]$Domain = "getdomain",

    [Parameter(Mandatory=$false,
        ParameterSetName='ConnectToExchangeOnPremises')]
    [ValidateNotNullOrEmpty()]
    [string]$ADSite = "getsite"
)

#endregion Parameters

################################################
# Common space for functions, global variables #
################################################
#region Functions, Global variables

### LogsToAnalyze (Scope: Script) variable will contain mailbox migration logs for all affected users
[System.Collections.ArrayList]$script:LogsToAnalyze = @()
### EXOCommandsPrefix (Scope: Script) variable will be used to create a new PSSession to Exchange Online.
### When importing the PSSession, the script will use "MAEXO" (Migration Analyzer EXO) as Prefix for each command
[string]$script:EXOCommandsPrefix = "MAEXO"
### EXOPSSessionCreated (Scope: Script) variable will be used to check if the Exchange Online PSSession was successfully created
[bool]$script:EXOPSSessionCreated = $false
### ExOnPremCommandsPrefix (Scope: Script) variable will be used to create a new PSSession to Exchange OnPremises.
### When importing the PSSession, the script will use "MAExOnP" (Migration Analyzer Exchange OnPremises) as Prefix for each command
[string]$script:ExOnPremCommandsPrefix = "MAExOnP"
### ExOnPremPSSessionCreated (Scope: Script) variable will be used to check if the Exchange OnPremises PSSession was successfully created
[bool]$script:ExOnPremPSSessionCreated = $false


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
### Show-Menu function is used if the script is started without any parameters
### </summary>
### <param name="WorkingDirectory">WorkingDirectory parameter is used get the location on which the LogFile will be created.</param>
function Show-Menu {

    $menu=@"

1 => If you have the migration logs in an .xml file
2 => If you want to connect to Exchange Online in order to collect the logs
3 => If you need to connect to Exchange On-Premises and collect the logs
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

        ### If "2" is selected, the script will connect you to Exchange Online
        "2" {
            Write-Log "[INFO] || You selected to connect to Exchange Online and collect from there correct migration logs to be analyzed."
            Selected-ConnectToExchangeOnlineOption
        }
 
        ### If "3" is selected, you started the script from On-Premises Exchange Management Shell
        "3" {
            Write-Log "[INFO] || You selected to connect to Exchange On-Premises and collect from there correct migration logs to be analyzed."
            Selected-ConnectToExchangeOnPremisesOption
        }

        ### If "Q" is selected, the script will exit
        "Q" {
            throw "You selected to quit the menu"
         }
 
        ### If you selected anything different than "1", "2", "3" or "Q", the Menu will reload
        default {
            Write-Log "[WARNING] || You selected an option that is not present in the menu (Value inserted from keyboard: `"$SwitchFromKeyboard`")" -ForegroundColor Yellow
            Write-Log "[INFO] || Press any key to re-load the menu"
            Read-Host
            Show-Menu
        }
    } 
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
            $void = New-Item -ItemType Directory -Force -Path $WorkingDirectory -ErrorAction Stop            
            $WorkingDirectoryToUse = $WorkingDirectory
        }
        catch {
            ### In case of error, we will retry to create the Working directory, for maximum 5 times.
            if ($NumberOfChecks -le 5) {
                if (Test-Path $WorkingDirectory){
                    $WorkingDirectoryToUse = $WorkingDirectory
                }
                else {
                    $NumberOfChecks++
                    $WorkingDirectoryToUse = Create-WorkingDirectory -NumberOfChecks $NumberOfChecks    
                }
            }
            ### In case we will not be able to create the Working directory even after 5 times, we will set the value of WorkingDirectoryToUse
            ### variable to NotAbleToCreateTheWorkingDirectory.
            else {
                $WorkingDirectoryToUse = "NotAbleToCreateTheWorkingDirectory"
            }
        }
    }

    ### Checking if we were able to create the Working Directory in the desired location. If not, we will ask to insert the path where it can be created,
    ### from the keyboard.
    if ($WorkingDirectoryToUse -eq "NotAbleToCreateTheWorkingDirectory") {
        Write-Host
        Write-Host "We were unable to create the working directory with " -ForegroundColor Red -NoNewline
        Write-Host "`"<MMddyyyy_HHmmss>`"" -ForegroundColor White -NoNewline
        Write-Host " format, under " -ForegroundColor Red -NoNewline
        Write-Host "`"%temp%\MigrationAnalyzer\`"" -ForegroundColor White
        Write-Host
        Write-Host "Please provide a location on which you have permissions to create folders/files." -ForegroundColor Cyan
        Write-Host "In it we will log the actions the script will take." -ForegroundColor Cyan
        Write-Host "`tPath: " -ForegroundColor Cyan -NoNewline
        $WorkingDirectoryToUse = Read-Host
    
        ### If entered value will be empty, the script will exit.
        if (-not ($WorkingDirectoryToUse)) {
            throw "No valid path was provided."
        }
        else {
            ### Doing 1-time effort to create the Working Directory in the location inserted from keyboard
            try {
                $void = New-Item -ItemType Directory -Force -Path $WorkingDirectoryToUse -ErrorAction Stop            
            }
            catch {
                ### In case of error, we will exit the script.
                throw "We were unable to create the Working Directory under: $WorkingDirectoryToUse"
            }
        }
    }
    ### We successfully created a Working Directory. We will set it as current path (Set-Location -Path $WorkingDirectoryToUse)
    else {
        Write-Host
        Write-Host "We successfully created the following working directory:" -ForegroundColor Green
        Write-Host "`tFull path: " -ForegroundColor Cyan -NoNewline
        Write-Host $WorkingDirectoryToUse -ForegroundColor White
        Write-Host "`tShort path: " -ForegroundColor Cyan -NoNewline
        $TheShortPath = ($WorkingDirectoryToUse -split "MigrationAnalyzer")[1]
        Write-Host "`%temp`%\MigrationAnalyzer$TheShortPath" -ForegroundColor White

        # Keep track of the old location so we can restore it at the end
        $script:OriginalLocation = Get-Location
        Set-Location -Path $WorkingDirectoryToUse
    }

    ### Create-WorkingDirectory function will return the Path of the Working directory, or NotAbleToCreateTheWorkingDirectory in case
    ### we were unable to create the Working directory.
    return $WorkingDirectoryToUse
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
        $void = New-Item -ItemType "file" -Path "$script:LogFile" -Force -ErrorAction Stop
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
        ( "[" + $date + "] || " + $string) | Write-Host -ForegroundColor $ForegroundColor
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
    ### If ConnectToExchangeOnline parameter of the script was used, we will continue on this path.
    elseif ($ConnectToExchangeOnline) {
        Write-Log ("[INFO] || The script was started with the ConnectToExchangeOnline parameter: `"-ConnectToExchangeOnline:$true -AffectedUsers $AffectedUsers -MigrationType $MigrationType -TheAdminAccount $AdminAccount`"")
        Selected-ConnectToExchangeOnlineOption -AffectedUser $AffectedUsers -MigrationType $MigrationType -TheAdminAccount $EXOAdminAccount
    }
    ### If ConnectToExchangeOnPremises parameter of the script was used, we will continue on this path.
    elseif ($ConnectToExchangeOnPremises) {
        Write-Log ("[INFO] || The script was started with the ConnectToExchangeOnPremises parameter.")
        Selected-ConnectToExchangeOnPremisesOption -TheAffectedUsers $AffectedUsers
    }
    ### If the script was started without any parameters, we will provide a menu in order to continue
    else {
        Show-Menu
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
        [string]
        $FilePath
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
            [string]$PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $NumberOfChecks
        }
    }
    ### If no FilePath was provided, the script will ask to provide the full path of the .xml file
    else{
        [string]$PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks
    }

    ### If PathOfXMLFile variable will match "NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed", we will continue the data collection
    ### using other methods.
    if ($PathOfXMLFile -match "NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed") {
        [int]$TheNumberOfChecks = 1
    
        ### TheAffectedUser variable will represent the Affected user for which we will try to collect mailbox migration related logs
        Write-Log "[INFO] || Trying to collect the AffectedUser..."
        [string]$TheAffectedUser = Ask-ForDetailsAboutUser -NumberOfChecks $TheNumberOfChecks
        
        ### TheMigrationType variable will represent the Migration type for which the logs have to be investigated
        Write-Log "[INFO] || Trying to collect the Migration Type..."
        [string]$TheMigrationType = Ask-DetailsAboutMigrationType -NumberOfChecks $TheNumberOfChecks -AffectedUser $TheAffectedUser

        ### TheMigrationLogs variable will represent MigrationLogs collected using the Selected-ConnectToExchangeOnlineOption function.
        Write-Log "[INFO] || Trying to collect the Migration Logs using Selected-FileOption -> Selected-ConnectToExchangeOnlineOption function..."
        $script:TheMigrationLogs = Selected-ConnectToExchangeOnlineOption -AffectedUser $TheAffectedUser -MigrationType $TheMigrationType
    }
    else {
        ### TheMigrationLogs variable will represent MigrationLogs collected using the Collect-MigrationLogs function.
        Write-Log "[INFO] || Trying to collect the Migration Logs using Selected-FileOption -> Collect-MigrationLogs function..."
        Collect-MigrationLogs -XMLFile $PathOfXMLFile
    }

}


### <summary>
### Selected-ConnectToExchangeOnlineOption function is used to connect to Exchange Online, and collect from there the mailbox migration logs,
### for the affected user, by running the correct commands, based on the migration type
### </summary>
### <param name="AffectedUser">AffectedUser represents the affected user for which we collect the mailbox migration logs </param>
### <param name="MigrationType">MigrationType represents the migration type for which we collect the mailbox migration logs </param>
### <param name="TheAdminAccount">TheAdminAccount represents username of an Admin that we will use in order to connect to Exchange Online </param>
function Selected-ConnectToExchangeOnlineOption {
    [CmdletBinding()]
    Param (
        [string]
        $AffectedUser,        
        [string]
        $MigrationType,
        [string]
        $TheAdminAccount
    )

    ### We will try to connect to Exchange Online
    Test-EXOSession -TheAdminAccount $TheAdminAccount

}


### <summary>
### Selected-ConnectToExchangeOnPremisesOption function is used to collect mailbox migration logs from Exchange On-Premises (MoveHistory from Get-MailboxStatistics)
### If this option will be selected, the script have to be started from the On-Premises Exchange Management Shell
### </summary>
### <param name="AffectedUser">AffectedUser represents the affected user for which we collect the mailbox migration logs </param>
function Selected-ConnectToExchangeOnPremisesOption {
    [CmdletBinding()]
    Param (
        [string[]]
        $TheAffectedUsers
    )

    $script:OnPremisesPSSession = ConnectTo-ExchangeOnPremises -Prefix $script:ExOnPremCommandsPrefix -ExchangeURL $ExchangeURL -AuthenticationType $AuthenticationType -Domain $Domain -ADSite $ADSite -NumberOfChecks 1
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

    ### Validating if the path has a length greater than 4, and if it is a of an .xml file
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
            Write-Log ("[WARNING] || $XMLFilePath is not a valid .xml file. We will set: XMLFilePath = `"NotAValidXMLFile`"") -ForegroundColor Yellow
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
        [int]
        $NumberOfChecks
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
                Write-Host "We will continue to collect the migration logs using other methods" -ForegroundColor Red
                $PathOfXMLFile = "ValidationOfFileFailed"
            }
        }
        else {
            ### If NO was selected, the script will set PathOfXMLFile to ValidationOfFileFailed, which will be used to collect the logs using other methods
            Write-Host
            Write-Host "We will continue to collect the migration logs using other methods" -ForegroundColor Yellow
            $PathOfXMLFile = "ValidationOfFileFailed"
        }
    }
    
    ### The function returns the full path of the .xml file, or ValidationOfFileFailed
    return $PathOfXMLFile
}


### <summary>
### Ask-ForDetailsAboutUser function is used to collect the Affected user.
### </summary>
### <param name="NumberOfChecks">NumberOfChecks is used in order to provide different messages when collecting the affected user
### for the first time, or if you are re-asking for the affected user </param>
function Ask-ForDetailsAboutUser {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [int]
        $NumberOfChecks
    )    

    Write-Host
    if ($NumberOfChecks -eq "1") {
        ### Asking for the affected user, for the first time
        Write-Log "[INFO] || Asking to provide the affected user, for the first time." -NonInteractive $true
        Write-Host "Please provide the username of the affected user (Eg.: " -NoNewline -ForegroundColor Cyan
        Write-Host "User1@contoso.com" -NoNewline -ForegroundColor White
        Write-Host "): " -NoNewline -ForegroundColor Cyan
        $TheUserName = Read-Host
        $NumberOfChecks++
        Write-Log ("[INFO] || The affected user provided is: $TheUserName") -NonInteractive $true
    }
    else {
        ### Re-asking for the affected user
        Write-Log "[INFO] || Re-asking to provide the affected user." -NonInteractive $true
        Write-Host "Please provide again the username of the affected user (Eg.: " -NoNewline -ForegroundColor Cyan
        Write-Host "User1@contoso.com" -NoNewline -ForegroundColor White
        Write-Host "): " -NoNewline -ForegroundColor Cyan
        $TheUserName = Read-Host
        Write-Log ("[INFO] || The affected user provided is: $TheUserName") -NonInteractive $true
    }

    ### Validating if the user provided is the affected user
    Write-Host
    Write-Host "You entered " -NoNewline -ForegroundColor Cyan
    Write-Host "$TheUserName" -NoNewline -ForegroundColor White
    Write-Host " as being the affected user. Is this correct?" -ForegroundColor Cyan
    Write-Host "`t[Y] Yes     [N] No      (default is `"Y`"): " -NoNewline -ForegroundColor White
    $ReadFromKeyboard = Read-Host

    [bool]$TheKey = $true
    Switch ($ReadFromKeyboard) 
    { 
      Y {$TheKey=$true} 
      N {$TheKey=$false} 
      Default {$TheKey=$true} 
    }

    if ($TheKey) {
        ### Received confirmation that the user provided is the affected user.
        Write-Log ("[INFO] || Got confirmation that `"$TheUserName`" is indeed the affected user.") -NonInteractive $true
    }
    else {
        ### The user provided is not the affected user. Asking again for the affected user.
        Write-Log ("[WARNING] || `"$TheUserName`" is not the affected user. Starting over the process of asking for the affected user.") -NonInteractive $true
        [string]$TheUserName = Ask-ForDetailsAboutUser -NumberOfChecks $NumberOfChecks
    }

    ### The function will return the affected user
    return $TheUserName
}

### <summary>
### Ask-DetailsAboutMigrationType function is used to collect the Migration Type used to do the migration.
### </summary>
### <param name="NumberOfChecks">NumberOfChecks is used in order to provide different messages when collecting the migration type
### for the first time, or if you are re-asking for the migration type </param>
### <param name="AffectedUser">AffectedUser represents the affected user for which we collect the migration type </param>
function Ask-DetailsAboutMigrationType {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [int]
        $NumberOfChecks,
        [string]
        $AffectedUser
    )

    Write-Host
    if ($NumberOfChecks -eq "1") {
        ### Asking about the migration type, for the first time
        Write-Log "[INFO] || Asking about the migration type, for the first time" -NonInteractive $true
        Write-Host "Please select the " -NoNewline -ForegroundColor Cyan
        Write-Host "Migration Type" -NoNewline -ForegroundColor White
        Write-Host " used to migrate " -NoNewline -ForegroundColor Cyan
        Write-Host "$AffectedUser" -ForegroundColor White
        Write-Host "`t[1] Hybrid" -ForegroundColor White
        Write-Host "`t[2] IMAP" -ForegroundColor White
        Write-Host "`t[3] Cutover" -ForegroundColor White
        Write-Host "`t[4] Staged" -ForegroundColor White
        Write-Host "Select 1, 2, 3 or 4 (default is `"1`"): " -NoNewline -ForegroundColor Cyan
        $ReadFromKeyboard = Read-Host
        Write-Log ("[INFO] || You selected $ReadFromKeyboard") -NonInteractive $true

        Switch ($ReadFromKeyboard) 
        { 
          1 {$MigrationType="Hybrid"} 
          2 {$MigrationType="IMAP"} 
          3 {$MigrationType="Cutover"}
          4 {$MigrationType="Staged"}
          Default {$MigrationType="Hybrid"} 
        }

        $NumberOfChecks++
    }
    else {
        ### Re-asking about the migration type
        Write-Log ("[INFO] || Re-asking about the migration type ($NumberOfChecks)") -NonInteractive $true
        Write-Host "Please select again the " -NoNewline -ForegroundColor Cyan
        Write-Host "Migration Type" -NoNewline -ForegroundColor White
        Write-Host " used to migrate " -NoNewline -ForegroundColor Cyan
        Write-Host "$AffectedUser" -ForegroundColor White
        Write-Host "`t[1] Hybrid" -ForegroundColor White
        Write-Host "`t[2] IMAP" -ForegroundColor White
        Write-Host "`t[3] Cutover" -ForegroundColor White
        Write-Host "`t[4] Staged" -ForegroundColor White
        Write-Host "Type 1, 2, 3 or 4 (default is `"1`"): " -NoNewline -ForegroundColor Cyan
        $ReadFromKeyboard = Read-Host
        Write-Log ("[INFO] || You selected $ReadFromKeyboard") -NonInteractive $true

        Switch ($ReadFromKeyboard) 
        { 
          1 {$MigrationType="Hybrid"} 
          2 {$MigrationType="IMAP"} 
          3 {$MigrationType="Cutover"}
          4 {$MigrationType="Staged"}
          Default {$MigrationType="Hybrid"} 
        }
    }

    Write-Host
    Write-Host "You entered " -NoNewline -ForegroundColor Cyan
    Write-Host "$MigrationType" -NoNewline -ForegroundColor White
    Write-Host ". Is this correct?" -ForegroundColor Cyan
    Write-Host "`t[Y] Yes     [N] No      (default is `"Y`"): " -NoNewline -ForegroundColor White
    $ReadFromKeyboard = Read-Host
    Write-Log "[INFO] || You selected the following: `"Migration Type: $MigrationType`"; `"Is this correct? $ReadFromKeyboard`"" -NonInteractive $true

    [bool]$TheKey = $true
    Switch ($ReadFromKeyboard) 
    { 
      Y {$TheKey=$true} 
      N {$TheKey=$false} 
      Default {$TheKey=$true} 
    }

    if ($TheKey) {
        Write-Log ("[INFO] || The script will continue to investigate the details about the $MigrationType migration, for $AffectedUser user") -NonInteractive $true
    }
    else {
        Write-Log ("[INFO] || The script will try to collect again the Migration type") -NonInteractive $true
        [string]$MigrationType = Ask-DetailsAboutMigrationType -NumberOfChecks $NumberOfChecks -AffectedUser $AffectedUser
    }

    ### The function returns the migration type
    return $MigrationType
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
        [string]
        $XMLFile,
        [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $true)]
        [switch]$ConnectToExchangeOnline,
    
        [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $true)]
        [switch]$ConnectToExchangeOnPremises,
    
        [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
        [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
        [string[]]$AffectedUsers,
        
        [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
        [ValidateSet("Hybrid", "IMAP", "Cutover", "Staged")]
        [string]$MigrationType,
    
        [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
        [string]$AdminAccount
    )
    
    if ($XMLFile) {
        ### Importing data in the LogsToAnalyze (Scope: Script) variable
        Write-Log ("[INFO] || Importing data from `"$XMLFile`" file, in the LogsToAnalyze variable")
        $TheMigrationLogs = Import-Clixml $XMLFile
        foreach ($Log in $TheMigrationLogs) {
            $void = $script:LogsToAnalyze.Add($Log)
        }
    }
    elseif ($ConnectToExchangeOnline) {
        ### Connecting to Exchange Online in order to collect the needed/correct mailbox migration logs
        Write-Host "This part is not yet implemented" -ForegroundColor Red
    }
    elseif ($ConnectToExchangeOnPremises) {
        ### Connecting to Exchange On-Premises in order to collect the outputs of relevant MoveHistory from Get-MailboxStatistics
        Write-Log ("[INFO] || Collecting MoveHistory from Get-MailboxStatistics for each Affected users")
        Collect-MailboxStatistics -AffectedUsers $AffectedUsers -TheEnvironment 'Exchange OnPremises'
    }
}


function Collect-MailboxStatistics {
    param (
        [string[]]
        $AffectedUsers,
        [ValidateSet("Exchange Online", "Exchange OnPremises")]
        [string]
        $TheEnvironment
    )

    if ($TheEnvironment -eq "Exchange Online") {
        [string]$TheCommand = "(Get-"+ $script:EXOCommandsPrefix + "MailboxStatistics `$User -IncludeMoveReport -IncludeMoveHistory -ErrorAction Stop).MoveHistory | where {[string]`$(`$_.WorkloadType.Value) -eq `"Onboarding`"} | select -First 1"
    }
    else {
        [string]$TheCommand = "(Get-"+ $script:ExOnPremCommandsPrefix + "MailboxStatistics `$User -IncludeMoveReport -IncludeMoveHistory -ErrorAction Stop).MoveHistory | where {[string]`$(`$_.WorkloadType) -eq `"Offboarding`"} | select -First 1"
    }

    foreach ($User in $AffectedUsers) {
        try {
            Write-Log ("[INFO] || Running the following command:`n`t$TheCommand")
            $MailboxStatistics = Invoke-Expression $TheCommand
            Write-Log "[INFO] || MoveHistory successfully collected for `"$User`" user."
            $void = $script:LogsToAnalyze.Add($MailboxStatistics)
        }
        catch {
            Write-Log "[ERROR] || We were unable to collect MoveHistory from MailboxStatistics for `"$User`" user."
        }
    }
}



#endregion Functions, Global variables

###############
# Main script #
###############
#region Main script

try {
    Clear-Host

    $null = Show-Header

    $script:TheWorkingDirectory = Create-WorkingDirectory -NumberOfChecks 1
    Create-LogFile -WorkingDirectory $script:TheWorkingDirectory

    Check-Parameters

    #region ForTestPurposes - This will be removed

    Write-Host
    Write-Host "Details from the mailbox migration log:" -ForegroundColor Green
    foreach ($Entry in $script:LogsToAnalyze) {
        Write-Host "`tName: " -ForegroundColor Cyan -NoNewline
        Write-Host "$($Entry.MailboxIdentity.Name)" -ForegroundColor White
        Write-Host "`tStatus: " -ForegroundColor Cyan -NoNewline
        Write-Host "$([string]$Entry.Status)" -ForegroundColor White
        Write-Host "`tStatusDetails: " -ForegroundColor Cyan -NoNewline
        Write-Host "$([string]$Entry.StatusDetail)" -ForegroundColor White
        Write-Host "`tExchangeGuid: " -ForegroundColor Cyan -NoNewline
        Write-Host "$([string]$Entry.ExchangeGuid)" -ForegroundColor White
        Write-Host
    }
    #endregion ForTestPurposes

    #Show-Menu
} catch {
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