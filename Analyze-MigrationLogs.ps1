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
    [string]$OnPremAdminAccount
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
### ExOnPremCommandsPrefix (Scope: Script) variable will be used to create a new PSSession to Exchange OnPremises.
### When importing the PSSession, the script will use "MAExOnP" (Migration Analyzer Exchange OnPremises) as Prefix for each command
[string]$script:ExOnPremCommandsPrefix = "MAExOnP"


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
### ConnectTo-ExchangeOnline function is used to connect to Exchange Online
### </summary>
### <param name="TheAdminAccount">TheAdminAccount represents username of an Admin that we will use in order to connect to Exchange Online </param>
### <param name="NumberOfChecks">NumberOfChecks represents the number of times we try to connect to EXO, after which we will fail the script </param>
Function ConnectTo-ExchangeOnline {
    [CmdletBinding()]
    param (
        [string]
        $TheAdminAccount,
        [int]
        $NumberOfChecks
    )
    
    ### If we tried, without success, to connect to Exchange Online for more than 3 times we will fail the script
    if ($NumberOfChecks -gt 3) {
        throw "We were unable to connect to Exchange Online, after we tried for 3 times"
    }

    ### If we do not have the EXOCredential (Scopr: Script) set, we are asking for the EXO Admin's credentials
    ### The credentials will be dismissed just when the script will exit. During the time the script is still running, we will use the credentials in case
    ### we have to reconnect to Exchange Online
    $i = 0
    while ((-not ($script:EXOCredential)) -and ($i -lt 5)){
        $script:EXOCredential = Get-Credential $TheAdminAccount -Message "Please provide your Exchange Online Credentials:"
    }

    ### If we still don't have a credential object then abort
    if (-not ($script:EXOCredential)) {
        throw "Failed to get credentials for connecting to Exchange Online"
    }

    ### Destroy any outstanding EXO related PSSessions
    Write-Log "[INFO] || Removing all Exchange Online related PSSessions"
    $null = Get-PSSession | where {$_.ComputerName -like "outlook.office365.com"} | Remove-PSSession -Confirm:$false
    
    ### Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
    [System.GC]::Collect()
    
    ### Sleep 5s to allow the sessions to tear down fully
    Write-Log "[INFO] || Sleeping 5 seconds for Session Tear Down"
    Start-SleepWithProgress -SleepTime 5

    ### Create the EXO session
    Write-Log "[INFO] || Creating new Exchange Online PSSession"
    
    try {
        $script:EXOsession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $EXOCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
        Write-Log "[INFO] || We managed to successfully create the Exchange Online PSSession"
    }
    catch {
        ### If error, retry
        Write-Log "[ERROR] || We were unable to establish a connection to Exchange Online" -ForegroundColor Red
        Write-log ("[ERROR] || $Error") -ForegroundColor Red
        $script:EXOCredential = $null
        $NumberOfChecks++
        ConnectTo-ExchangeOnline -NumberOfChecks $NumberOfChecks
    }

    ### Import the PSSession
    try {
        $null = Import-PSSession $EXOsession -AllowClobber -Prefix $script:EXOCommandsPrefix
        Write-Log "[INFO] || We managed to successfully import the Exchange Online PSSession"
    }
    catch {
        ### If error, retry
        Write-Log "[ERROR] || We were unable to import the Exchange Online PSSession" -ForegroundColor Red
        Write-log ("[ERROR] || $Error") -ForegroundColor Red
        $script:EXOCredential = $null
        $NumberOfChecks++
        ConnectTo-ExchangeOnline -NumberOfChecks $NumberOfChecks
    }
}


### <summary>
### Start-SleepWithProgress function is used to sleep X seconds and display a progress bar
### </summary>
### <param name="sleeptime">sleeptime represents the number of seconds for which the script will sleep </param>
Function Start-SleepWithProgress {
    Param(
        [int]
        $sleeptime
    )

	### Loop Number of seconds you want to sleep
	For ($i=0;$i -le $sleeptime;$i++){
		$timeleft = ($sleeptime - $i);
		
		### Progress bar showing progress of the sleep
		Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		
		### Sleep 1 second
		start-sleep 1
	}
	
	Write-Progress -Completed -Activity "Sleeping"
}

### <summary>
### Test-EXOSession function is used to test if we have an Open PSSession to Exchange Online
### </summary>
### <param name="TheAdminAccount">TheAdminAccount represents username of an Admin that we will use in order to connect to Exchange Online </param>
Function Test-EXOSession {
    [CmdletBinding()]
    param (
        [string]
        $TheAdminAccount
    )

	### Reset and regather our session information
	$SessionInfo = $null
	$SessionInfo = Get-PSSession | where {$_.ComputerName -like "outlook.office365.com"}
	
	### Make sure we found a session
	if (-not ($SessionInfo)) { 
		Write-Log "[ERROR] || No Exchange Online related PSSession was found" -ForegroundColor Red
		Write-log "[INFO] || Recreating the session"
		ConnectTo-ExchangeOnline -TheAdminAccount $TheAdminAccount -NumberOfChecks 1
	}	
	### Make sure it is in an opened state if not log and recreate
	elseif ($SessionInfo.State -ne "Opened") {
		Write-Log "[ERROR] || The Exchange Online related PSSession is not in Open state" -ForegroundColor Red
		Write-log ($SessionInfo | fl | Out-String ) -ForegroundColor Red
		Write-log "[INFO] || Recreating the Exchange Online Session"
		ConnectTo-ExchangeOnline -TheAdminAccount $TheAdminAccount -NumberOfChecks 1
    }
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
        [string]$TheCommand = "(Get-MailboxStatistics `$User -IncludeMoveReport -IncludeMoveHistory -ErrorAction Stop).MoveHistory | where {[string]`$(`$_.WorkloadType) -eq `"Offboarding`"} | select -First 1"
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

    $script:OnPremisesPSSession = ConnectTo-ExchangeOnPremises -Prefix $script:ExOnPremCommandsPrefix
}

### <summary>
### ConnectTo-ExchangeOnPremises function is used to collect mailbox migration logs from Exchange On-Premises (MoveHistory from Get-MailboxStatistics)
### If this option will be selected, the script have to be started from the On-Premises Exchange Management Shell
### </summary>
### <param name="AffectedUser">AffectedUser represents the affected user for which we collect the mailbox migration logs </param>
function ConnectTo-ExchangeOnPremises
{
    <#
    .SYNOPSIS
        Connect to any local or remote exchange server. The source computer does not need to be domain joined.
    .DESCRIPTION
        Connect to any or multiple exchange servers in a local or remote domain or connect to a specific exchange server in a AD site. 
        The source computer does not need to be domain joined.when no switches used autoresolve will try to resolve exchange servers 
        in the local client site , adjacent sites or finaly any sites.
        
        This function is dependant on the function "Get-LdapObject"
    .EXAMPLE
        Connect-Exchangeservice

        Connect to a local domain exchange server of the version 2010 in the same site as your client.
    .EXAMPLE
        Connect-Exchangeservice -Version 2016

        Connect to a local domain exchange server of the version 2016 in the same site as your client.
    .EXAMPLE
        Connect-Exchangeservice -Prefix NY -ADsite "NY" ; Connect-Exchangeservice -version 2013 -Prefix AT -ADsite "AT"

        First connect to a local domain exchange server of the version 2010 in the AD site "NY" and prefix all commands with NY. 
        Seconly connect to a local domain exchange server of the version 2013 in the AD site "AT" and prefix all commands with AT.
    .EXAMPLE
        Connect-Exchangeservice -Domain "Tailspin.com" -Creds (get-credential)

        Connect to a Remote exchange 2010 server in the domain "tailspin.com" using credentials from a credential prompt.
    .NOTES
        Function name : Connect-ExchangeService
        Authors       : Martijn van Geffen
        Version       : 1.3
        
        Dependencies:
        This function is dependant on the function "Get-LdapObject"
        This function can be found on: Http://www.tech-savvy.nl
        
        Version Changes: 
        03-11-2016 V0.1 : Initial Script (MvG) 
        11-01-2017 V0.2 : Servers variable declared as type "array" due to environments with 1 server (MvG)
        17-01-2017 V0.2 : Updated function with the option to overwrite the exchange connection URL
                          Updated the function to have AD site support for non domain jioned pc`s
                          Updated the function to support non domain joined computers
                          Updated the functions paramater sets
        25-01-2017 V1.0 : Released to TechNet
        01-12-2017 V1.1 : Update code to retry any ADsite when a specific version is used without adsite 
                          and client site does not contain a exchangeserver of that version 
        01-04-2018 V1.2 : Redone the autoresolve logic 
                          Added support for searching adjacent AD sites if no server in own site
                          Updated code to skip Edge servers
                          Added some additional verbose code for trouble shooting
                          Added some additional debug code for trouble shooting
        02-20-2018 V1.3 : Changed behavior to always search adjacent sites if no servers found.
    #>

    [CmdletBinding(DefaultParameterSetName='resolve',
                   HelpUri = 'https://gallery.technet.microsoft.com/Connect-to-one-or-multiple-b850411d'
    )]

    Param(

    [Parameter(Mandatory=$false)]
    [validateScript({If ($_ -eq "notdeclared" -or $_ -eq "2007" -or $_ -eq "2010" -or $_ -eq "2013" -or $_ -eq "2016") 
                        {
                            $true
                        }
                        else
                        {
                            throw "$_ is not a valid version of exchange use 2007, 2010, 2013 or 2016"    
                        }
                    })]
    [string]$Version = "notdeclared",

    [Parameter(Mandatory=$false,
        ParameterSetName='resolve')]
    [ValidateNotNullOrEmpty()]
    [string]$ADSite = "getsite",

    [Parameter(Mandatory=$false,
        ParameterSetName='resolve')]
    [ValidateNotNullOrEmpty()]
    [string]$Domain = "getdomain",

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [System.Management.Automation.PsCredential]$Creds,

    [Parameter(Mandatory=$true,
        ParameterSetName='manual')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({$_ -match "[htp]{4}"})] 
    [string]$exchangeurl,

    [Parameter(Mandatory=$true,
        ParameterSetName='manual')]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Basic","Digest","negotiate","kerberos")]
    [string]$authenticationtype,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [String]$Prefix

    )

    #region Determine the exchange version

    switch ($version)
    {
        "notdeclared" {$versionnr = "notdeclared"}
        "2007"        {$versionnr = " 8 "}
        "2010"        {$versionnr = " 14."}
        "2013"        {$versionnr = " 15.0 "}
        "2016"        {$versionnr = " 15.1 "}
    }

    Write-Log ("[INFO] || Function is set to use exchange version: $versionnr")

    #endregion Determine the exchange version

    #region Determine if a url is being specified

    if (([string]::IsNullOrEmpty($exchangeurl))) {
        Write-Log "[INFO] || Function is set to use exchange URL: Auto-resolve"
        [boolean]$autoresolveurl = $true
    }
    else {
        Write-Log ("[INFO] || Function is set to use exchange URL: $exchangeurl")
        [boolean]$autoresolveurl = $false
    }

    #endregion Determine if a url is being specified

    
    if ($autoresolveurl) {
        #region Start auto resolve connection url
        #region Determine AD domain 

        if ($Domain -eq "getdomain") {
            try {
                Write-Log "[INFO] || Function will now query AD domain using .net"
                $Domain = ([system.directoryservices.activedirectory.domain]::GetCurrentDomain()).name
            }
            catch {
                Write-Log ("$($_.exception.message)")
                throw "[ERROR] || Function Could not resolve Active directory domain and -Domain switch is not used."
            }
        }

        #endregion Determine AD domain
        #region Determine AD site 

        if ($ADsite -eq "getsite") {
            $sitemanualset = $false
            try {
                Write-Log "[INFO] || Function will now query current computer AD site using .net"
                $ADsitename = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).name
            }
            catch {
                Write-Log ("$($_.exception.message)")
                Write-Log "[WARNING] || Could not resolve the client AD site. The Function will continue autoresolve using any AD site as filter. You might end up on a slow AD site. Please validate your IP to AD site binding settings" -ForegroundColor Yellow
                $ADsitename = "*"                
            }
        }
        else {
            $ADsitename = $ADsite
            $sitemanualset = $true
        }

        Write-Log  ("[INFO] || Function is set to use ADsite: $ADsitename")

        #endregion Determine AD site
        #region Build exchange search filter
        #region craft exchange site filter

        $FilterADSite = "(&(objectclass=site)(Name=$ADsitename))"
        $ADsiteobject = Get-Ldapobject -LDAPfilter $FilterADSite -configurationNamingContext -configurationNamingContextdomain $domain
        $ADsiteobjectdn = $ADsiteobject.properties.distinguishedname
        
        if ([string]::IsNullOrEmpty($ADsiteobjectdn)) {
            Write-Log ("[ERROR] || Failing AD query: $FilterADSite") -ForegroundColor Red
            throw "[ERROR] || Could not find the AD site. Please check you spelling if you used -ADsite parameter. If autoresolve is used a connectivity error occured"
        }

        #endregion craft exchange site filter

        if ($versionnr -eq "notdeclared"){
            $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*))"
        }
        else {
            $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*$versionnr*))"
        }
        
        #endregion Build exchange search filter
        #region Harvest exchange servers
        [array]$Servers =@()
        $tempallServers = Get-Ldapobject -LDAPfilter $Filterexservers -configurationNamingContext -configurationNamingContextdomain $domain -Findall $true
        [array]$Servers += $tempallServers
        
        if ($Servers.count -eq 0) {       
            Write-Log ("[WARNING] || Function did resolve 0 servers in the ADsite $adsit using filter: $Filterexservers ") -ForegroundColor Yellow
            Write-Log "[INFO] || Retrying with next closest AD sites"
            $AdjacentSites = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).AdjacentSites.name
            Write-Log ("[INFO] || Found ajecentsites AD sites: $($AdjacentSites -join ", ") ")

            foreach ($site in $AdjacentSites) {
                Write-Log "[INFO] || Function trying to find servers in AD sites: $site "

                #clean iterative variables
                              
                $FilterADSite = $null
                $ADsiteobject = $null
                $ADsiteobjectdn = $null
                $Filterexservers = $null
                $tempallServers = $null
                $selectedsiteServers = $null

                #region retry Build exchange search filter
                
                #region craft exchange site filter retry

                $FilterADSite = "(&(objectclass=site)(Name=$site))"
                $ADsiteobject = Get-Ldapobject -LDAPfilter $FilterADSite -configurationNamingContext -configurationNamingContextdomain $domain
                $ADsiteobjectdn = ($ADsiteobject).properties.distinguishedname
        
                if ([string]::IsNullOrEmpty($ADsiteobjectdn)) {
                    write-verbose -Message "failing AD query: $FilterADSite"
                    throw "Error - Could not find the AD site. Please check you spelling if you used -ADsite parameter. If autoresolve is used a connectivity error occured"
                    break    
                }

                #endregion craft exchange site filter retry

                if ($versionnr -eq "notdeclared") {
                    $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*))"
                }
                else {
                    $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*$versionnr*))"
                }
                
                #endregion retry Build exchange search filter

                #region Retry Harvest exchange servers
                
                $tempallServers = Get-Ldapobject -LDAPfilter $Filterexservers -configurationNamingContext -configurationNamingContextdomain $domain -Findall $true
                $selectedsiteServers = $tempallServers.properties.name -join ", "
                if ($tempallServers.count -ge 1) {
                    Write-Log ("[INFO] || Function found new servers in site $site. Adding servers: $($selectedsiteServers -join ", ")")
                    $Servers += $tempallServers
                }
                #endregion Retry Harvest exchange servers
            }

            #region Final attempt Harvest exchange servers

            if ($Servers.count -eq 0) {
                Write-Log ("[INFO] || Function did resolve 0 servers in adjacent sites to ADsite $adsitename using filter: $Filterexservers ")
                Write-Log "Function last attempt: Any site , any version"
                $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer))"
                [array]$Servers +=  Get-Ldapobject -LDAPfilter $Filterexservers -configurationNamingContext -configurationNamingContextdomain $domain -Findall $true
                if ($Servers.count -eq 0) {
                    throw "Function was unable to identify any Exchange servers in your organization"
                }                    
            }

            #endregion Final attempt Harvest exchange servers
        }

        Write-Log ("[INFO] || Function found the following exchange servers to try to connect to: $($Servers.properties.name -join ", ")")
    }

         

    do {
        try {
            if (!($exchangeurl)) { 
                if (!([string]::IsNullOrWhiteSpace($Servers))) {
                    Write-Verbose -Message "The following servers have been found $($servers.properties.name)" 
                    $server = get-random $servers
                }
                else {
                    write-output -InputObject "There are 0 exchange servers of version $version $tempversion in the site $adsite"
                    throw
                }
                $ip = ($server.properties.networkaddress | ?{$_ -like "ncacn_ip_tcp*" }).split(":")[1]
                $serverconnection = "http://$ip/powershell"

            }else {
                $serverconnection = $exchangeurl
            }

            if ([string]::IsNullOrWhiteSpace($creds.UserName)) {
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -ErrorAction STOP
            }
            elseif ([string]::IsNullOrWhiteSpace($authenticationtype)) {
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -Credential $creds -ErrorAction STOP
            }
            else {
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -Credential $creds -Authentication $authenticationtype -ErrorAction STOP
            }

            if ([string]::IsNullOrWhiteSpace($prefix)) {
                Import-PSSession $Session 4>&1 3>&1 | out-null
            }
            else {
                Import-PSSession $Session -prefix $prefix 4>&1 3>&1 | out-null
            }

            write-Log ("[INFO] || Connected and imported session from server: $serverconnection")

        }
        catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
            write-output (" tried connecting to $serverconnection but could not connect ")    
            if ($_.exception.message -like "*Access is denied*" -or $_.exception.message -like "*Access denied*")
            {
                Write-Log "[ERROR] || Access is denied. invalid or no credentials. Provide credentials" -ForegroundColor Red
            }

            $connectionerrorcount ++
            if ( $connectionerrorcount -ge 2 )
            {
                $session = "failed"
                Write-Log "$($_.exception.message)"
                Write-Log ("[ERROR] || Tried connecting 2 times but could not connect due to invalid credentials. last exchange server we tried : $serverconnection") -ForegroundColor Red
            }
        }
        catch {
            $connectionerrorcount ++
            if ( $_.exception.Message -like "No command proxies have been created*") {
                Write-Log "$($_.exception.message)" -ForegroundColor Red
                Write-Log "[INFO] || No command proxies have been created, because all of the requested remote commands would shadow existing local commands."
                Write-Log ("[INFO] || Connected and imported new commands from server: $serverconnection")

            }
            elseif ( $_.exception.Message -like "*The attribute cannot be added because variable*") {
                #Catch Powershell Bug object validation: http://stackoverflow.com/questions/19775779/powershell-getnewclosure-and-cmdlets-with-validation
                Write-Log ("[ERROR] || Powershell Bug object validation plz ignore bug raport is created at microsoft: $($_.exception.message)")
            }
            else {
                Write-Log ("$($_.exception.message)") -ForegroundColor Red
                Write-Log ("[ERROR] || Tried connecting to $serverconnection but could not connect " + $_.exception.Message)  -ForegroundColor Red
            }
            if ( $connectionerrorcount -ge 5 ) {
                $session = "failed"
                Write-Log ("$($_.exception.message)") -ForegroundColor Red
                Write-Log ("[INFO] || tried connecting 5 times but could not connect. last exchange server we tried : $serverconnection")     
            }
        }
        finally
        {
            if ( $ADsite -is [System.IDisposable]){ 
                $ADsite.Dispose()
            }
            if ( $domain -is [System.IDisposable]) { 
                $domain.Dispose()
            }
        }
    }
    until ($session)
} 


function Get-Ldapobject
{
    <#
    .SYNOPSIS
        Search LDAP directorys using .NET LDAP searcher. The function supports query`s from any pc no matter if it is joined to the domain.
        The function has support for all  partition types and multi domain / forest setups.
    .DESCRIPTION
        Search AD configuration or naming partition or using .NET AD searcher 
    .EXAMPLE
        Get-Ldapobject -LDAPfilter "(&(name=henk*)(diplayname=*))"

        Search the current domain with the LDAP filter "(&(name=Henk*)(diplayname=*))". Return all properties.
        Return only 1 result
    .EXAMPLE
        Get-Ldapobject -LDAPfilter "(&(name=henk*)(diplayname=*))" -properties Displayname,samaccountname -Findall $true

        Search the current domain with the LDAP filter "(&(name=henk*)(diplayname=*))". Return Displayname and samaccountname.
        Return all result 
    .EXAMPLE
        Get-Ldapobject -OU "OU=users,DC=contoso,DC=com" -DC "DC01" -LDAPfilter "(&(name=henk*)(diplayname=*))" -properties samaccountname

        Search the OU "users" in the domain "contoso.com" using DC01 and the LDAP filter "(&(name=henk*)(diplayname=*))". Return the
        samaccountname. Return only 1 result
    .EXAMPLE
        Get-Ldapobject -OU "CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=tailspin,DC=com" -LDAPfilter 
        "(&(objectclass=msExchExchangeServer)(serialnumber=*15*))" -Findall $true -$configurationNamingContext

        Search the current AD domain for all exchange 2013 and 2016 servers in the configuration partition of AD.
        Return all result 
    .EXAMPLE
        Get-Ldapobject -OU "CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=tailspin,DC=com" -LDAPfilter 
        "(objectclass=msExchExchangeServer)" -Findall $true -ConfigurationNamingContext -ConfigurationNamingContextdomain "tailspin.com"

        Search the Remote AD domain "tailspin.com" for all exchange servers in the configuration partition of AD.
        Return all result
    .NOTES
        -----------------------------------------------------------------------------------------------------------------------------------
        Function name : Get-Ldapobject
        Authors       : Martijn van Geffen
        Version       : 1.2
        dependancies  : None
        -----------------------------------------------------------------------------------------------------------------------------------
        -----------------------------------------------------------------------------------------------------------------------------------
        Version Changes:
        Date: (dd-MM-YYYY)    Version:     Changed By:           Info:
        12-12-2016            V1.0         Martijn van Geffen    Initial Script.
        06-01-2017            V1.1         Martijn van Geffen    Released on Technet
        26-02-2018            V1.2         Martijn van Geffen    Set the default OU to the forest root to better support multi domain
                                                                 and multi forest
        -----------------------------------------------------------------------------------------------------------------------------------
    .COMPONENT
        None
    .ROLE
        None
    .FUNCTIONALITY
        Search LDAP directorys using .NET LDAP searcher
    #>

    [CmdletBinding(HelpUri='https://gallery.technet.microsoft.com/scriptcenter/Search-AD-LDAP-from-domain-c0131588')]
    [Alias("glo")]
    [OutputType([System.Array])]

    param(

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$OU,
    
    [Parameter(Mandatory=$false)]
    [string]$DC,
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$LDAPfilter,

    [Parameter(Mandatory=$false)]
    [array]$Properties = "*",

    [Parameter(Mandatory=$false)]
    [boolean]$Findall = $false,
        
    [Parameter(Mandatory=$false)]
    [string]$Searchscope = "Subtree",

    [Parameter(Mandatory=$false)]
    [int32]$PageSize = "900",

    [Parameter(Mandatory=$false)]
    [switch]$ConfigurationNamingContext,

    [Parameter(Mandatory=$false)]
    [string]$ConfigurationNamingContextdomain,

    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PsCredential]$Cred   

    )
    
    If ( $cred )
    {
        $username = $Cred.username
        $password = $Cred.GetNetworkCredential().password
    }

    if ( !$DC )
    {
        try 
        {
            $DC = ([system.directoryservices.activedirectory.domain]::GetCurrentDomain()).name
            write-verbose -message "Current "
        }
        catch
        {
            Write-error "Variable DC can not be empty if you run this from a non domain joined computer. Use a DC or Use Get-dc function here from https://gallery.technet.microsoft.com/scriptcenter/Find-a-working-domain-fe731b4f"
        }
    }

    if ( !$OU )
    {
        try 
        {
            $OU = "DC=" + ([string]([system.directoryservices.activedirectory.domain]::GetCurrentDomain()).forest).Replace(".",",DC=")
        }
        catch
        {
            Write-error "Variable OU can not be empty if you run this from a non domain joined computer. Use a DC or Use Get-dc function here from https://gallery.technet.microsoft.com/scriptcenter/Find-a-working-domain-fe731b4f"
        }
    }

    Try
    {
        if ( $cred )
        {
            $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DC/$OU",$username,$password)
        }else
        {
            $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DC/$OU")
        } 
        
        if ( $configurationNamingContext.IsPresent )
        {
        
            try
            {
                if (!$ConfigurationNamingContextdomain)
                {
                    $ConfigurationNamingContextdomain = [system.directoryservices.activedirectory.domain]::GetCurrentDomain()
                }
                $tempconfigurationNamingContextdomain = $configurationNamingContextdomain
            }
            catch
            {
                Write-error "Variable ConfigurationNamingContextdomain can not be empty if you run this from a not domain joined computer"
            }

            try
            {
                do
                {
                    if ( $cred )
                    {
                        $tempdomain = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$tempconfigurationNamingContextdomain,$username,$password)
                    }else
                    {
                        $tempdomain = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$tempconfigurationNamingContextdomain)
                    }
                    $domain = [system.directoryservices.activedirectory.domain]::GetDomain($tempdomain)
                    $configurationNamingContextdomain = $domain.forest.name
                    $tempconfigurationNamingContextdomain = $domain.parent
                }while ( $domain.parent )

                $configurationdn = "CN=configuration,DC=" + $configurationNamingContextdomain.Replace(".",",DC=")
                if ( $cred )
                {
                    $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DC/$configurationdn",$username,$password)
                }else
                {
                    $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DC/$configurationdn")
                }
                      
            }
            Finally
            {
                if (  $domain -is [System.IDisposable])
                { 
                     $domain.Dispose()
                }
                if ( $configurationNamingContextdomain -is [System.IDisposable])
                { 
                     $configurationNamingContextdomain.Dispose()
                }
            }
        
        }
                   
        $searcher = new-object DirectoryServices.DirectorySearcher($root)
        $searcher.filter = $LDAPfilter
        $searcher.PageSize = $PageSize
        $searcher.searchscope = $searchscope
        $searcher.PropertiesToLoad.addrange($properties)

        if ($findall)
        {
            [System.Array]$object = $searcher.Findall()
        }
    
        if (!$findall)
        {
            [System.Array]$object = $searcher.Findone()
        }

    }
    Finally
    {        
        if ( $searcher -is [System.IDisposable])
        { 
            $searcher.Dispose()
        }
        if ( $OU -is [System.IDisposable])
        { 
            $OU.Dispose()
        }
        if ( $DC -is [System.IDisposable])
        { 
            $DC.Dispose()
        }
    }
    return $object
}

### <summary>
### Selected-ConnectToExchangeOnPremisesOption function is used to collect mailbox migration logs from Exchange On-Premises (MoveHistory from Get-MailboxStatistics)
### If this option will be selected, the script have to be started from the On-Premises Exchange Management Shell
### </summary>
### <param name="AffectedUser">AffectedUser represents the affected user for which we collect the mailbox migration logs </param>
function OldSelected-ConnectToExchangeOnPremisesOption {
    [CmdletBinding()]
    Param (
        [string[]]
        $TheAffectedUsers
    )

    [bool]$ExchangeManagementShell = CheckIf-ExchangeManagementShell -ExchangeOnPremises

    if (-not ($ExchangeManagementShell)) {
        throw "Please start again the script, but, from OnPremises Exchange Management Shell."
    }
    else {
        if (-not ($TheAffectedUsers)) {
            Write-Log "[INFO] || Selected-ConnectToExchangeOnPremisesOption function was called, but, no affected user was provided."
            Write-Log "[INFO] || The script will ask for an affected user."
            [string]$TheAffectedUsers = Ask-ForDetailsAboutUser -NumberOfChecks $TheNumberOfChecks
        }

        [string[]]$TheAffectedUsers = Extract-CorrectListOfUsersForMailboxStatistics -TheAffectedUsers $TheAffectedUsers -TheEnvironment 'Exchange OnPremises'
        Collect-MigrationLogs -ConnectToExchangeOnPremises -AffectedUsers $TheAffectedUsers
    }
}


### <summary>
### Extract-CorrectListOfUsersForMailboxStatistics function is used to check if the list of affected users can be used in order to collect MailboxStatistics output
### </summary>
### <param name="TheAffectedUsers">TheAffectedUsers represent the list of affected users for which we do the check </param>
### <param name="TheEnvironment">TheEnvironment informs us about the environment in which we want to run the Get-MailboxStatistics command </param>
function Extract-CorrectListOfUsersForMailboxStatistics {
    param (
        [string[]]
        $TheAffectedUsers,
        [ValidateSet("Exchange Online", "Exchange OnPremises")]
        [string]$TheEnvironment
    )

    [System.Collections.ArrayList]$UsersOK = @()
    [System.Collections.ArrayList]$UsersNotOK = @()
    ### For each user, checking if they are UserMailbox as RecipientType in the mentioned environment
    foreach ($User in $TheAffectedUsers) {
        Write-Log ("[INFO] || Verifying if the $User user is an UserMailbox in $TheEnvironment")
        try {
            if ($TheEnvironment -eq "Exchange Online") {
                [string]$TheCommand = "Get-"+ $script:EXOCommandsPrefix + "Recipient `$User -ResultSize Unlimited -ErrorAction Stop"
            }
            else {
                [string]$TheCommand = "Get-Recipient `$User -ResultSize Unlimited -ErrorAction Stop"
            }
            $GetRecipient = Invoke-Expression $TheCommand
            Write-Log ("[INFO] || $User user is an UserMailbox in $TheEnvironment")
            Write-Log ("[INFO] || Details about the user:`n`tUserPrincipalName: $($GetUser.UserPrincipalName)`n`tSamAccountName: $($GetUser.SamAccountName)`n`tOrganizationalUnit: $($GetUser.OrganizationalUnit)`n`tDistinguishedName: $($GetUser.DistinguishedName)`n`tGuid: $($GetUser.Guid)`n`tRecipientTypeDetails: $($GetUser.RecipientTypeDetails)") -NonInteractive $true
            Write-Host "Details about the user:" -ForegroundColor Green
            Write-Host "`tPrimarySMTPAddress: " -ForegroundColor Cyan -NoNewline
            Write-Host "$($GetRecipient.PrimarySMTPAddress)" -ForegroundColor White
            Write-Host "`tOrganizationalUnit: " -ForegroundColor Cyan -NoNewline
            Write-Host "$($GetRecipient.OrganizationalUnit)" -ForegroundColor White
            Write-Host "`tDistinguishedName: " -ForegroundColor Cyan -NoNewline
            Write-Host "$($GetRecipient.DistinguishedName)" -ForegroundColor White
            Write-Host "`tGuid: " -ForegroundColor Cyan -NoNewline
            Write-Host "$($GetRecipient.Guid)" -ForegroundColor White
            Write-Host "`tRecipientTypeDetails: " -ForegroundColor Cyan -NoNewline
            Write-Host "$($GetRecipient.RecipientTypeDetails)" -ForegroundColor White
            Write-Host

            Write-Log ("[INFO] || Adding $User user to the UsersOK variable")

            ### If found, added to the UsersOK list
            $void = $UsersOK.Add($User)
        }
        catch {
            ### If not found, added to the UsersNotOK list
            Write-Log ("[WARNING] || Adding $User user to the UsersNotOK variable") -ForegroundColor Yellow
            $void = $UsersNotOK.Add($User)
        }
    }

    ### Throwing error if none of the affected users are UserMailboxes in the environment
    if ($($UsersOK.Count) -eq 0) {
        throw "The users provided do not have UserMailbox as RecipientType.`nIf the affected mailbox is in the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online, and provide the PrimarySMTPAddress of the affected user.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch), and provide the PrimarySMTPAddress of the affected user."
    }
    elseif ($($UsersOK.Count) -gt 0) {
        Write-Log ("[INFO] || List of users for which we will continue to collect the mailbox migration logs, using this method:`n`t$UsersOK")
    }
    
    if ($($UsersNotOK.Count) -gt 0) {
        Write-Log ("[WARNING] || List of users for which we will not continue to collect the mailbox migration logs:`n`t$UsersNotOK") -ForegroundColor Yellow
    }

    ### This function returns the list of users for which we can collect MailboxStatistics
    return $UsersOK
}


### <summary>
### CheckIf-ExchangeManagementShell function is used to check if the script was started from Exchange Management Shell.
### </summary>
### <param name="ExchangeOnPremises">ExchangeOnPremises will be used if the script is started with the ConnectToExchangeOnPremises parameter </param>
function CheckIf-ExchangeManagementShell {
    param (
        [switch]
        $ExchangeOnPremises
    )

    [bool]$ExchangeManagementShell = $false
    ### Checking if the script was started in Exchange Management Shell
    try {
        $null = Get-Command Get-ExBlog -ErrorAction Stop
        Write-Log "[INFO] || The script was started from Exchange Management Shell"
        $ExchangeManagementShell = $true
    }
    catch {
        Write-Log "[INFO] || The script was not started from Exchange Management Shell"
    }
    
    if ($ExchangeManagementShell) {
        if (-not ($ExchangeOnPremises)) {
            ### In case the script was not started with the ConnectToExchangeOnPremises parameter, the script will exit
            throw "You started the script from Exchange Management Shell, even if you do not specifically want to run it in Exchange Management Shell.`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
        }
        else {
            ### In case the script was started with ConnectToExchangeOnPremises parameter, checking if we can use this EMS to continue
            ### Checking how many modules were imported in this EMS using Import-PSSession command
            $TheModules = Get-Module | where {($_.ModuleType -eq "Script") -and ($_.Name -like "tmp*")}
    
            if ($($TheModules.Count) -gt 0) {
                ### If we find more than 1 module
                Write-Log "[INFO] || Found $($TheModules.Count) modules of type `"Script`", for which the Name starts with `"tmp`""
                [System.Collections.ArrayList]$script:EXOModules = @()
                [System.Collections.ArrayList]$script:OnPremModules = @()
                [System.Collections.ArrayList]$script:NotUsefulModules = @()
                foreach ($Module in $TheModules) {
                    ### Checking details for each module (if it's using any prefix, and if it is related to Exchange Online, or Exchange OnPremises)
                    Write-Host
                    Write-Host "Checking the following module: " -ForegroundColor Cyan -NoNewline
                    Write-Host "$($Module.Name)" -ForegroundColor White
                    $Prefix = $Module.Prefix

                    if ($Prefix) {
                        Write-Log ("[INFO] || The prefix used for this module (`"$($Module.Name)`") is: `"$Prefix`"") -NonInteractive $True
                        Write-Host "`tThe prefix used for this module is: " -ForegroundColor Cyan -NoNewline
                        Write-Host "$Prefix" -ForegroundColor White
                    }
                    else {
                        Write-Log ("[INFO] || This module (`"$($Module.Name)`") doesn't have any prefix") -ForegroundColor Green
                    }

                    Write-Log ("[INFO] || Checking if this module (`"$($Module.Name)`") is related to Exchange Online, or Exchange On-Premises") -ForegroundColor Cyan

                    ### Creating an Exchange Online specific command, to check later
                    [string]$EXOCommand = "Get-" + $Prefix + "SyncRequest"
                    ### Creating an Exchange OnPremises specific command, to check later
                    [string]$OnPremCommand = "Get-" + $Prefix + "ExchangeServer"

                    ### Checking if the module is Exchange Online related
                    if ($($Module.ExportedCommands["$EXOCommand"])) {
                        Write-Log ("[INFO] || This module (`"$($Module.Name)`") is related to Exchange Online") -ForegroundColor Green
                        $void = $script:EXOModules.Add($Module)
                    }
                    ### Checking if the module is Exchange OnPremises related
                    elseif ($($Module.ExportedCommands["$OnPremCommand"])) {
                        Write-Log ("[INFO] || This module (`"$($Module.Name)`") is related to Exchange On-Premises") -ForegroundColor Green
                        $void = $script:OnPremModules.Add($Module)
                    }
                    else {
                        Write-Log ("[WARNING] || This module (`"$($Module.Name)`") is not related to Exchange Online, or Exchange OnPremises") -ForegroundColor Green
                        $void = $script:NotUsefulModules.Add($Module)
                    }
                }
            }

            ### Throwing relevant errors
            if ([int]$($script:EXOModules.Count) -eq 1) {
                throw "In this Exchange Management Shell session, the script found also 1 module with Exchange Online related commands.`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
            }
            elseif ([int]$($script:EXOModules.Count) -gt 1) {
                throw "In this Exchange Management Shell session, the script found more than 1 module with Exchange Online related commands (found $($script:EXOModules.Count) modules).`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
            }
            elseif ([int]$($script:OnPremModules.Count) -eq 1) {
                throw "In this Exchange Management Shell session, the script found also 1 module with Exchange OnPremises related commands.`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
            }
            elseif ([int]$($script:OnPremModules.Count) -gt 1) {
                throw "In this Exchange Management Shell session, the script found more than 1 module with Exchange OnPremises related commands (found $($script:OnPremModules.Count) modules).`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
            }
        }
    }
    ### This function returns if the script started in Exchange Management Shell ($true), or no ($false)
    return $ExchangeManagementShell
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