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
A.	Necessary modules:
    1.	Collect the migration logs (related to one or multiple affected mailboxes):
        a.	From an existing .xml file;
        b.	From Exchange Online, by using the correct command:
            i.      For Hybrid;
            ii.     For IMAP;
            iii.    For Cutover / Staged.

        c.	From Get-MailboxStatistics output, if we speak about Remote moves (Hybrid), in case customer already removed the MoveRequest:
            i.	From Exchange Online, if we speak about an Onboarding;
            ii.	From Exchange On-Premises, if we speak about an Offboarding.

    2.	Download the JSON file from GitHub, and, based on the error received, based on the Migration type, we will provide recommendation about the actions that they can take to solve the issue.

B.	Good to have modules:
    1.	Performance analyzer. Similar to what Karahan provided in his script;
    2.	DiagnosticInfo analyzer.
        a.	Using Build-TimeTrackerTable function from Angus’s module, I’ll parse the DiagnosticInfo details, and provide some information to customer.
        b.	Using the idea described here, I’ll create a function that will provide a Column/Bar Chart similar to (this is screen shot provided by Angus long time ago, from a Pivot Table created in Excel, based on some information created with the above mentioned function):

            EURPRD10> $timeline = Build-TimeTrackerTable -MrsJob $stat
            EURPRD10> $timeline | Export-Csv 'tmp.csv'


C.	Priority of modules:
    Should be present in Version 1:     A.1., A.2., B.2.a.
    Can be introduced in Version 2.:    B.1., B.2.b.


D.	Resource estimates:

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
    [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
    [string]$AdminAccount
)

#endregion Parameters

################################################
# Common space for functions, global variables #
################################################
#region Functions, Global variables


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
    
    Clear-Host
    Write-Host "We are creating the working folder with " -ForegroundColor Cyan -NoNewline
    Write-Host "`"<MMddyyyy_HHmmss>`"" -ForegroundColor White -NoNewline
    Write-Host " format, under " -ForegroundColor Cyan -NoNewline
    Write-Host "`"%temp%\MigrationAnalyzer\`"" -ForegroundColor White

    ### <summary>
    ### TheDateToUse variable is used to collect the current date&time in the "MMddyyyy_HHmmss" format.
    ### </summary>
    $TheDateToUse = (Get-Date).ToString("MMddyyyy_HHmmss")

    ### <summary>
    ### WorkingDirectory variable is initialized to "%temp%\MigrationAnalyzer\<MMddyyyy_HHmmss>".
    ### </summary>
    $WorkingDirectory = "$env:temp\MigrationAnalyzer\$TheDateToUse"

    ### <summary>
    ### Creating the Working directory in the desired format.
    ### </summary>
    if (-not (Test-Path $WorkingDirectory)) {
        try {
            $void = New-Item -ItemType Directory -Force -Path $WorkingDirectory -ErrorAction Stop            
            $WorkingDirectoryToUse = $WorkingDirectory
        }
        catch {
            ### <summary>
            ### In case of error, we will retry to create the Working directory, for maximum 5 times.
            ### </summary>
            if ($NumberOfChecks -le 5) {
                if (Test-Path $WorkingDirectory){
                    $WorkingDirectoryToUse = $WorkingDirectory
                }
                else {
                    $NumberOfChecks++
                    $WorkingDirectoryToUse = Create-WorkingDirectory -NumberOfChecks $NumberOfChecks    
                }
            }
            ### <summary>
            ### In case we will not be able to create the Working directory even after 5 times, we will set the value of WorkingDirectoryToUse
            ### variable to NotAbleToCreateTheWorkingDirectory.
            ### </summary>            
            else {
                $WorkingDirectoryToUse = "NotAbleToCreateTheWorkingDirectory"
            }
        }
    }

    ### <summary>
    ### Checking if we were able to create the Working Directory in the desired location. If not, we will ask to insert the path where it can be created,
    ### from the keyboard.
    ### </summary>      
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
    
        ### <summary>
        ### If entered value will be empty, the script will exit.
        ### </summary>  
        if (-not ($WorkingDirectoryToUse)) {
            Write-Host
            Write-Host "No valid path was provided. The script will exit now." -ForegroundColor Red
            Exit
        }
        else {
            ### <summary>
            ### Doing 1-time effort to create the Working Directory in the location inserted from keyboard
            ### </summary>  
            try {
                $void = New-Item -ItemType Directory -Force -Path $WorkingDirectoryToUse -ErrorAction Stop            
            }
            catch {
                ### <summary>
                ### In case of error, we will exit the script.
                ### </summary>
                Write-Host
                Write-Host "We were unable to create the Working Directory under: " -ForegroundColor Red -NoNewline
                Write-Host "$WorkingDirectoryToUse" -ForegroundColor White
                Write-Host "The script will exit now." -ForegroundColor Red
                Exit
            }
        }
    }
    ### <summary>
    ### We successfully created a Working Directory. We will set it as current path (Set-Location -Path $WorkingDirectoryToUse)
    ### </summary>  
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

    ### <summary>
    ### Create-WorkingDirectory function will return the Path of the Working directory, or NotAbleToCreateTheWorkingDirectory in case
    ### we were unable to create the Working directory.
    ### </summary>  
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
        $NonInteractive
    )
	
	### <summary>
    ### Collecting the current date
    ### </summary>
	[string]$date = Get-Date -Format G
		
	### <summary>
    ### Write everything to LogFile
    ### </summary>
	( "[" + $date + "] || " + $string) | Out-File -FilePath $script:LogFile -Append
	
	### <summary>
    ### In case NonInteractive is not True, write on display, too
    ### </summary>
	if (!($NonInteractive)){
		( "[" + $date + "] || " + $string) | Write-Host
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

    ### <summary>
    ### LogFile variable (Scope: Script) is initialized to "$WorkingDirectory\MigrationAnalyzer.log".
    ### </summary>
    $script:LogFile = "$WorkingDirectory\MigrationAnalyzer.log"

    try {
        ### <summary>
        ### Creating the LogFile.
        ### </summary>
        $void = New-Item -ItemType "file" -Path "$script:LogFile" -Force -ErrorAction Stop
    }
    catch {
        ### <summary>
        ### In case of error, the script will exit.
        ### </summary>
        Write-Host
        Write-Host "You do not have permissions to create files under: " -ForegroundColor Red -NoNewline
        Write-Host "$WorkingDirectory" -ForegroundColor White
        Write-Host "The script will exit now..." -ForegroundColor Red

        Restore-OriginalState
        Exit
    }

    ### <summary>
    ### In case of success, we will log the first entry in the LogFile.
    ### </summary>    
    Write-Log "[INFO] || Logfile successfully created. Its location is $script:LogFile"
}

### <summary>
### Check-Parameters function is checking if the script was started with specific parameters.
### </summary>
function Check-Parameters {

    ### <summary>
    ### If FilePath parameter of the script was used, we will continue on this path.
    ### </summary>
    if ($FilePath){
        Write-Log ("[INFO] || The script was started with the FilePath parameter: `"-FilePath $FilePath`"")
        Selected-FileOption -FilePath $FilePath
    }
    ### <summary>
    ### If ConnectToExchangeOnline parameter of the script was used, we will continue on this path.
    ### </summary>
    elseif (ConnectToExchangeOnline) {
        Selected-ConnectToExchangeOnlineOption -FilePath $FilePath
        New-CleanO365Session 
    }
    ### <summary>
    ### If ConnectToExchangeOnPremises parameter of the script was used, we will continue on this path.
    ### </summary>
    elseif (ConnectToExchangeOnPremises) {
        
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
    ### <summary>
    ### If FilePath was provided, the script will use it in order to validate if the information from this variable is a correct
    ### full path of an .xml file.
    ### </summary>
    if ($FilePath){
        try {
            ### <summary>
            ### The script validates that the path provided is of a valid .xml file.
            ### </summary>
            Write-Log "[INFO] || We are validating if `"$FilePath`" is the full path of a .xml file"
            [string]$PathOfXMLFile = Validate-XMLPath -XMLFilePath $FilePath
        }
        catch {
            ### <summary>
            ### In case of error, the script will ask to provide again the full path of the .xml file
            ### </summary>
            [string]$PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $NumberOfChecks
        }
    }
    ### <summary>
    ### If no FilePath was provided, the script will ask to provide the full path of the .xml file
    ### </summary>
    else{
        [string]$PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks
    }

    ### <summary>
    ### If PathOfXMLFile variable will match "NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed", we will continue the data collection
    ### using other methods.
    ### </summary>    
    if ($PathOfXMLFile -match "NotAValidXMLFile|NotAValidPath|ValidationOfFileFailed") {
        [int]$TheNumberOfChecks = 1
    
        ### <summary>
        ### TheAffectedUser variable will represent the Affected user for which we will try to collect mailbox migration related logs
        ### </summary> 
        Write-Log "[INFO] || Trying to collect the AffectedUser..."
        [string]$TheAffectedUser = Ask-ForDetailsAboutUser -NumberOfChecks $TheNumberOfChecks
        
        ### <summary>
        ### TheMigrationType variable will represent the Migration type for which the logs have to be investigated
        ### </summary>
        Write-Log "[INFO] || Trying to collect the Migration Type..."
        [string]$TheMigrationType = Ask-DetailsAboutMigrationType -NumberOfChecks $TheNumberOfChecks -AffectedUser $TheAffectedUser

        ### <summary>
        ### TheMigrationLogs variable will represent MigrationLogs collected using the Selected-ConnectToExchangeOnlineOption function.
        ### </summary>
        Write-Log "[INFO] || Trying to collect the Migration Logs using Selected-FileOption -> Selected-ConnectToExchangeOnlineOption function..."
        $script:TheMigrationLogs = Selected-ConnectToExchangeOnlineOption -AffectedUser $TheAffectedUser -MigrationType $TheMigrationType
    }
    else {
        ### <summary>
        ### TheMigrationLogs variable will represent MigrationLogs collected using the Collect-MigrationLogs function.
        ### </summary>
        Write-Log "[INFO] || Trying to collect the Migration Logs using Selected-FileOption -> Collect-MigrationLogs function..."
        $script:TheMigrationLogs = Collect-MigrationLogs -XMLFile $PathOfXMLFile
    }

    return $script:TheMigrationLogs
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

    ### <summary>
    ### Validating if the path has a length greater than 4, and if it is a of an .xml file
    ### </summary>
    Write-Log "[INFO] || Checking if the FilePath is a valid .xml file, from PowerShell's perspective"
    if (($XMLFilePath.Length -gt 4) -and ($XMLFilePath -like "*.xml")) {
        ### <summary>
        ### Validating if the .xml file was created by PowerShell
        ### </summary>
        $fileToCheck = new-object System.IO.StreamReader($XMLFilePath)
        if ($fileToCheck.ReadLine() -like "*http://schemas.microsoft.com/powershell*") {
            Write-Host
            Write-Host $XMLFilePath -ForegroundColor Cyan -NoNewline
            Write-Host " seems to be a valid .xml file. We will use it to continue the investigation." -ForegroundColor Green
            Write-Log ("[INFO] || $XMLFilePath seems to be a valid .xml file. We will use it to continue the investigation.") -NonInteractive $true
        }
        ### <summary>
        ### If not, the script will set the XMLFilePath to NotAValidXMLFile. This will help in next checks, in order to start collecting the mailbox 
        ### migration logs using other methods
        ### </summary>
        else {
            Write-Host $XMLFilePath -ForegroundColor Cyan -NoNewline
            Write-Host " is not a valid .xml file." -ForegroundColor Yellow
            $XMLFilePath = "NotAValidXMLFile"
            Write-Log ("[WARNING] || $XMLFilePath is not a valid .xml file. We will set: XMLFilePath = `"NotAValidXMLFile`"") -NonInteractive $true
        }
        
        $fileToCheck.Close()

    }
    ### <summary>
    ### If the path's length is not greater than 4 characters and the file is not an .xml file the script will set XMLFilePath to NotAValidPath.
    ### This will help in next checks, in order to start collecting the mailbox migration logs using other methods
    ### </summary>
    else {
        Write-Host $XMLFilePath -ForegroundColor Cyan -NoNewline
        Write-Host " is not a valid path." -ForegroundColor Yellow
        $XMLFilePath = "NotAValidPath"
        Write-Log ("[WARNING] || $XMLFilePath is not a valid .xml file. We will set: XMLFilePath = `"NotAValidPath`"") -NonInteractive $true
    }
    
    ### <summary>
    ### The script returns the value of XMLFilePath 
    ### </summary>
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
        ### <summary>
        ### Asking to provide the full path of the .xml file for the first time
        ### </summary>
        Write-Host
        Write-Log "[INFO] || We are asking to provide the path of the .xml file" -NonInteractive $true
        Write-Host "Please provide the path of the .xml file: " -ForegroundColor Cyan
        Write-Host "`t" -NoNewline
        try {
            ### <summary>
            ### PathOfXMLFile variable will contain the full path of the .xml file, if it will be validated (it will be inserted from the keyboard)
            ### </summary>
            $PathOfXMLFile = Validate-XMLPath -XMLFilePath (Read-Host)
        }
        catch {
            ### <summary>
            ### If error, the script is doing the 1-time effort to collect again the full path of the .xml file
            ### </summary>
            $NumberOfChecks++
            $PathOfXMLFile = Ask-ForXMLPath -NumberOfChecks $NumberOfChecks
        }
    }
    else {
        ### <summary>
        ### The script is doing the 1-time effort to collect again the full path of the .xml file
        ### </summary>
        Write-Host
        Write-Log "[INFO] || Asking to provide the path of the .xml file again" -NonInteractive $true
        Write-Host "Would you like to provide the path of the .xml file again?" -ForegroundColor Cyan
        Write-Host "`t[Y] Yes`t`t[N] No`t`t(default is `"N`"): " -NoNewline -ForegroundColor White
        $ReadFromKeyboard = Read-Host

        ### <summary>
        ### Checking if the path will be provided again, or no. If no, we will continue to collect the mailbox migration logs, using other methods.
        ### </summary>
        [bool]$TheKey = $false
        Switch ($ReadFromKeyboard) 
        { 
          Y {$TheKey=$true} 
          N {$TheKey=$false} 
          Default {$TheKey=$false} 
        }

        if ($TheKey) {
            ### <summary>
            ### If YES was selected, we are asking to provide the path of the .xml file again
            ### </summary>
            Write-Host
            Write-Host "Please provide again the path of the .xml file: " -ForegroundColor Cyan
            Write-Host "`t" -NoNewline
            try {
                ### <summary>
                ### Validating the path of the .xml file
                ### </summary>
                $PathOfXMLFile = Validate-XMLPath -XMLFilePath (Read-Host)
            }
            catch {
                ### <summary>
                ### If error, the script will set PathOfXMLFile to ValidationOfFileFailed, which will be used to collect the logs using other methods
                ### </summary>
                Write-Host "We will continue to collect the migration logs using other methods" -ForegroundColor Red
                $PathOfXMLFile = "ValidationOfFileFailed"
            }
        }
        else {
            ### <summary>
            ### If NO was selected, the script will set PathOfXMLFile to ValidationOfFileFailed, which will be used to collect the logs using other methods
            ### </summary>
            Write-Host
            Write-Host "We will continue to collect the migration logs using other methods" -ForegroundColor Yellow
            $PathOfXMLFile = "ValidationOfFileFailed"
        }
    }
    
    ### <summary>
    ### The function returns the full path of the .xml file, or ValidationOfFileFailed
    ### </summary>
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
        ### <summary>
        ### Asking for the affected user, for the first time
        ### </summary>
        Write-Log "[INFO] || Asking to provide the affected user, for the first time." -NonInteractive $true
        Write-Host "Please provide the username of the affected user (Eg.: " -NoNewline -ForegroundColor Cyan
        Write-Host "User1@contoso.com" -NoNewline -ForegroundColor White
        Write-Host "): " -NoNewline -ForegroundColor Cyan
        $TheUserName = Read-Host
        $NumberOfChecks++
        Write-Log ("[INFO] || The affected user provided is: $TheUserName") -NonInteractive $true
    }
    else {
        ### <summary>
        ### Re-asking for the affected user
        ### </summary>
        Write-Log "[INFO] || Re-asking to provide the affected user." -NonInteractive $true
        Write-Host "Please provide again the username of the affected user (Eg.: " -NoNewline -ForegroundColor Cyan
        Write-Host "User1@contoso.com" -NoNewline -ForegroundColor White
        Write-Host "): " -NoNewline -ForegroundColor Cyan
        $TheUserName = Read-Host
        Write-Log ("[INFO] || The affected user provided is: $TheUserName") -NonInteractive $true
    }

    ### <summary>
    ### Validating if the user provided is the affected user
    ### </summary>
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
        ### <summary>
        ### Received confirmation that the user provided is the affected user.
        ### </summary>
        Write-Log ("[INFO] || Got confirmation that `"$TheUserName`" is indeed the affected user.") -NonInteractive $true
    }
    else {
        ### <summary>
        ### The user provided is not the affected user. Asking again for the affected user.
        ### </summary>
        Write-Log ("[WARNING] || `"$TheUserName`" is not the affected user. Starting over the process of asking for the affected user.") -NonInteractive $true
        [string]$TheUserName = Ask-ForDetailsAboutUser -NumberOfChecks $NumberOfChecks
    }

    ### <summary>
    ### The function will return the affected user
    ### </summary>
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
        ### <summary>
        ### Asking about the migration type, for the first time
        ### </summary>
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
        ### <summary>
        ### Re-asking about the migration type
        ### </summary>
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
    Write-Log "[IMFO] || You selected the following: `"Migration Type: $MigrationType`"; `"Is this correct? $ReadFromKeyboard`"" -NonInteractive $true

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

    ### <summary>
    ### The function returns the migration type
    ### </summary>
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

    ### <summary>
    ### We will try to connect to Exchange Online
    ### </summary>
    $ThePSSession = ConnectTo-ExchangeOnline -TheAdminAccount $TheAdminAccount

<#
    [int]$TheNumberOfChecks = 1
    [string]$ThePath = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks

    if ($ThePath -match "ValidationOfFileFailed") {
        [int]$TheNumberOfChecks = 1
        [string]$TheUser = Ask-ForDetailsAboutUser -NumberOfChecks $TheNumberOfChecks
        [int]$TheNumberOfChecks = 1
        [string]$TheMigrationType = Ask-DetailsAboutMigrationType -NumberOfChecks $TheNumberOfChecks -AffectedUser $TheUser
        Select-Option2
        #$ThePSSession = Connect-ToExchangeOnline #### Not done yet
        #$TheMigrationLogs = Collect-MigrationLogs -UserName $TheUser -MigrationType $TheMigrationType #### Not done yet
    }
    else {
        #$TheMigrationLogs = Collect-MigrationLogs  #### Not done yet
    }
#>  
}

### <summary>
### Selected-ConnectToExchangeOnlineOption function is used to connect to Exchange Online, and collect from there the mailbox migration logs,
### for the affected user, by running the correct commands, based on the migration type
### </summary>
### <param name="TheAdminAccount">TheAdminAccount represents username of an Admin that we will use in order to connect to Exchange Online </param>
Function ConnectTo-ExchangeOnline {
    [CmdletBinding()]
    param (
        [string]
        $TheAdminAccount
    )
    
    if ($TheAdminAccount) {
        $script:Credential = Get-Credential -UserName $TheAdminAccount -Message "Please provide credentials to connect to Exchange Online:"
    }

	# If we don't have a credential then prompt for it
	$i = 0
	while (($script:Credential -eq $Null) -and ($i -lt 5)){
		$script:Credential = Get-Credential -Message "Please provide your Exchange Online Credentials"
		$i++
	}
	
	# If we still don't have a credentail object then abort
	if ($Credential -eq $null){
		Write-Log "[Error] || Failed to get credentials"
	}

	Write-Log "[INFO] || Removing all PS Sessions"

	# Destroy any outstanding PS Session
	Get-PSSession | Remove-PSSession -Confirm:$false
	
	# Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
	[System.GC]::Collect()
	
	# Sleep 15s to allow the sessions to tear down fully
	Write-Log "[INFO] || Sleeping 15 seconds for Session Tear Down"
	Start-SleepWithProgress -SleepTime 15

	# Clear out all errors
	$Error.Clear()
	
	# Create the session
	Write-Log "[INFO] || Creating new PS Session"
	
	$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection
		
	# Check for an error while creating the session
	if ($Error.Count -gt 0){
	
		Write-Log "[ERROR] || Error while setting up session"
		Write-log ("[ERROR] || $Error")
		
		# Increment our error count so we abort after so many attempts to set up the session
		$ErrorCount++
		
		# if we have failed to setup the session > 3 times then we need to abort because we are in a failure state
		if ($ErrorCount -gt 3){
		
			Write-log "[ERROR] || Failed to setup session after multiple tries"
			Write-log "[ERROR] || Aborting Script"
			exit
		
		}
		
		# If we are not aborting then sleep 60s in the hope that the issue is transient
		Write-Log "[INFO] || Sleeping 60s so that issue can potentially be resolved"
		Start-SleepWithProgress -sleeptime 60
		
		# Attempt to set up the sesion again
		New-CleanO365Session
	}
	
	# If the session setup worked then we need to set $errorcount to 0
	else {
		$ErrorCount = 0
	}
	
	# Import the PS session
	$null = Import-PSSession $session -AllowClobber -Prefix EXO
	
	# Set the Start time for the current session
	Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
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
        [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
        [string]$AdminAccount
    )
    
    if ($XMLFile) {
        ### <summary>
        ### Importing data in the MigrationLogsToAnalyze (Scope: Script) variable
        ### </summary>
        Write-Log ("[INFO] || Importing data from `"$XMLFile`" file, in the MigrationLogsToAnalyze variable" )
        $script:MigrationLogsToAnalyze = Import-Clixml $XMLFile
    }
    elseif ($ConnectToExchangeOnline) {
        Write-Host "This part is not yet implemented" -ForegroundColor Red
    }
    elseif ($ConnectToExchangeOnPremises) {
        Write-Host "This part is not yet implemented" -ForegroundColor Red
    }
}

#endregion Functions, Global variables

###############
# Main script #
###############
#region Main script

$script:TheWorkingDirectory = Create-WorkingDirectory -NumberOfChecks 1
Create-LogFile -WorkingDirectory $script:TheWorkingDirectory

Check-Parameters
#Show-Menu

Restore-OriginalState

#endregion Main script

############################
#####################################
# Create / update .xml / .json file #
#####################################