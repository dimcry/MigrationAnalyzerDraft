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

########################################
# Common space for script's parameters #
########################################

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


################################################
# Common space for functions, global variables #
################################################


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

    $TheDateToUse = (Get-Date).ToString("MMddyyyy_HHmmss")    
    $WorkingDirectory = "$env:temp\MigrationAnalyzer\$TheDateToUse"

    if (-not (Test-Path $WorkingDirectory)) {
        try {
            $void = New-Item -ItemType Directory -Force -Path $WorkingDirectory -ErrorAction Stop            
            $WorkingDirectoryToUse = $WorkingDirectory
        }
        catch {
            if ($NumberOfChecks -le 5) {
                if (Test-Path $WorkingDirectory){
                    $WorkingDirectoryToUse = $WorkingDirectory
                }
                else {
                    $NumberOfChecks++
                    $WorkingDirectoryToUse = Create-WorkingDirectory -NumberOfChecks $NumberOfChecks    
                }
            }
            else {
                $WorkingDirectoryToUse = "NotAbleToCreateTheWorkingDirectory"
            }
        }
    }

    return $WorkingDirectoryToUse
}

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
	
	# Get the current date
	[string]$date = Get-Date -Format G
		
	# Write everything to our log file
	( "[" + $date + "] || " + $string) | Out-File -FilePath $LogFile -Append
	
	# If NonInteractive true then supress host output
	if (!($NonInteractive)){
		( "[" + $date + "] || " + $string) | Write-Host
	}
}

function Show-Menu {

    $menu=@"

1 => If you have the migration logs in an .xml file
2 => If you want to connect to Exchange Online in order to collect the logs
3 => If you need to connect to Exchange On-Premises and collect the logs
Q => Quit

Select a task by number, or, Q to quit: 
"@
    $menuprompt = $null

    Write-Log "[INFO] || Loading the menu..." -NonInteractive $true

    Clear-Host
    $title = "=== Mailbox migration analyzer ==="
    if (!($menuprompt)) 
    {
        $menuprompt+="="*$title.Length
    }
    Write-Host $menuprompt
    Write-Host $title
    Write-Host $menuprompt
    Write-Host $menu -ForegroundColor Cyan -NoNewline
    $SwitchFromKeyboard = Read-Host

    Switch ($SwitchFromKeyboard) {

        "1" {
            Write-Log "[INFO] || You selected to provide an .xml to be analyzed."
            Select-Option1
        }

        "2" {
            Write-Log "[INFO] || You selected to connect to Exchange Online and collect from there correct migration logs to be analyzed."
            Select-Option2
        }
 
        "3" {
            Write-Log "[INFO] || You selected to connect to Exchange On-Premises and collect from there correct migration logs to be analyzed."
            Select-Option3
        }

        "Q" {
            Write-Log "[INFO] || You selected to quit the menu. The script will exit now."
            try
            {
                Exit
            }
            Catch {}
         }
 
        default {
            Write-Log "[INFO] || You selected an option that is not present in the menu."
            Write-Log "[INFO] || Press any key to re-load the menu"
            Read-Host
            Show-Menu
        }
    } 
}

function Select-Option1 {
    [CmdletBinding()]
    Param
    (        
        [string]
        $FilePath
    )

    [int]$TheNumberOfChecks = 1
    if ($FilePath){
        try {
            $PathOfXMLFile = Validate-XMLPath -filePath $FilePath
        }
        catch {
            $NumberOfChecks++
            Ask-ForXMLPath -NumberOfChecks $NumberOfChecks
        }
    }
    else{
        [string]$ThePath = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks
    }

    if ($ThePath -match "ValidationOfFileFailed") {
        [int]$TheNumberOfChecks = 1
        [string]$TheUser = Ask-ForDetailsAboutUser -NumberOfChecks $TheNumberOfChecks
        [int]$TheNumberOfChecks = 1
        [string]$TheMigrationType = Ask-DetailsAboutMigrationType -NumberOfChecks $TheNumberOfChecks -AffectedUser $TheUser
        Select-Option2 -AffectedUser $TheUser -MigrationType $TheMigrationType
        #$ThePSSession = Connect-ToExchangeOnline #### Not done yet
        #$TheMigrationLogs = Collect-MigrationLogs -UserName $TheUser -MigrationType $TheMigrationType #### Not done yet
    }
    else {
        $TheMigrationLogs = Collect-MigrationLogs -XMLFile $ThePath
    }
    
}

function Select-Option2 {
    [CmdletBinding()]
	Param (
        [string]
        $AffectedUser,        
        [string]
        $MigrationType
    )

    Clear-Host
    $ThePSSession = New-CleanO365Session

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
        Clear-Host
        Write-Log "[INFO] || We are asking to provide the path of the .xml file" -NonInteractive $true
        Write-Host "Please provide the path of the .xml file: " -ForegroundColor Cyan
        Write-Host "`t" -NoNewline
        try {
            $PathOfXMLFile = Validate-XMLPath -filePath (Read-Host)
        }
        catch {
            $NumberOfChecks++
            Ask-ForXMLPath -NumberOfChecks $NumberOfChecks
        }
    }
    else {
        Write-Host

        Write-Host "Would you like to provide the path of the .xml file again?" -ForegroundColor Cyan
        Write-Host "`t[Y] Yes`t`t[N] No`t`t(default is `"N`"): " -NoNewline -ForegroundColor White
        $ReadFromKeyboard = Read-Host

        [bool]$TheKey = $false
        Switch ($ReadFromKeyboard) 
        { 
          Y {$TheKey=$true} 
          N {$TheKey=$false} 
          Default {$TheKey=$false} 
        }

        if ($TheKey) {
            Write-Host
            Write-Host "Please provide again the path of the .xml file: " -ForegroundColor Cyan
            Write-Host "`t" -NoNewline
            try {
                $PathOfXMLFile = Validate-XMLPath -filePath (Read-Host)
            }
            catch {
                Write-Host "We will continue to collect the migration logs using other methods" -ForegroundColor Red
                $PathOfXMLFile = "ValidationOfFileFailed"
            }
        }
        else {
            Write-Host
            Write-Host "We will continue to collect the migration logs using other methods" -ForegroundColor Yellow
            $PathOfXMLFile = "ValidationOfFileFailed"
        }
    }
    return $PathOfXMLFile
}

function Validate-XMLPath {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_})]
        [string]
        $XMLFilePath
    )

    $fileToCheck = new-object System.IO.StreamReader($XMLFilePath)
    if (($XMLFilePath.Length -gt 4) -and ($XMLFilePath -like "*.xml")) {
        if ($fileToCheck.ReadLine() -like "*http://schemas.microsoft.com/powershell*") {
            Write-Host
            Write-Host $XMLFilePath -ForegroundColor Cyan -NoNewline
            Write-Host " seems to be a valid .xml file. We will use it to continue the investigation." -ForegroundColor Green
        }
        else {
            Write-Host $XMLFilePath -ForegroundColor Cyan -NoNewline
            Write-Host " is not a valid .xml file." -ForegroundColor Yellow
            $XMLFilePath = "NotAValidXMLFile"
        }
    }
    else {
        Write-Host $XMLFilePath -ForegroundColor Cyan -NoNewline
        Write-Host " is not a valid path." -ForegroundColor Yellow
        $XMLFilePath = "NotAValidPath"
    }

    $fileToCheck.Close()

    return $XMLFilePath
}

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
        
        Write-Host "Please provide the username of the affected user (Eg.: " -NoNewline -ForegroundColor Cyan
        Write-Host "User1@contoso.com" -NoNewline -ForegroundColor White
        Write-Host "): " -NoNewline -ForegroundColor Cyan
        $TheUserName = Read-Host
        $NumberOfChecks++
    }
    else {
        Write-Host "Please provide again the username of the affected user (Eg.: " -NoNewline -ForegroundColor Cyan
        Write-Host "User1@contoso.com" -NoNewline -ForegroundColor White
        Write-Host "): " -NoNewline -ForegroundColor Cyan
        $TheUserName = Read-Host
    }

    Clear-Host
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
        return $TheUserName
    }
    else {
        [string]$TheUser = Ask-ForDetailsAboutUser -NumberOfChecks $NumberOfChecks
        return $TheUser
    }

}

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

    Clear-Host
    if ($NumberOfChecks -eq "1") {
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

        Switch ($ReadFromKeyboard) 
        { 
          1 {$MigrationType="Hybrid"} 
          2 {$MigrationType="IMAP"} 
          3 {$MigrationType="Cutover"}
          4 {$MigrationType="Staged"}
          Default {$MigrationType="Hybrid"} 
        }
    }

    Clear-Host
    Write-Host "You entered " -NoNewline -ForegroundColor Cyan
    Write-Host "$MigrationType" -NoNewline -ForegroundColor White
    Write-Host ". Is this correct?" -ForegroundColor Cyan
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
        return $MigrationType
    }
    else {
        [string]$AMigrationType = Ask-DetailsAboutMigrationType -NumberOfChecks $NumberOfChecks -AffectedUser $AffectedUser
        return $AMigrationType
    }
}

function Collect-MigrationLogs {
    [CmdletBinding()]
	Param (
        [parameter(Mandatory=$true,
        ParameterSetName="XMLFile")]
        [string]
        $XMLFile,
        [parameter(Mandatory=$true,
        ParameterSetName="ConnectToExchangeOnline")]
        [string]
        $ConnectToExchangeOnline,
        [parameter(Mandatory=$true,
        ParameterSetName="ConnectToExchangeOnPremises")]
        [string]
        $ConnectToExchangeOnPremises
    )
    
    if ($XMLFile) {
        $script:MigrationLogsToAnalyze = Import-Clixml $XMLFile
    }
    elseif ($ConnectToExchangeOnline) {
        #$ThePSSession = Connect-ToExchangeOnline #### Not done yet
        #$TheMigrationLogs = Collect-MigrationLogs -UserName $TheUser -MigrationType $TheMigrationType #### Not done yet
    }
    elseif ($ConnectToExchangeOnPremises) {
        #$ThePSSession = Connect-ToExchangeOnPremises #### Not done yet
        #$TheMigrationLogs = Collect-MigrationLogs -UserName $TheUser -MigrationType $TheMigrationType #### Not done yet
    }
}


# Sleeps X seconds and displays a progress bar
Function Start-SleepWithProgress {
	Param([int]$sleeptime)

	# Loop Number of seconds you want to sleep
	For ($i=0;$i -le $sleeptime;$i++){
		$timeleft = ($sleeptime - $i);
		
		# Progress bar showing progress of the sleep
		Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		
		# Sleep 1 second
		start-sleep 1
	}
	
	Write-Progress -Completed -Activity "Sleeping"
}

# Setup a new O365 Powershell Session
Function New-CleanO365Session {
    [CmdletBinding()]
    param (
        [System.Management.Automation.PSCredential]
        $EXOCredential
    )
    
    if ($EXOCredential) {
        $Credential = $EXOCredential
    }

	# If we don't have a credential then prompt for it
	$i = 0
	while (($Credential -eq $Null) -and ($i -lt 5)){
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

# Verifies that the connection is healthy
# Goes ahead and resets it every $ResetSeconds number of seconds either way
Function Test-O365Session {
	
	# Get the time that we are working on this object to use later in testing
	$ObjectTime = Get-Date
	
	# Reset and regather our session information
	$SessionInfo = $null
	$SessionInfo = Get-PSSession
	
	# Make sure we found a session
	if ($SessionInfo -eq $null) { 
		Write-Log "[ERROR] || No Session Found"
		Write-log "[INFO] || Recreating Session"
		New-CleanO365Session
	}	
	# Make sure it is in an opened state if not log and recreate
	elseif ($SessionInfo.State -ne "Opened"){
		Write-Log "[ERROR] - Session not in Open State"
		Write-log ("[INFO] || $($SessionInfo | fl | Out-String)")
		Write-log "[INFO] || Recreating Session"
		New-CleanO365Session
	}
	# If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
	elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds){
		Write-Log ("[INFO] || Session Has been active for greater than " + $ResetSeconds + " seconds" )
		Write-Log "[INFO] || Rebuilding Connection"
		
		# Estimate the throttle delay needed since the last session rebuild
		# Amount of time the session was allowed to run * our activethrottle value
		# Divide by 2 to account for network time, script delays, and a fudge factor
		# Subtract 15s from the results for the amount of time that we spend setting up the session anyway
		[int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)
		
		# If the delay is >15s then sleep that amount for throttle to recover
		if ($DelayinSeconds -gt 0){
		
			Write-Log ("[INFO] || Sleeping " + $DelayinSeconds + " additional seconds to allow throttle recovery")
			Start-SleepWithProgress -SleepTime $DelayinSeconds
		}
		# If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
		else {
			Write-Log ("[INFO] || Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
		}
				
		# new O365 session and reset our object processed count
		New-CleanO365Session
	}
	else {
		# If session is active and it hasn't been open too long then do nothing and keep going
	}
	
	# If we have a manual throttle value then sleep for that many milliseconds
	if ($ManualThrottle -gt 0){
		Write-log ("[INFO] || Sleeping " + $ManualThrottle + " milliseconds")
		Start-Sleep -Milliseconds $ManualThrottle
	}
}

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

    ### $script:MigrationBatch = Get-EXOMigrationBatch $AffectedUser -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose"
}

function Check-Parameters {

    if ($FilePath){
        Select-Option1 -FilePath $FilePath
    }
    elseif (ConnectToExchangeOnline) {
        New-CleanO365Session 
    }
    elseif (ConnectToExchangeOnPremises) {
        
    }
}

###############
# Main script #
###############

$TheWorkingDirectory = Create-WorkingDirectory -NumberOfChecks 1
if ($TheWorkingDirectory -eq "NotAbleToCreateTheWorkingDirectory") {
    Write-Host "We were unable to create the working directory with " -ForegroundColor Red -NoNewline
    Write-Host "`"<MMddyyyy_HHmmss>`"" -ForegroundColor White -NoNewline
    Write-Host " format, under " -ForegroundColor Red -NoNewline
    Write-Host "`"%temp%\MigrationAnalyzer\`"" -ForegroundColor White
    Write-Host
    Write-Host "The script will exit now!" -ForegroundColor Red
    Exit
}
else {
    Write-Host "We successfully created the following working directory:" -ForegroundColor Green
    Write-Host "`tFull path: " -ForegroundColor Cyan -NoNewline
    Write-Host $TheWorkingDirectory -ForegroundColor White
    Write-Host "`tShort path: " -ForegroundColor Cyan -NoNewline
    $TheShortPath = ($TheWorkingDirectory -split "MigrationAnalyzer")[1]
    Write-Host "`%temp`%\MigrationAnalyzer$TheShortPath" -ForegroundColor White
    Set-Location -Path $TheWorkingDirectory
}

$LogFile = "$TheWorkingDirectory\MigrationAnalyzer.log"
Write-Log "[INFO] || Logfile successfully created. Its location is $LogFile"

Check-Parameters
Show-Menu

############################
#####################################
# Create / update .xml / .json file #
#####################################