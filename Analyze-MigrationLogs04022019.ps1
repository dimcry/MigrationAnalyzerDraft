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

<#
[CmdletBinding()]
Param(
    [Parameter(ParameterSetName = "FilePath", Mandatory = $true)]
    [String]$FilePath = $null,

    [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $true)]
    $ConnectToExchangeOnline = $null,

    [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $true)]
    $ConnectToExchangeOnPremises = $null,

    [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
    [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
    [string[]]$AffectedUser = $null,
    
    [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
    [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
    [ValidateSet("Hybrid", "IMAP", "Cutover", "Staged")]
    [string[]]$MigrationType = $null,

    [Parameter(ParameterSetName = "ConnectToExchangeOnline", Mandatory = $false)]
    [System.Management.Automation.PSCredential]$AdminCredentials,

    [Parameter(ParameterSetName = "ConnectToExchangeOnPremises", Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential
)
#>

################################################
# Common space for functions, global variables #
################################################


function Show-Menu {

    # Read cred stored
    $global:CredPath = "$HOME\PSSecureCredentials"


     
    if((test-path $global:CredPath) -eq $false) {
        md $CredPath
    }


    $menu=@"
1 => If you have the migration logs in an .xml file
2 => If you want to connect to Exchange Online in order to collect the logs
3 => If you need to connect to Exchange On-Premises and collect the logs
Q => Quit
Select a task by number or Q to quit: 
"@
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
    Write-Host $menu -ForegroundColor Cyan -NoNewline
    $SwitchFromKeyboard = Read-Host

    Switch ($SwitchFromKeyboard) {

        "1" {
            [int]$TheNumberOfChecks = 1
            [string]$ThePath = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks
        }

        "2" {
            Write-Host "=== Connect to AADRM (AIP), MSOL and Exchange Online ===" -ForegroundColor Green
            Connect-AADRMandEXO
            sleep -Seconds 2
            Check-OMEStatus
            Read-Host "Press [ENTER] to reload the menu"
            Show-Menu
        }
 
        "3" {
            Write-Host "`n=== Configure OME to use only v1 ===" -ForegroundColor Green
            Config-OMEv1
            Read-Host "Press [ENTER] to reload the menu"
            Show-Menu
        }

        "Q" {
            Write-Host "`n=== Quitting ===" -ForegroundColor Cyan
            try
            {
                Disconnect-AadrmService
                Write-Host "Disconnecting Exchange Online PS Session" -ForegroundColor Cyan
                Remove-PSSession $global:session
                Exit
            }
            Catch {}
         }
 
        default {
            Write-Host "`n=== I don't understand what you want to do ===" -ForegroundColor Yellow
            Read-Host "Press [Enter] to re-load the menu"
            Show-Menu
        }
    } 
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
        Write-Host "The path you provided is not valid!" -ForegroundColor Red

        Write-Host "Would you like to provide it again?" -ForegroundColor Cyan
        Write-Host "`t[Y] Yes     [N] No      (default is `"N`"): " -NoNewline -ForegroundColor White
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
                Write-Host "The path you provided is still not valid" -ForegroundColor Red
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
        $filePath
    )

    if (($filePath.Length -gt 4) -and ($filePath -like "*.xml")) {
        Write-Host
        Write-Host $filePath -ForegroundColor Cyan -NoNewline
        Write-Host " is a valid .xml file. We will use it to continue the investigation" -ForegroundColor Green
    }
    else {
        $filePath = "NotAValidPath"
    }

    return $filePath
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
function Connect-ToExchangeOnline {

    Clear-Host
    Write-Host "We are connecting you to Exchange Online..." -ForegroundColor Cyan
    Write-Host

    Write-Host "We will need credentials of an Admin in order to collect the correct migration logs related "
    $cred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection

    return $Session
    
}


###############
# Main script #
###############

Get-PSSession | Remove-PSSession

Clear-Host
Show-Menu
<#
[int]$TheNumberOfChecks = 1
[string]$ThePath = Ask-ForXMLPath -NumberOfChecks $TheNumberOfChecks

if ($ThePath -match "ValidationOfFileFailed") {
    [int]$TheNumberOfChecks = 1
    [string]$TheUser = Ask-ForDetailsAboutUser -NumberOfChecks $TheNumberOfChecks
    [int]$TheNumberOfChecks = 1
    [string]$TheMigrationType = Ask-DetailsAboutMigrationType -NumberOfChecks $TheNumberOfChecks -AffectedUser $TheUser
    #$ThePSSession = Connect-ToExchangeOnline #### Not done yet
    #$TheMigrationLogs = Collect-MigrationLogs -UserName $TheUser -MigrationType $TheMigrationType #### Not done yet
}
else {
    #$TheMigrationLogs = Collect-MigrationLogs  #### Not done yet
}

#>

############################
#####################################
# Create / update .xml / .json file #
#####################################
<#
### If .xml, try not to use Import/Export-Clixml

[System.Collections.ArrayList]$ErrorsAndRecommendations = @()

$TheXML = New-Object -TypeName PSObject -ErrorAction SilentlyContinue
$TheXML | Add-Member -Type NoteProperty -Name "Failure Type" -Value "NoValue"
$TheXML | Add-Member -Type NoteProperty -Name "Actual Issue" -Value "NoValue"
$TheXML | Add-Member -Type NoteProperty -Name "Recommendation" -Value "NoValue"
$void = $ErrorsAndRecommendations.Add($TheXML)

$TheXML1 = New-Object -TypeName PSObject -ErrorAction SilentlyContinue
$TheXML1 | Add-Member -Type NoteProperty -Name "Failure Type" -Value "NoValue"
$TheXML1 | Add-Member -Type NoteProperty -Name "Actual Issue" -Value "NoValue"
$TheXML1 | Add-Member -Type NoteProperty -Name "Recommendation" -Value "NoValue"
$void = $ErrorsAndRecommendations.Add($TheXML1)

$TheXML2 = New-Object -TypeName PSObject -ErrorAction SilentlyContinue
$TheXML2 | Add-Member -Type NoteProperty -Name "Failure Type" -Value "NoValue"
$TheXML2 | Add-Member -Type NoteProperty -Name "Actual Issue" -Value "NoValue"
$TheXML2 | Add-Member -Type NoteProperty -Name "Recommendation" -Value "NoValue"
$void = $ErrorsAndRecommendations.Add($TheXML2)

$ErrorsAndRecommendations | Export-Clixml .\ErrorsAndRecommendations1.xml -Force
$JSon_ErrorsAndRecommendations = $ErrorsAndRecommendations | ConvertTo-Json -Depth 10
$JSon_ErrorsAndRecommendations | Out-File .\JSon_ErrorsAndRecommendations.json -Force
#>
############################