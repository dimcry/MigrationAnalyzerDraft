function CheckIf-ExchangeManagementShell {
    param (
        [switch]
        $ExchangeOnPremises
    )

    [bool]$ExchangeManagementShell = $false
    try {
        Get-Command Get-ExBlog -ErrorAction Stop
        Write-Log "[INFO] || The script was started from Exchange Management Shell"
        $ExchangeManagementShell = $true
    }
    catch {
        Write-Log "[INFO] || The script was not started from Exchange Management Shell"
    }
    
    if ($ExchangeManagementShell) {
        if (-not ($ExchangeOnPremises)) {
            throw "[ERROR] || You started the script from Exchange Management Shell, even if you do not specifically want to run it in Exchange Management Shell.`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
        }
        else {
            $TheModules = Get-Module | where {($_.ModuleType -eq "Script") -and ($_.Name -like "tmp*")}
    
            if ($($TheModules.Count) -gt 0) {
                Write-Log "[INFO] || Found $($TheModules.Count) modules of type `"Script`", for which the Name starts with `"tmp`""
                [System.Collections.ArrayList]$script:EXOModules = @()
                [System.Collections.ArrayList]$script:OnPremModules = @()
                [System.Collections.ArrayList]$script:NotUsefulModules = @()
                foreach ($Module in $TheModules) {
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
                    [string]$EXOCommand = "Get-" + $Prefix + "SyncRequest"
                    [string]$OnPremCommand = "Get-" + $Prefix + "ExchangeServer"

                
                    if ($($Module.ExportedCommands["$EXOCommand"])) {
                        Write-Log ("[INFO] || This module (`"$($Module.Name)`") is related to Exchange Online") -ForegroundColor Green
                        $void = $script:EXOModules.Add($Module)
                    }
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

            if ([int]$($script:EXOModules.Count) -eq 1) {
                throw "[ERROR] || In this Exchange Management Shell session, the script found also 1 module with Exchange Online related commands.`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
            }
            elseif ([int]$($script:EXOModules.Count) -gt 1) {
                throw "[ERROR] || In this Exchange Management Shell session, the script found more than 1 module with Exchange Online related commands (found $($script:EXOModules.Count) modules).`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
            }
            elseif ([int]$($script:OnPremModules.Count) -eq 1) {
                throw "[ERROR] || In this Exchange Management Shell session, the script found also 1 module with Exchange OnPremises related commands.`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
            }
            elseif ([int]$($script:OnPremModules.Count) -gt 1) {
                throw "[ERROR] || In this Exchange Management Shell session, the script found more than 1 module with Exchange OnPremises related commands (found $($script:OnPremModules.Count) modules).`nIf the script has to run on the Exchange Online environment, please restart the script from a PowerShell window that is not already connected to Exchange Online.`nIf the script has to run on the Exchange OnPremises environment, please start a new Exchange Management Shell window, and start the script directly from it (using the -ConnectToExchangeOnPremises switch)."
            }
        }
    }
    return $ExchangeManagementShell
}

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
        ( "[" + $date + "] || " + $string) | Write-Host -ForegroundColor $ForegroundColor
    }
}

try {
    $script:LogFile = ".\Log.log"
    [bool]$EMS = CheckIf-ExchangeManagementShell -ExchangeOnPremises
}
catch {
    Write-Log "[ERROR] || $_" -ForegroundColor Red
    Write-Host
    Write-Log "[ERROR] || Script will now exit" -ForegroundColor Red
}