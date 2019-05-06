### ExOnPremCommandsPrefix (Scope: Script) variable will be used to create a new PSSession to Exchange OnPremises.
### When importing the PSSession, the script will use "MAExOnP" (Migration Analyzer Exchange OnPremises) as Prefix for each command
[string]$script:ExOnPremCommandsPrefix = "MAExOnP"
### ExOnPremPSSessionCreated (Scope: Script) variable will be used to check if the Exchange OnPremises PSSession was successfully created
[bool]$script:ExOnPremPSSessionCreated = $false

function ConnectTo-ExchangeOnPremises {

    [CmdletBinding(DefaultParameterSetName='resolve',
                   HelpUri = 'https://gallery.technet.microsoft.com/Connect-to-one-or-multiple-b850411d'
    )]

    Param(

    [Parameter(Mandatory=$false)]
    [string]$ExchangeURL,

    [Parameter(Mandatory=$false)]
    [string]$AuthenticationType,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$Domain = "getdomain",

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$ADSite = "getsite",

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [String]$Prefix,

    [Parameter(Mandatory = $false)]
    [string]$OnPremAdminAccount,

    [Parameter(Mandatory = $false)]
    [int]$NumberOfChecks

    )

    #region Determine if a url is being specified

    if (([string]::IsNullOrEmpty($ExchangeURL))) {
        Write-Log "[INFO] || Function is set to use exchange URL: Auto-resolve"
        [bool]$autoresolveurl = $true
    }
    else {
        Write-Log ("[INFO] || Function is set to use exchange URL: $exchangeurl")
        [bool]$autoresolveurl = $false
    }

    #endregion Determine if a url is being specified

    
    if ($autoresolveurl) {
        #region Start auto resolve connection url
        #region Determine AD domain 

        if ($Domain -eq "getdomain") {
            Write-Log "[INFO] || Function will now query AD domain using .net"
            try {
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
            throw "Could not find the AD site. Please check you spelling if you used -ADsite parameter. If autoresolve is used a connectivity error occured"
        }

        #endregion craft exchange site filter     

        $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*))"
 
        #endregion Build exchange search filter
        #region Harvest exchange servers
        [Array]$Servers =@()
        $tempallServers = Get-Ldapobject -LDAPfilter $Filterexservers -configurationNamingContext -configurationNamingContextdomain $domain -Findall $true
        [Array]$Servers += $tempallServers

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
        }

        #region Final attempt Harvest exchange servers

        if ($Servers.count -eq 0) {
            Write-Log ("[INFO] || Function did resolve 0 servers in adjacent sites to ADsite $adsitename using filter: $Filterexservers ")
            Write-Log "Function last attempt: Any site , any version"
            $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer))"
            [Array]$Servers +=  Get-Ldapobject -LDAPfilter $Filterexservers -configurationNamingContext -configurationNamingContextdomain $domain -Findall $true
        }

        if ($Servers.count -eq 0) {
            throw "Function was unable to identify any Exchange servers in your organization"
        }
        else {
            [System.Collections.ArrayList]$E19Servers = @()
            [System.Collections.ArrayList]$E16Servers = @()
            [System.Collections.ArrayList]$E15Servers = @()
            [System.Collections.ArrayList]$E14Servers = @()
            [System.Collections.ArrayList]$OtherVersionServers = @()
            foreach ($Server in $Servers) {
                if ($($Server.Properties["serialnumber"]) -like "*Version 15.2*") {
                    $null = $E19Servers.Add($Server)
                }
                elseif ($($Server.Properties["serialnumber"]) -like "*Version 15.1*") {
                    $null = $E16Servers.Add($Server)
                }
                elseif ($($Server.Properties["serialnumber"]) -like "*Version 15.0*") {
                    $null = $E15Servers.Add($Server)
                }
                elseif ($($Server.Properties["serialnumber"]) -like "*Version 14.*") {
                    $null = $E14Servers.Add($Server)
                }
                else {
                    $null = $OtherVersionServers.Add($Server)
                }
            }
        }

        [Array]$Servers = @()
        if ($E19Servers) {
            foreach ($Server in $E19Servers) {
                $Servers += $Server
            }
        }
        elseif ($E16Servers) {
            foreach ($Server in $E16Servers) {
                $Servers += $Server
            }
        }
        elseif ($E15Servers) {
            foreach ($Server in $E15Servers) {
                $Servers += $Server
            }
        }
        elseif ($E14Servers) {
            foreach ($Server in $E14Servers) {
                $Servers += $Server
            }
        }
        elseif ($OtherVersionServers) {
            throw "In your Organization we were unable to identify any supported versions of Exchange server (2019 / 2016 / 2013 / 2010).`nWe identified $($OtherVersionServers.Count) servers with a version older than 2010.`nIf you have, a supported version of Exchange, please restart the script with the `"-ConnectToExchangeOnPremises -ExchangeURL http://mail.contoso.com/PowerShell`" parameters, and correct values."
        }
        else {
            throw "In your Organization we were unable to identify any Exchange server.`nIf you have, a supported version of Exchange server (2019 / 2016 / 2013 / 2010), please restart the script at least with the `"-ConnectToExchangeOnPremises -ExchangeURL http://mail.contoso.com/PowerShell`" parameters, and correct values."
        }


        #endregion Final attempt Harvest exchange servers



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
            }
            else {
                $serverconnection = $exchangeurl
            }

            $i = 0
            while ((-not ($script:ExOnPremCredential)) -and ($i -lt 5)){
                $script:ExOnPremCredential = Get-Credential $OnPremAdminAccount -Message "Please provide your Exchange OnPremises Credentials:"
            }

            if ([string]::IsNullOrWhiteSpace($OnPremAdminAccount)) {
                try {
                    $script:ExOnPremPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -ErrorAction Stop
                }
                catch {
                    Write-Log "[ERROR] || The script was unable to connect to Exchange On-Premises without explicit credentials"
                }
            }

            #if ($script:ExOnPremCredential) {
                if ([string]::IsNullOrWhiteSpace($authenticationtype)) {
                    try {
                        $script:ExOnPremPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -Credential $script:ExOnPremCredential -ErrorAction Stop
                    }
                    catch {
                        Write-Log "[ERROR] || The script was unable to connect to Exchange On-Premises using the provided credentials (of user $($script:ExOnPremCredential.UserName))"
                        try {
                            $script:ExOnPremPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -Credential $script:ExOnPremCredential -Authentication Basic -ErrorAction Stop
                        }
                        catch {
                            Write-Log "[ERROR] || The script was unable to connect to Exchange On-Premises using the provided credentials (of user $($script:ExOnPremCredential.UserName)), using the provided authentication type ($authenticationtype)"
                        }
                    }
                }
                else {
                    try {
                        $script:ExOnPremPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -Credential $script:ExOnPremCredential -Authentication $authenticationtype -ErrorAction Stop
                    }
                    catch {
                        Write-Log "[ERROR] || The script was unable to connect to Exchange On-Premises using the provided credentials (of user $($script:ExOnPremCredential.UserName)), using the provided authentication type ($authenticationtype)"
                    }
                }
            #}

            try {
                $null = Import-PSSession $script:ExOnPremPSSession -AllowClobber -Prefix $script:ExOnPremCommandsPrefix
                $script:ExOnPremPSSessionCreated = $true
                Write-Log "[INFO] || We managed to successfully import the Exchange OnPremises PSSession"
            }
            catch {
                ### If error, retry
                Write-Log "[ERROR] || We were unable to import the Exchange OnPremises PSSession" -ForegroundColor Red
                Write-log ("[ERROR] || $Error") -ForegroundColor Red
                $script:ExOnPremCredential = $null
                $NumberOfChecks++
                ConnectTo-ExchangeOnPremises -Prefix $script:ExOnPremCommandsPrefix -ExchangeURL $ExchangeURL -AuthenticationType $AuthenticationType -Domain $Domain -ADSite $ADSite -NumberOfChecks 1
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
                $script:ExOnPremPSSession = "failed"
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
                $script:ExOnPremPSSession = "failed"
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
    until ($script:ExOnPremPSSession)
} 

function Get-Ldapobject {
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
Â Â       12-12-2016            V1.0         Martijn van Geffen    Initial Script.
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
    [bool]$Findall = $false,
        
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

ConnectTo-ExchangeOnPremises -ExchangeURL https://owa.dimcry.ro/PowerShell

# Florin