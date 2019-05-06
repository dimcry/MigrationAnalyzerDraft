function Connect-ExchangeService
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

    Write-Verbose -Message "Function is set to use exchange version: $versionnr"

    #endregion Determine the exchange version

    #region Determine if a url is being specified

    if (([string]::IsNullOrEmpty($exchangeurl)))
    {
        Write-Verbose -Message "Function is set to use exchange URL: Auto-resolve"
        [boolean]$autoresolveurl = $true
    }else
    {
        Write-Verbose -Message "Function is set to use exchange URL: $exchangeurl"
        [boolean]$autoresolveurl = $false
    }

    #endregion Determine if a url is being specified

    
    if ($autoresolveurl)
    {
        #region Start auto resolve connection url

        #region Determine AD domain 

        if ($Domain -eq "getdomain")
        {
            try
            {
                Write-Debug -Message "Function will now query AD domain using .net"
                $Domain = ([system.directoryservices.activedirectory.domain]::GetCurrentDomain()).name
            }catch
            {
                Write-Verbose -Message "$($_.exception.message)"
                Write-Error -Message "Function Could not resolve Active directory domain and -Domain switch is not used."
                Throw "Aborting - Function Could not resolve Active directory domain and -Domain switch is not used."
                break
            }
        }

        #endregion Determine AD domain

        #region Determine AD site 

        if ($ADsite -eq "getsite")
        {
            $sitemanualset = $false
            try
            {
                Write-Debug -Message "Function will now query current computer AD site using .net"
                $ADsitename = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).name
            }catch
            {
                Write-Verbose -Message "$($_.exception.message)"
                Write-warning -Message "WARNING - Could not resolve the client AD site. The Function will continue autoresolve using any AD site as filter. You might end up on a slow AD site. Please validate your IP to AD site binding settings"
                $ADsitename = "*"                
            }
        }else
        {
            $ADsitename = $ADsite
            $sitemanualset = $true
        }

        Write-Verbose -Message  "Function is set to use ADsite: $ADsitename"

        #endregion Determine AD site

        #region Build exchange search filter

        #region craft exchange site filter

        $FilterADSite = "(&(objectclass=site)(Name=$ADsitename))"
        $ADsiteobject = Get-Ldapobject -LDAPfilter $FilterADSite -configurationNamingContext -configurationNamingContextdomain $domain
        $ADsiteobjectdn = $ADsiteobject.properties.distinguishedname
        
        if ([string]::IsNullOrEmpty($ADsiteobjectdn))
        {
            write-verbose -Message "failing AD query: $FilterADSite"
            throw "Error - Could not find the AD site. Please check you spelling if you used -ADsite parameter. If autoresolve is used a connectivity error occured"
            break    
        }

        #endregion craft exchange site filter

        if ($versionnr -eq "notdeclared")
        {
            $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*))"
        }else
        {
            $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*$versionnr*))"
        }
        
        #endregion Build exchange search filter

        #region Harvest exchange servers
        [array]$Servers =@()
        $tempallServers = Get-Ldapobject -LDAPfilter $Filterexservers -configurationNamingContext -configurationNamingContextdomain $domain -Findall $true
        [array]$Servers += $tempallServers
        
        if ($Servers.count -eq 0)
        {       
            Write-output -InputObject "Function did resolve 0 servers in the ADsite $adsit using filter: $Filterexservers "
            Write-Output -InputObject "Retrying with next closest AD sites"
            $adjecentsites = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).AdjacentSites.name
            Write-Output -InputObject "Found ajecentsites AD sites: $($adjecentsites -join ", ") "

            foreach ($site in $adjecentsites)
            {
                Write-debug -Message "Function trying to find servers in AD sites: $site "

                #clean itterative variables
                              
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
        
                if ([string]::IsNullOrEmpty($ADsiteobjectdn))
                {
                    write-verbose -Message "failing AD query: $FilterADSite"
                    throw "Error - Could not find the AD site. Please check you spelling if you used -ADsite parameter. If autoresolve is used a connectivity error occured"
                    break    
                }

                #endregion craft exchange site filter retry

                if ($versionnr -eq "notdeclared")
                {
                    $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*))"
                }else
                {
                    $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer)(msExchServerSite=$ADsiteobjectdn)(serialnumber=*$versionnr*))"
                }
                
                #endregion retry Build exchange search filter

                #region Retry Harvest exchange servers
                
                $tempallServers = Get-Ldapobject -LDAPfilter $Filterexservers -configurationNamingContext -configurationNamingContextdomain $domain -Findall $true
                $selectedsiteServers = $tempallServers.properties.name -join ", "
                if ($tempallServers.count -ge 1)
                {
                    Write-Verbose -Message "Function found new servers in site $site Adding servers: $($selectedsiteServers -join ", ")"
                    $Servers += $tempallServers
                }
                #endregion Retry Harvest exchange servers
            }

            #region Final attempt Harvest exchange servers

            if ($Servers.count -eq 0)
            {
                Write-output -InputObject "Function did resolve 0 servers in adjecent sites to ADsite $adsitename using filter: $Filterexservers "
                Write-output -InputObject "Function last attempt: Any site , any version"
                $Filterexservers = "(&(serverrole=*)(objectclass=msExchExchangeServer))"
                [array]$Servers +=  Get-Ldapobject -LDAPfilter $Filterexservers -configurationNamingContext -configurationNamingContextdomain $domain -Findall $true
                if ($Servers.count -eq 0)
                {
                    Write-error -Message "Function giving up, are you kidding me do you even have exchange installed. Please contact me info@tech-savvy.nl"
                    break
                }                    
            }

            #endregion Final attempt Harvest exchange servers
        }

        Write-Verbose "Function found the following exchange servers to try to connect to: $($Servers.properties.name -join ", ")"
    }

         

    do
    {
        try
        {
            if (!($exchangeurl))
            { 
                if (!([string]::IsNullOrWhiteSpace($Servers)))
                {
                    Write-Verbose -Message "The following servers have been found $($servers.properties.name)" 
                    $server = get-random $servers
                }
                else
                {
                    write-output -InputObject "There are 0 exchange servers of version $version $tempversion in the site $adsite"
                    throw
                }
                $ip = ($server.properties.networkaddress | ?{$_ -like "ncacn_ip_tcp*" }).split(":")[1]
                $serverconnection = "http://$ip/powershell"

            }else
            {
                $serverconnection = $exchangeurl
            }

            if ([string]::IsNullOrWhiteSpace($creds.UserName))
            {
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -ErrorAction STOP
            }elseif ([string]::IsNullOrWhiteSpace($authenticationtype))
            {
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -Credential $creds -ErrorAction STOP
            }else
            {
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $serverconnection -Credential $creds -Authentication $authenticationtype -ErrorAction STOP
            }

            if ([string]::IsNullOrWhiteSpace($prefix))
            {
                Import-PSSession $Session 4>&1 3>&1 | out-null
            }else
            {
                Import-PSSession $Session -prefix $prefix 4>&1 3>&1 | out-null
            }

            write-output -InputObject "Connected and imported session from server: $serverconnection"

        }catch [System.Management.Automation.Remoting.PSRemotingTransportException]
        {
            write-output (" tried connecting to $serverconnection but could not connect ")    
            if ($_.exception.message -like "*Access is denied*" -or $_.exception.message -like "*Access denied*")
            {
                Write-error "Error : Access is denied. invalid or no credentials. Provide credentials"
            }

            $connectionerrorcount ++
            if ( $connectionerrorcount -ge 2 )
            {
                $session = "failed"
                Write-Verbose -Message "$($_.exception.message)"
                write-error (" tried connecting 2 times but could not connect due to invalid credentials. last exchange server we tried : $serverconnection")     
            }
        }catch
        {
            $connectionerrorcount ++
            if ( $_.exception.Message -like "No command proxies have been created*")
            {
                Write-Verbose -Message "$($_.exception.message)"
                write-warning -message "No command proxies have been created, because all of the requested remote commands would shadow existing local commands."
                write-output -InputObject "Connected and imported new commands from server: $serverconnection"

            }elseif ( $_.exception.Message -like "*The attribute cannot be added because variable*")
            {
                #Catch Powershell Bug object validation: http://stackoverflow.com/questions/19775779/powershell-getnewclosure-and-cmdlets-with-validation
                Write-Verbose -Message "Powershell Bug object validation plz ignore bug rapport is created at microsoft: $($_.exception.message)"
            }else
            {
                Write-Verbose -Message "$($_.exception.message)"
                write-output -InputObject (" tried connecting to $serverconnection but could not connect " + $_.exception.Message)  
            }
            if ( $connectionerrorcount -ge 5 )
            {
                $session = "failed"
                Write-Verbose -Message "$($_.exception.message)"
                write-error (" tried connecting 5 times but could not connect. last exchange server we tried : $serverconnection")     
            }
        }
        finally
        {
            if ( $ADsite -is [System.IDisposable])
            { 
                $ADsite.Dispose()
            }
            if ( $domain -is [System.IDisposable])
            { 
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


###############
# Main script #
###############

Connect-ExchangeService