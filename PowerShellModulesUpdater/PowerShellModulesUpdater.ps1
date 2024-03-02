#Requires -Version 5.1
<#
    .NAME
        PowerShellModulesUpdater.ps1

    .SYNOPSIS
        Updates, installs and removes PowerShell modules on your system based on settings in the "Settings & Variables" region.

    .DESCRIPTION
        Updates, installs and removes PowerShell modules on your system based on settings in the "Settings & Variables" region.

        # How to use:
        ## 1. Remember to allow script execution!
        Set-ExecutionPolicy -ExecutionPolicy 'Unrestricted' -Scope 'Process' -Force -Confirm:$false
        ## 2. Some modules requires you to accept a license. This can be passed as a parameter to "Install-Module".
        Set $AcceptLicenses to $true (default) if you want these modules to be installed as well.
        Some modules that requires $AcceptLicenses to be $true:
        * Az.ApplicationMonitor

    .NOTES
        Author:   Olav Rønnestad Birkeland | github.com/o-l-a-v
        Created:  190310
        Modified: 240302

    .EXAMPLE
        # Run from PowerShell ISE or Visual Studio Code, user context
        Set-ExecutionPolicy -Scope 'Process' -ExecutionPolicy 'Bypass' -Force
        & $(Try{$psEditor.GetEditorContext().CurrentFile.Path}Catch{$psISE.CurrentFile.FullPath}) -SystemContext $false

    .EXAMPLE
        # Run from PowerShell ISE, system context, bypass script execution policy
        Set-ExecutionPolicy -Scope 'Process' -ExecutionPolicy 'Bypass' -Force
        & $(Try{$psEditor.GetEditorContext().CurrentFile.Path}Catch{$psISE.CurrentFile.FullPath}) -SystemContext $true
#>



# Input parameters
[OutputType([System.Void])]
Param (
    [Parameter(HelpMessage = 'System/device context, else current user only.')]
    [bool] $SystemContext = $false,

    [Parameter(HelpMessage = 'Whether to accept licenses when installing modules that requires it.')]
    [bool] $AcceptLicenses = $true,

    [Parameter(HelpMessage = 'Security, whether to skip checking signing of the module against alleged publisher.')]
    [bool] $SkipPublisherCheck = $true,

    [Parameter(HelpMessage = 'Whether to install missing modules, specified in variable "$ModulesWanted".')]
    [bool] $InstallMissingModules = $true,

    [Parameter(HelpMessage = 'Whether to install missing sub modules.')]
    [bool] $InstallMissingSubModules = $true,

    [Parameter(HelpMessage = 'Whether to update outdated modules.')]
    [bool] $InstallUpdatedModules = $true,

    [Parameter(HelpMessage = 'Whether to uninstall outdated modules.')]
    [bool] $UninstallOutdatedModules = $true,

    [Parameter(HelpMessage = 'Whether to uninstall unwanted modules, specified in variable "$ModulesUnwanted".')]
    [bool] $UninstallUnwantedModules = $true,

    [Parameter(HelpMessage = 'Whether to do PowerShell scripts.')]
    [bool] $DoScripts = $true
)



#region    Settings & Variables
# List of modules
## Modules you want to install and keep installed
$ModulesWanted = [string[]]$(
    'AIPService',                             # Microsoft. Used for managing Microsoft Azure Information Protection (AIP).
    'Az',                                     # Microsoft. Used for Azure Resources. Combines and extends functionality from AzureRM and AzureRM.Netcore.
    'AzSK',                                   # Microsoft. Azure Secure DevOps Kit. https://azsk.azurewebsites.net/00a-Setup/Readme.html
    'AzViz',                                  # Prateek Singh. Used for visualizing a Azure environment.
    'AWSPowerShell.NetCore',                  # Amazon AWS
    'ConfluencePS',                           # Atlassian, for interacting with Atlassian Confluence Rest API.
    'DefenderMAPS',                           # Alex Verboon, for testing connectivity to "MAPS connection for Microsoft Windows Defender".
    'Evergreen',                              # By Aaron Parker @ Stealth Puppy. For getting URL etc. to latest version of various applications on Windows.
    'ExchangeOnlineManagement',               # Microsoft. Used for managing Exchange Online.
    'GetBIOS',                                # Damien Van Robaeys. Used for getting BIOS settings for Lenovo, Dell and HP.
    'ImportExcel',                            # dfinke.    Used for import/export to Excel.
    'Intune.USB.Creator',                     # Ben Reader @ powers-hell.com. Used to create Autopilot WinPE.
    'IntuneBackupAndRestore',                 # John Seerden. Uses "MSGraphFunctions" module to backup and restore Intune config.
    'Invokeall',                              # Santhosh Sethumadhavan. Multithread PowerShell commands.
    'JWTDetails',                             # Darren J. Robinson. Used for decoding JWT, JSON Web Tokens.
    'Mailozaurr',                             # Przemyslaw Klys. Used for various email related tasks.
    'Microsoft.Graph',                        # Microsoft. Works with PowerShell Core.
    'Microsoft.Graph.Beta',                   # Microsoft. Works with PowerShell Core.
    'Microsoft.Online.SharePoint.PowerShell', # Microsoft. Used for managing SharePoint Online.
    'Microsoft.PowerShell.ConsoleGuiTools',   # Microsoft.
    'Microsoft.PowerShell.PSResourceGet',     # Microsoft, successor to PowerShellGet and PackageManagement.
    'Microsoft.PowerShell.SecretManagement',  # Microsoft. Used for securely managing secrets.
    'Microsoft.PowerShell.SecretStore',       # Microsoft. Used for securely storing secrets locally.
    'Microsoft.PowerShell.ThreadJob',         # Microsoft. Successfor of "ThreadJob" originally by Paul Higinbotham.
    'Microsoft.RDInfra.RDPowerShell',         # Microsoft. Used for managing Windows Virtual Desktop.
    'Microsoft.WinGet.Client',                # Microsoft.
    'MicrosoftGraphSecurity',                 # Microsoft. Used for interacting with Microsoft Graph Security API.
    'MicrosoftPowerBIMgmt',                   # Microsoft. Used for managing Power BI.
    'MicrosoftTeams',                         # Microsoft. Used for managing Microsoft Teams.
    'MSAL.PS',                                # Microsoft. Used for MSAL authentication.
    'MSGraphFunctions',                       # John Seerden. Wrapper for Microsoft Graph Rest API.
    'MSOnline',                               # (DEPRECATED, "AzureAD" is it's successor) Microsoft. Used for managing Microsoft Cloud Objects (Users, Groups, Devices, Domains...)
    'Nevergreen',                             # Dan Gough. Evergreen alternative that scrapes websites for getting latest version and URL to a package.
    'newtonsoft.json',                        # Serialize/Deserialize Json using Newtonsoft.json
    'Office365DnsChecker',                    # Colin Cogle. Checks a domain's Office 365 DNS records for correctness.
    'Optimized.Mga',                          # Bas Wijdenes. Microsoft Graph batch operations.
    'PartnerCenter',                          # Microsoft. Used for interacting with PartnerCenter API.
    'platyPS',                                # Microsoft. Used for converting markdown to PowerShell XML external help files.
    'PnP.PowerShell',                         # Microsoft. Used for managing SharePoint Online.
    'PolicyFileEditor',                       # Microsoft. Used for local group policy / gpedit.msc.
    'PoshRSJob',                              # Boe Prox. Used for parallel execution of PowerShell.
    'powershell-yaml'                         # Cloudbase. Serialize and deserialize YAML, using YamlDotNet.
    'PSIntuneAuth',                           # Nickolaj Andersen. Get auth token to Intune.
    'PSPackageProject',                       # Microsoft. Help with building and publishing PowerShell packages.
    'PSPKI',                                  # Vadims Podans. Used for infrastructure and certificate management.
    'PSReadLine',                             # Microsoft. Used for helping when scripting PowerShell.
    'PSScriptAnalyzer',                       # Microsoft. Used to analyze PowerShell scripts to look for common mistakes + give advice.
    'PSWindowsUpdate',                        # Michal Gajda. Used for updating Windows.
    'RunAsUser',                              # Kelvin Tegelaar. Allows running as current user while running as SYSTEM using impersonation.
    'SetBIOS',                                # Damien Van Robaeys. Used for setting BIOS settings for Lenovo, Dell and HP.
    'SharePointPnPPowerShellOnline',          # Microsoft. Used for managing SharePoint Online.
    'SpeculationControl',                     # Microsoft, by Matt Miller. To query speculation control settings (Meltdown, Spectr).
    'VSTeam',                                 # Donovan Brown. Adds functionality for working with Azure DevOps and Team Foundation Server.
    'WindowsAutoPilotIntune'                  # Michael Niehaus @ Microsoft. Used for Intune AutoPilot stuff.
)

## Modules you don't want - Will Remove Every Related Module, for AzureRM for instance will also search for AzureRM.*
$ModulesUnwanted = [string[]]$(
    'AnyPackage',                             # AnyPackage / Thomas Nieto. Spiritual successor to OneGet / PackageManagement.
    'AnyPackage.PSResourceGet'                # AnyPackage / Thomas Nieto. PSResourceGet for AnyPackage.
    'Az.Insights',                            # Name changed to "Az.Monitor": https://learn.microsoft.com/en-us/powershell/azure/migrate-az-1.0.0#module-name-changes
    'Az.Profile',                             # Name changed to "Az.Accounts": https://learn.microsoft.com/en-us/powershell/azure/migrate-az-1.0.0#module-name-changes
    'Az.Tags',                                # Functionality merged into "Az.Resources": https://learn.microsoft.com/en-us/powershell/azure/migrate-az-1.0.0#module-name-changes
    'Azure',                                  # (DEPRECATED, "Az" is its' successor) Microsoft. Used for managing Classic Azure resources/ objects.
    'AzureAD',                                # (DEPRECATED, "Microsoft.Graph" is its' successor) Microsoft. Used for managing Azure Active Directory resources/ objects.
    'AzureADPreview',                         # (DEPRECATED, "Microsoft.Graph" is its' successor) -^-
    'AzureAutomationAuthoringToolkit',        # Microsoft, Azure Automation Account add-on for PowerShell ISE.
    'AzureRM',                                # (DEPRECATED, "Az" is its' successor). Microsoft. Used for managing Azure Resource Manager resources/ objects.
    'ISESteroids',                            # Power The Shell, ISE Steroids. Used for extending PowerShell ISE functionality.
    'PackageManagement',                      # (REPLACED by "Microsoft.PowerShell.PSResourceGet"). Microsoft. Used for installing/ uninstalling modules.
    'PartnerCenterModule',                    # (DEPRECATED, "PartnerCenter" is it's successor). Microsoft. Used for interacting with Partner Center API.
    'Microsoft.Graph.Intune',                 # (REPLACED by "Replaced by Microsoft.Graph.Device"). Microsoft, old Microsoft Graph module for Intune.
    'PowerShellGet',                          # Microsoft. Used for installing updates.
    'ThreadJob',                              # (REPLACED, "Microsoft.PowerShell.ThreadJob" by Microsoft) Paul Higinbotham, running multiple processes using threads.
    'VMware'                                  # VMware, to run commands against VMware vSphere.
)

## Modules you don't want to get updated - Will not update named modules in this list
$ModulesDontUpdate = [string[]]$(
    ''
)

## Module versions you don't want removed
$ModulesVersionsDontRemove = [ordered]@{
    #'Az.Resources' = [System.Version[]] '2.3.0'
}



# List of wanted scripts
$ScriptsWanted = [string[]](
    'Get-WindowsAutoPilotInfo',                # Microsoft, Michael Niehaus. Get Windows AutoPilot info.
    'Get-AutopilotDiagnostics',                # Microsoft, Michael Niehaus. Display diagnostics information.
    'Upload-WindowsAutopilotDeviceInfo'        # Nickolaj Andersen. Upload autopilot hash straight to Intune.
)



# Settings - PowerShell Output Streams
## Set
$ConfirmPreference = 'High'
$DebugPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$WarningPreference = 'Continue'

## Fix for Progress if run using F8 on the example line in script header
$Global:ProgressPreference = $ProgressPreference

## Indentation because tabulation ("`t") is trated differently in PowerShell ISE than the PowerShell terminal
$null = Set-Variable -Option 'AllScope', 'ReadOnly' -Force -Name 'Indentation' -Value '  '
#endregion Settings & Variables




#region    Functions
#region    Get-ModuleInstalledVersions
function Get-ModuleInstalledVersions {
    <#
        .SYNOPSIS
            Gets all installed versions of a module.

        .DESCRIPTION
            Gets all installed versions of a module.
            * Includes workaround to handle versions that can't be parsed as [System.Version]

        .PARAMETER ModuleName
            String, name of the module you want to check.
    #>
    [CmdletBinding(SupportsPaging = $false)]
    [OutputType([System.Version[]])]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $ModuleName
    )

    # Begin
    Begin {}

    # Process
    Process {
        # Assets
        $Path = [string] [System.IO.Path]::Combine($Script:ModulesPath,$ModuleName)

        # Return installed versions of $ModuleName
        [System.Version[]](
            [System.IO.Directory]::GetDirectories($Path).Where{
                [System.IO.File]::Exists([System.IO.Path]::Combine($_,'PSGetModuleInfo.xml'))
            }.ForEach{
                $_.Split([System.IO.Path]::DirectorySeparatorChar)[-1]
            }.Where{
                Try {
                    $null = $_ -as [System.Version]
                    $?
                }
                Catch {
                    $false
                }
            }
        )
    }

    # End
    End {
    }
}
#endregion Get-ModuleInstalledVersions


#region    Get-PowerShellGalleryPackageLatestVersion
function Get-PowerShellGalleryPackageLatestVersion {
    <#
        .SYNAPSIS
            Fetches latest version number of a given pacakge (module or script) from PowerShellGallery.

        .PARAMETER ModuleName
            String, name of the package (module or script) you want to check.
    #>

    # Input parameters
    [CmdletBinding(SupportsPaging = $false)]
    [OutputType([System.Version])]
    param(
        [Parameter(Mandatory, HelpMessage = 'Name of the PowerShellGallery package (module or script) to fetch.')]
        [ValidateNotNullOrEmpty()]
        [Alias('ModuleName','ScriptName')]
        [string] $PackageName
    )

    # Begin
    Begin {
        if (-not$(Try{$null = Get-Variable -Name 'Random' -Scope 'Script' -ErrorAction 'SilentlyContinue'; $?}Catch{$false})) {
            $null = Set-Variable -Name 'Random' -Scope 'Script' -Option 'ReadOnly' -Force -Value (
                [random]::new()
            )
        }
    }

    # Process
    Process {
        # Access the main module page, and add a random number to trick proxies
        $Url = [string]('https://www.powershellgallery.com/packages/{0}/?dummy={1}' -f ($PackageName,$Script:Random.Next(9999)))
        Write-Debug -Message ('URL for module "{0}" = "{1}".' -f ($PackageName,$Url))


        # Create Request Url
        $Request = [System.Net.WebRequest]::Create($Url)


        # Do not allow to redirect. The result is a "MovedPermanently"
        $Request.'AllowAutoRedirect' = $false
        $Request.'Proxy' = $null


        # Try to get published version number
        Try {
            # Send the request
            $Response = $Request.GetResponse()

            # Get back the URL of the true destination page, and split off the version
            $Version = [System.Version]$($Response.GetResponseHeader('Location').Split('/')[-1])
        }
        Catch {
            # Write warning if it failed & return blank version number.
            Write-Warning -Message ($_.'Exception'.'Message')
            $Version = [System.Version]('0.0.0.0')
        }
        Finally {
            # Make sure to clean up connection
            $Response.Close()
            $Response.Dispose()
        }


        # Return Version
        return $Version
    }

    # End
    End {
    }
}
#endregion Get-PowerShellGalleryPackageLatestVersion



#region    Modules
#region    Get-ModulesInstalled
function Get-ModulesInstalled {
    <#
        .SYNAPSIS
            Gets all currently installed modules.

        .NOTES
            * Includes some workarounds to handle versions that can't be parsed as [System.Version], like beta/ pre-release.
    #>

    # Input parameters
    [CmdletBinding(SupportsPaging = $false)]
    [OutputType([System.Void])]
    Param(
        [Parameter()]
        [switch] $AllVersions,

        [Parameter()]
        [switch] $ForceRefresh,

        [Parameter()]
        [switch] $IncludePreReleaseVersions
    )

    # Begin
    Begin {
    }

    # Process
    Process {
        # Check if variable needs refresh
        if (
            $ForceRefresh -or
            ($Script:ModulesInstalled.'Count' -lt 0 -and [System.IO.Directory]::GetDirectories($Script:ModulesPath).'Count' -gt 0) -or
            $Script:ModulesInstalledNeedsRefresh -or
            -not [bool]$($null = Get-Variable -Name 'ModulesInstalledNeedsRefresh' -Scope 'Script' -ErrorAction 'SilentlyContinue'; $?)
        ) {
            # Write information
            Write-Information -MessageData 'Refreshing list of installed modules.'

            # Get installed modules given the scope
            $InstalledModulesGivenScope = [PSCustomObject[]](
                $(
                    Try {
                        Microsoft.PowerShell.PSResourceGet\Get-InstalledPSResource -Path $Script:ModulesPath | `
                            Where-Object -FilterScript {
                            $_.'Type' -eq 'Module' -and
                            $_.'Repository' -eq 'PSGallery' -and
                            -not ($IncludePreReleaseVersions -and $_.'IsPrerelease')
                        } | `
                            ForEach-Object -Process {Add-Member -InputObject $_ -MemberType 'AliasProperty' -Name 'Path' -Value 'InstalledLocation' -PassThru} | `
                            Group-Object -Property 'Name' | `
                            Select-Object -Property 'Name',@{'Name'='Versions';'Expression'='Group'}
                    }
                    Catch {
                    }
                )
            )

            # Set variable
            $null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'ModulesInstalled' -Value (
                $(
                    if ($AllVersions) {
                        $InstalledModulesGivenScope
                    }
                    else {
                        $InstalledModulesGivenScope.ForEach{
                            $_.'Versions' | Sort-Object -Property 'Version' | Select-Object -Last 1 -Property 'Name','Version','Author','Path'
                        }
                    }
                )
            )

            # Reset Script Scrope Variable "ModulesInstalledNeedsRefresh" to $false
            $null = Set-Variable -Scope 'Script' -Option 'None' -Force -Name 'ModulesInstalledNeedsRefresh' -Value ([bool]$false)
        }
    }

    # End
    End {
    }
}
#endregion Get-ModulesInstalled



#region    Update-ModulesInstalled
function Update-ModulesInstalled {
    <#
        .SYNAPSIS
            Fetches latest version number of a given module from PowerShellGallery.
    #>

    # Input parameters
    [CmdletBinding(SupportsPaging = $false)]
    [OutputType([System.Void])]
    Param()

    # Begin
    Begin {
    }

    # Process
    Process {
        # Refresh Installed Modules variable
        $null = Get-ModulesInstalled -ForceRefresh

        # Skip if no installed modules was found
        if ($Script:ModulesInstalled.'Count' -le 0) {
            Write-Information -MessageData ('No installed modules where found, no modules to update.')
            return
        }

        # Help Variables
        $C = [uint16] 1
        $CTotal = [string] $Script:ModulesInstalled.'Count'
        $Digits = [string] '0' * $CTotal.'Length'
        $ModulesInstalledNames = [string[]]($Script:ModulesInstalled.'Name' | Sort-Object)

        # Update Modules
        :ForEachModule foreach ($ModuleName in $ModulesInstalledNames) {
            # Get Latest Available Version
            $VersionAvailable = [System.Version]$(Get-PowerShellGalleryPackageLatestVersion -PackageName $ModuleName)

            # Get Version Installed
            $VersionInstalled = [System.Version]$($Script:ModulesInstalled.Where{$_.'Name' -eq $ModuleName}.'Version')

            # Get Version Installed - Get fresh version number if newer version is available and current module is a sub module
            if (
                [System.Version]($VersionAvailable) -gt [System.Version]$($VersionInstalled) -and
                $ModuleName -like '*?.?*' -and
                [string[]]$($ModulesInstalledNames) -contains [string]$($ModuleName.Replace(('.{0}' -f ($ModuleName.Split('.')[-1])),''))
            ) {
                $VersionInstalled = [System.Version](Get-ModuleInstalledVersions -ModuleName $ModuleName | Sort-Object -Descending | Select-Object -First 1)
            }

            # Present Current Module
            Write-Information -MessageData (
                '{0}/{1} {2} v{3}' -f (
                    ($C++).ToString($Digits),
                    $CTotal,
                    $ModuleName,
                    $VersionInstalled.ToString()
                )
            )

            # Compare Version Installed vs Version Available
            if ([System.Version]$($VersionInstalled) -ge [System.Version]$($VersionAvailable)) {
                Write-Information -MessageData ('{0}Current version is latest version.' -f $Indentation)
                Continue ForEachModule
            }
            else {
                Write-Information -MessageData ('{0}Newer version available, v{1}.' -f $Indentation, $VersionAvailable.ToString())
                if ([bool]$($null = Microsoft.PowerShell.PSResourceGet\Find-PSResource -Type 'Module' -Repository 'PSGallery' -Name $ModuleName -ErrorAction 'SilentlyContinue';$?)) {
                    if ($ModulesDontUpdate -contains $ModuleName) {
                        Write-Information -MessageData ('{0}Will not update as module is specified in script settings. ($ModulesDontUpdate).' -f ($Indentation*2), $Success.ToString)
                    }
                    else {
                        # Install module
                        $Success = [bool]$(
                            Try {
                                $null = Microsoft.PowerShell.PSResourceGet\Save-PSResource -Repository 'PSGallery' -TrustRepository -IncludeXml `
                                    -Path $Script:ModulesPath -Name $ModuleName -SkipDependencyCheck -Confirm:$false -Verbose:$false 2>$null
                                $?
                            }
                            Catch {
                                $false
                            }
                        )

                        # Double check for success
                        if ($Success) {
                            $Success = [bool](
                                [System.IO.Directory]::Exists(
                                    [System.IO.Path]::Combine($Script:ModulesPath, $ModuleName, $VersionAvailable.ToString())
                                )
                            )
                        }

                        # Output success
                        Write-Information -MessageData ('{0}Install success? {1}' -f ($Indentation * 2), $Success.ToString())

                        # Count as installed if success
                        if ($Success) {
                            # Stats
                            $Script:ModulesUpdated += [string[]]$($ModuleName)
                            # Updated cache of installed modules and version if current module has sub modules
                            if ([uint16]$([string[]]$($ModulesInstalledNames.Where{$_ -like ('{0}.?*' -f $Module)}).'Count') -ge 1) {
                                Get-ModulesInstalled
                            }
                            # Else, set flag to update cache of installed modules later
                            else {
                                $Script:ModulesInstalledNeedsRefresh = [bool] $true
                            }
                        }
                        else {
                            Write-Warning -Message (
                                'Failed to install "{0}". Here is the error:{1}"{2}"' -f (
                                    $ModuleName,
                                    [System.Environment]::NewLine,
                                    $Error[0].ToString()
                                )
                            )
                        }
                    }
                }
                else {
                    Write-Warning -Message ('Did not find "{0}" in PSGallery, probably deprecated, delisted or something similar. Will skip.' -f ($ModuleName))
                }
            }
        }
    }

    # End
    End {
    }
}
#endregion Update-ModulesInstalled


#region    Install-ModulesMissing
function Install-ModulesMissing {
    <#
        .SYNAPSIS
            Installs missing modules by comparing installed modules vs input parameter $ModulesWanted.

        .PARAMETER ModulesWanted
            A string array containing names of wanted modules.
    #>

    # Input parameters
    [CmdletBinding(SupportsPaging = $false)]
    [OutputType([System.Void])]
    Param(
        [Parameter(Mandatory)]
        [string[]] $ModulesWanted
    )

    # Begin
    Begin {}

    # Process
    Process {
        # Refresh Installed Modules variable
        $null = Get-ModulesInstalled


        # Help Variables
        $C = [uint16](1)
        $CTotal = [string]($ModulesWanted.'Count')
        $Digits = [string]('0' * $CTotal.'Length')
        $ModulesInstalledNames = [string[]]($Script:ModulesInstalled.'Name' | Sort-Object)


        # Loop each wanted module. If not found in installed modules: Install it
        foreach ($ModuleName in $ModulesWanted) {
            Write-Information -MessageData ('{0}/{1} {2}' -f (($C++).ToString($Digits),$CTotal,$ModuleName))

            # Install if not already installed
            if ($ModulesInstalledNames.Where{$_ -eq $ModuleName}.'Count' -ge 1) {
                Write-Information -MessageData ('{0}Already Installed. Next!' -f $Indentation)
            }
            else {
                Write-Information -MessageData ('{0}Not already installed. Installing.' -f $Indentation)
                $PSResource = Microsoft.PowerShell.PSResourceGet\Find-PSResource -Type 'Module' -Repository 'PSGallery' -Name $ModuleName -ErrorAction 'SilentlyContinue'
                if ($? -and -not [string]::IsNullOrEmpty($PSresource.'Name')) {
                    # Special case if Unix
                    if (
                        $PSVersionTable.'Platform' -eq 'Unix' -and
                        $PSResource.'Tags'.Where({$_ -in 'Linux','Mac','MacOS','PSEdition_Core'},'First').'Count' -le 0
                    ) {
                        Write-Information -MessageData ('{0}This module does not seem to support Unix, thus skipping it.' -f ($Indentation*2))
                        Continue
                    }
                    # Install missing module
                    ## Assets
                    $ModulesToInstall = [string[]]($ModuleName)
                    ## Find dependencies
                    $ModulesToInstall += [string[]](
                        $PSResource.'Dependencies'.'Name'.Where{
                            $_.StartsWith('{0}.' -f $ModuleName)
                        }
                    )
                    ## Remove empty results and sort object
                    $ModulesToInstall = [string[]]($ModulesToInstall.Where{-not [string]::IsNullOrEmpty($_)} | Sort-Object -Unique)
                    ## Install $ModuleName and all $Dependencies with -SkipDependencyCheck
                    $Success = [bool]$(
                        Try {
                            $null = Microsoft.PowerShell.PSResourceGet\Save-PSResource -Repository 'PSGallery' -TrustRepository -IncludeXml `
                                -Path $Script:ModulesPath -Name $ModulesToInstall -SkipDependencyCheck -Confirm:$false -Verbose:$false 2>$null
                            $?
                        }
                        Catch {
                            $false
                        }
                    )

                    # Output success
                    Write-Information -MessageData ('{0}Install success? {1}' -f ($Indentation * 2),$Success.ToString())

                    # Count if success
                    if ($Success) {
                        # Stats
                        $Script:ModulesInstalledMissing += [string[]]$($ModuleName)
                        # Make sure list of installed modules gets refreshed
                        $Script:ModulesInstalledNeedsRefresh = $true
                    }
                    else {
                        Write-Warning -Message ('Failed to install "{0}". Here is the error:{1}"{2}"' -f ($ModuleName,[System.Environment]::NewLine,$Error[0].ToString()))
                    }
                }
                else {
                    Write-Warning -Message ('Did not find "{0}" in PSGallery, probably deprecated, delisted or something similar. Will skip.' -f ($ModuleName))
                }
            }
        }
    }

    # End
    End {
    }
}
#endregion Install-ModulesMissing


#region    Install-SubModulesMissing
function Install-SubModulesMissing {
    <#
        .SYNAPSIS
            Installs eventually missing submodules
    #>
    # Input parameters
    [CmdletBinding(SupportsPaging = $false)]
    [OutputType([System.Void])]
    Param()

    # Begin
    Begin {}

    # Process
    Process {
        # Refresh Installed Modules variable
        $null = Get-ModulesInstalled

        # Skip if no installed modules was found
        if ($Script:ModulesInstalled.'Count' -le 0) {
            Write-Information -MessageData ('No installed modules where found, no modules to check against.')
            return
        }

        # Help Variables - Both Foreach
        $ParentModulesInstalled = [array](
            $ModulesInstalled.Where{
                [bool]($_.'Name' -in $ModulesWanted) -or
                [bool]($_.'Name' -notlike '*.*') -or
                [bool](
                    $_.'Name' -like '*.*' -and
                    $_.'Name' -notlike '*.*.*' -and
                    [string[]]$($ModulesInstalled.'Name') -notcontains [string]$($_.'Name'.Replace(('.{0}' -f ($_.'Name'.Split('.')[-1])),''))
                )
            } | Select-Object -Property 'Name','Author' | Sort-Object -Property 'Name' -Unique
        )

        # Help Variables - Outer Foreach
        $OC = [uint16] 1
        $OCTotal = [string] $ParentModulesInstalled.'Count'.ToString()
        $ODigits = [string]('0' * $OCTotal.'Length')

        # Loop Through All Installed Modules
        :ForEachModule foreach ($Module in $ParentModulesInstalled) {
            # Present Current Module
            Write-Information -MessageData ('{0}/{1} {2} by "{3}"' -f ($OC++).ToString($ODigits), $OCTotal, $Module.'Name', $Module.'Author')

            # Get all installed sub modules
            $SubModulesInstalled = [string[]](
                $ModulesInstalled.Where{
                    $_.'Name' -like ('{0}.*' -f $Module.'Name') -and $_.'Author' -eq $Module.'Author'
                }.'Name' | Sort-Object
            )

            # Get all available sub modules
            $SubModulesAvailable = [PSCustomObject[]](
                [array](
                    Microsoft.PowerShell.PSResourceGet\Find-PSResource -Type 'Module' -Repository 'PSGallery' -Name ('{0}.*' -f $Module.'Name') | `
                        Where-Object -FilterScript {
                        $_.'Author' -eq $Module.'Author' -and
                        $_.'Name' -notin $ModulesUnwanted
                    } | Select-Object -Property 'Name','Version' | Sort-Object -Unique -Property 'Name'
                )
            )

            # If either $SubModulesAvailable is 0, Continue Outer Foreach
            if ($SubModulesAvailable.'Count' -eq 0) {
                Write-Information -MessageData (
                    '{0}Found {1} available sub module{2}.' -f (
                        $Indentation,
                        $SubModulesAvailable.'Count'.ToString(),
                        [string]$(if ($SubModulesAvailable.'Count' -ne 1){'s'})
                    )
                )
                Continue ForEachModule
            }

            # Compare objects to see which are missing
            $SubModulesMissing = [PSCustomObject[]]$(
                if ($SubModulesInstalled.'Count' -eq 0) {
                    $SubModulesAvailable
                }
                else {
                    $SubModulesAvailable.Where{
                        $_.'Name' -notin $SubModulesInstalled
                    }
                }
            )
            ## Output result
            Write-Information -MessageData (
                '{0}Found {1} missing sub module{2}.' -f (
                    $Indentation,
                    $SubModulesMissing.'Count'.ToString(),
                    [string]$(if ($SubModulesMissing.'Count' -ne 1){'s'})
                )
            )

            # Install missing sub modules
            if ($SubModulesMissing.'Count' -gt 0) {
                # Help Variables - Inner Foreach
                $IC = [uint16] 1
                $ICTotal = [string] $SubModulesMissing.'Count'.ToString()
                $IDigits = [string]('0' * $ICTotal.'Length')

                # Install Modules
                :ForEachSubModule foreach ($SubModule in $SubModulesMissing) {
                    # Present Current Sub Module
                    Write-Information -MessageData (
                        '{0}{1}/{2} {3}' -f (
                            ($Indentation * 2),
                            ($IC++).ToString($IDigits),
                            $ICTotal,
                            $SubModule.'Name'
                        )
                    )


                    # Install The Missing Sub Module
                    $Success = [bool]$(
                        Try {
                            $null = Microsoft.PowerShell.PSResourceGet\Save-PSResource -Repository 'PSGallery' -TrustRepository -IncludeXml `
                                -Path $Script:ModulesPath -Name $SubModule.'Name' -SkipDependencyCheck -Confirm:$false -Verbose:$false 2>$null
                            $?
                        }
                        Catch {
                            $false
                        }
                    )


                    # Double check for success
                    if ($Success) {
                        $Success = [bool](
                            [System.IO.Directory]::Exists(
                                [System.IO.Path]::Combine($Script:ModulesPath, $SubModule.'Name', $SubModule.'Version'.ToString())
                            )
                        )
                    }

                    # Output success
                    Write-Information -MessageData ('{0}Install success? {1}' -f ($Indentation * 3), $Success.ToString())

                    # Count as installed if success
                    if ($Success) {
                        # Stats
                        $Script:ModulesSubInstalledMissing += [string[]]$($SubModule.'Name')
                        # Make sure list of installed modules gets refreshed
                        $Script:ModulesInstalledNeedsRefresh = $true
                    }
                    else {
                        Write-Warning -Message ('Failed to install "{0}". Here is the error:{1}"{2}"' -f $SubModule.'Name', [System.Environment]::NewLine, $Error[0].ToString())
                    }
                }
            }
        }
    }

    # End
    End {
    }
}
#endregion Install-SubModulesMissing


#region    Uninstall-ModuleManually
function Uninstall-ModuleManually {
    <#
        .SYNOPSIS
            Uninstalls a module version by trying to find the folder, then delete it. If that fails "Remove-Package" is called instead.

        .DESCRIPTION
            Uninstalls a module version by trying to find the folder, then delete it. If that fails "Remove-Package" is called instead.
                * Created because "Uninstall-Module" and "Remove-Package" is very slow.
    #>
    [CmdletBinding()]
    [OutputType([System.Void])]
    Param (
        # Mandatory
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $ModuleName,

        [Parameter(Mandatory)]
        [System.Version] $Version
    )

    # Begin
    Begin {
        # Check if running as admin if $SystemContext
        if ($Script:SystemContext -and -not $([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Throw 'ERROR - Admin permissions are required if running in system context.'
        }
    }

    # Process
    Process {
        # Create path to module
        $ModulePath = [string] [System.IO.Path]::Combine($Script:ModulesPath, $ModuleName, $Version.ToString())

        # If path does not exist, try to check for versions and create path more dynamically
        if (-not [System.IO.Directory]::Exists($ModulePath)) {
            # Create module path
            $ModulePath = [string] [System.IO.Path]::Combine($Script:ModulesPath, $ModuleName)

            # Get all versions
            $Versions = [PSCustomObject[]](
                $([array](Get-ChildItem -Path $ModulePath -Directory -Depth 0)).ForEach{
                    Try {
                        [PSCustomObject]@{
                            'Path'    = [string] $_.'FullName'
                            'Version' = [System.Version] $_.'Name'
                        }
                    }
                    Catch {
                    }
                }
            )

            # Cut ending 0 if more version numbers than input $Version
            $Versions.ForEach{
                if ($_.'Version'.ToString().Split('.').'Count' -gt $Version.ToString().Split('.').'Count' -and $([uint16]$_.'Version'.ToString().Split('.')[-1]) -eq 0) {
                    $_.'Version' = [System.Version][string]($_.'Version'.ToString().Split('.')[0..$($([byte]($Version.ToString().Split('.').'Count')) - 1)] -join '.')
                }
            }

            # Set path
            if ([System.Version[]]($Versions.'Version') -contains $Version) {
                $ModulePath = [string]($Versions.Where{$_.'Version' -eq $Version}.'Path')
            }
        }

        # Unload module if currently in use
        if (
            $(Microsoft.PowerShell.Core\Get-Module -Name $ModuleName).Where{
                $_.'Path'.StartsWith(($ModulePath + [System.IO.Path]::DirectorySeparatorChar))
            }.'Count' -gt 0
        ) {
            $null = Remove-Module -Name $ModuleName -Force
        }

        # Delete folder if it exists / Uninstall module version
        if (
            [bool]$(
                Try{$null = [System.Version]$ModulePath.Split([System.IO.Path]::DirectorySeparatorChar)[-1];$?}Catch{$false}) -and
            [System.IO.Directory]::Exists($ModulePath)
        ) {
            $null = [System.IO.Directory]::Delete($ModulePath,$true)
            if ($? -and -not [System.IO.Directory]::Exists($ModulePath)) {
                return
            }
        }
    }

    # End
    End {
    }
}
#endregion Uninstall-ModuleManually


#region    Uninstall-ModulesOutdated
function Uninstall-ModulesOutdated {
    <#
        .SYNOPSIS
            Uninstalls outdated modules / currently installed modules with more than one version.

        .Description
            Uninstalls outdated modules / currently installed modules with more than one version.
            * Uses workaround from function Get-ModulesInstalled to handle version numbers that can't be parsed as [System.Version].
    #>

    # Input parameters
    [CmdletBinding(SupportsPaging = $false)]
    [OutputType([System.Void])]
    Param()

    # Begin
    Begin {
    }

    # Process
    Process {
        # Refresh Installed Modules variable
        $null = Get-ModulesInstalled -AllVersions -ForceRefresh

        # Skip if no installed modules were found
        if ($Script:ModulesInstalled.'Count' -le 0) {
            Write-Information -MessageData ('No installed modules where found, no outdated modules to uninstall.')
            Return
        }

        # Only care about modules installed with multiple versions
        $ModulesInstalledWithMultipleVersions = [PSCustomObject[]](
            $Script:ModulesInstalled.Where{$_.'Versions'.'Count' -ge 2} | Sort-Object -Property 'Name'
        )

        # Skip if no modules have multiple versions installed
        if ($ModulesInstalledWithMultipleVersions.'Count' -le 0) {
            Write-Information -MessageData ('None of the installed modules in current scope have multiple versions installed.')
            Return
        }

        # Help Variables
        $C = [uint16](1)
        $CTotal = [string]($ModulesInstalledWithMultipleVersions.'Count')
        $Digits = [string]('0' * $CTotal.'Length')

        # Get Versions of Installed Main Modules
        :ForEachModuleName foreach ($Module in $ModulesInstalledWithMultipleVersions) {
            Write-Information -MessageData ('{0}/{1} {2}' -f (($C++).ToString($Digits),$CTotal,$Module.'Name'))

            # Get all versions installed
            $VersionsAll = [System.Version[]](
                $Module.'Versions'.'Version'
            )
            Write-Information -MessageData (
                '{0}{1} got {2} installed version{3} ({4}).' -f (
                    $Indentation,
                    $Module.'Name',
                    $VersionsAll.'Count'.ToString(),
                    [string]$(if ($VersionsAll.'Count' -gt 1){'s'}),
                    [string]($VersionsAll -join ', ')
                )
            )

            # Find newest version
            $ModuleOldVersions = [array]($Module.'Versions' | Sort-Object -Property 'Version' -Descending | Select-Object -Skip 1)

            # Uninstall all but newest
            foreach ($ModuleVersion in $ModuleOldVersions) {
                Write-Information -MessageData ('{0}Uninstalling module "{1}" version "{2}".' -f ($Indentation*2), $Module.'Name', $ModuleVersion.'Version'.ToString())

                # Check if current version is not to be uninstalled / specified in $ModulesVersionsDontRemove
                if ($([System.Version[]]($ModulesVersionsDontRemove.$($Module.'Name')) -contains $([System.Version]($ModuleVersion.'Version')))) {
                    Write-Information -MessageData ('{0}This version is not to be uninstalled because it`s specified in $ModulesVersionsDontRemove.' -f ($Indentation*3))
                }
                else {
                    # Uninstall
                    $Success = [bool]$(
                        Try {
                            $null = Uninstall-ModuleManually -ModuleName $Module.'Name' -Version $ModuleVersion.'Version' `
                                -WarningAction 'SilentlyContinue' -ErrorAction 'SilentlyContinue'
                            $?
                        }
                        Catch {
                            $false
                        }
                    )

                    # Check for success
                    $Success = [bool](
                        $Success -and -not [System.IO.Directory]::Exists([System.IO.Path]::Combine($ModuleVersion.'InstalledLocation', $ModuleVersion.'Version'))
                    )

                    # Output
                    Write-Information -MessageData ('{0}Success? {1}.' -f ($Indentation*3), $Success.ToString())
                    if ($Success) {
                        # Stats
                        $Script:ModulesUninstalledOutdated += [string[]]($Module.'Name')
                        # Make sure list of installed modules gets refreshed
                        $Script:ModulesInstalledNeedsRefresh = $true
                    }
                    else {
                        # Output warning
                        Write-Warning -Message ('Failed to uninstall "{0}" v{1}' -f $Module.'Name', $Module.'Version') -WarningAction 'Continue'
                        Continue ForEachModuleName
                    }
                }
            }
        }
    }

    # End
    End {
    }
}
#endregion Uninstall-ModulesOutdated


#region    Uninstall-ModulesUnwanted
function Uninstall-ModulesUnwanted {
    <#
        .SYNAPSIS
            Uninstalls installed modules that matches any value in the input parameter $ModulesUnwanted.

        .PARAMETER ModulesUnwanted
            String Array containig modules you don't want to be installed on your system.
    #>
    [CmdletBinding(SupportsPaging = $false)]
    [OutputType([System.Void])]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ModulesUnwanted
    )

    # Begin
    Begin {
    }

    # Process
    Process {
        # Refresh Installed Modules variable
        $null = Get-ModulesInstalled -ForceRefresh

        # Skip if no installed modules was found
        if ($Script:ModulesInstalled.'Count' -le 0) {
            Write-Information -MessageData ('No installed modules where found, no modules to uninstall.')
            return
        }

        # Find out if we got unwated modules installed based on input parameter $ModulesUnwanted vs $InstalledModulesAll
        $ModulesToRemove = [string[]]$(
            :ForEachInstalledModule foreach ($ModuleInstalled in [string[]]$($Script:ModulesInstalled.'Name')) {
                :ForEachUnwantedModule foreach ($ModuleUnwanted in $ModulesUnwanted) {
                    if ($ModuleInstalled -eq $ModuleUnwanted -or $ModuleInstalled -like ('{0}.*' -f $ModuleUnwanted)) {
                        $ModuleInstalled
                        Continue ForEachInstalledModule
                    }
                }
            }
        ) | Sort-Object

        # Write Out How Many Unwanted Modules Was Found
        Write-Information -MessageData (
            'Found {0} unwanted module{1}.{2}' -f (
                $ModulesToRemove.'Count'.ToString(),
                $(if ($ModulesToRemove.'Count' -ne 1){'s'}),
                $(if ($ModulesToRemove.'Count' -gt 0){
                        ' Will proceed to uninstall {0}.' -f $(if ($ModulesToRemove.'Count' -eq 1){'it'}else{'them'})
                    })
            )
        )

        # Uninstall Unwanted Modules If More Than 0 Was Found
        if ([uint16]$($ModulesToRemove.'Count') -gt 0) {
            # Assets
            $C = [uint16]$(1)
            $CTotal = [string]$($ModulesToRemove.'Count'.ToString())
            $Digits = [string]$('0' * $CTotal.'Length')
            $LengthModuleLongestName = [byte]($([byte[]]($([string[]]($ModulesToRemove)).ForEach{$_.'Length'} | Sort-Object)) | Select-Object -Last 1)

            # Loop
            foreach ($Module in $ModulesToRemove) {
                # Write information
                Write-Information -MessageData (
                    '{0}{1}/{2} {3}{4}' -f (
                        $Indentation,
                        ($C++).ToString($Digits),
                        $CTotal,
                        $Module,
                        (' ' * [byte]$([byte]$($LengthModuleLongestName) - [byte]$($Module.'Length')))
                    )
                )

                # Remove Current Module
                $null = [System.IO.Directory]::Delete([System.IO.Path]::Combine($Script:ModulesPath, $Module),$true) 2>$null
                $Success = [bool]$($?)

                # Write Out Success
                Write-Information -MessageData ('{0}Success? {1}.' -f ($Indentation*2), $Success.ToString())
                if ($Success) {
                    # Stats
                    $Script:ModulesUninstalledUnwanted += [string[]]($Module)
                    # Make sure list of installed modules gets refreshed
                    $Script:ModulesInstalledNeedsRefresh = $true
                }
            }
        }
    }

    # End
    End {
    }
}
#endregion Uninstall-ModulesUnwanted
#endregion Modules



#region    Scripts
#region    Install-ScriptsMissing
function Install-ScriptsMissing {
    <#
        .SYNOPSIS
            Installs scripts that is not already installed.'
    #>

    # Input parameters
    [OutputType([System.Void])]
    Param(
        [Parameter(Mandatory, HelpMessage = 'String array with script names from PowerShell Gallery.')]
        [string[]] $Scripts
    )

    # Begin
    Begin {
        # Make sure script folder exists by creating "InstalledScriptInfos" if it does not already exist
        $InstalledScriptsInfos = [string] [System.IO.Path]::Combine($Script:ScriptsPath,'InstalledScriptInfos')
        if (-not [System.IO.Directory]::Exists($InstalledScriptsInfos)) {
            $null = [System.IO.Directory]::CreateDirectory($InstalledScriptsInfos)
        }

        # Get installed scripts
        $InstalledScripts = [array](
            $(
                Try {
                    [System.IO.Directory]::GetFiles([System.IO.Path]::Combine($Script:ScriptsPath,'InstalledScriptInfos'),'*.xml').ForEach{
                        [Microsoft.PowerShell.PSResourceGet.UtilClasses.TestHooks]::ReadPSGetResourceInfo(
                            $_
                        )
                    }.Where{$_.'Repository' -eq 'PSGallery' -and $_.'Type' -eq 'Script'}
                }
                Catch {
                }
            )
        )
    }

    # Process
    Process {
        foreach ($Script in $Scripts) {
            Write-Information -MessageData (
                '{0} / {1} "{2}"' -f (
                    (1 + $Scripts.IndexOf($Script)).ToString('0' * $Scripts.'Count'.ToString().'Length'),
                    $Scripts.'Count'.Tostring(),
                    $Script
                )
            )
            if ($InstalledScripts.'Name' -contains $Script) {
                Write-Information -MessageData 'Already installed.'
            }
            else {
                Write-Information -MessageData 'Not already installed.'
                $PSResource = Microsoft.PowerShell.PSResourceGet\Find-PSResource -Type 'Script' -Repository 'PSGallery' -Name $Script -ErrorAction 'SilentlyContinue'
                if ($? -and -not [string]::IsNullOrEmpty($PSResource.'Name')) {
                    # Special case if Unix
                    if (
                        $PSVersionTable.'Platform' -eq 'Unix' -and
                        $PSResource.'Tags'.Where({$_ -in 'Linux','Mac','MacOS','PSEdition_Core'},'First').'Count' -le 0
                    ) {
                        Write-Information -MessageData ('{0}This script does not seem to support Unix, thus skipping it.' -f ($Indentation*2))
                        Continue
                    }
                    # Install script
                    $null = Microsoft.PowerShell.PSResourceGet\Save-PSResource -Repository 'PSGallery' -TrustRepository `
                        -IncludeXml -SkipDependencyCheck -Path $Script:ScriptsPath -Name $Script
                    # Move XML file to "InstalledScriptInfos"
                    $Local:Source = [string] [System.IO.Path]::Combine(
                        $Script:ScriptsPath,
                        '{0}_InstalledScriptInfo.xml' -f $Script
                    )
                    $Local:Destination = [string] [System.IO.Path]::Combine(
                        $InstalledScriptsInfos,
                        '{0}_InstalledScriptInfo.xml' -f $Script
                    )
                    if ([System.IO.File]::Exists($Local:Source)) {
                        $null = Move-Item -Path $Local:Source -Destination $Local:Destination -Force
                    }
                    # Output success
                    if ([System.IO.File]::Exists($Local:Destination)) {
                        Write-Information -MessageData ('{0}Successfully installed.' -f $Indentation)
                        # Add to stats
                        $Script:ScriptsMissing += [string[]]($Script)
                    }
                    else {
                        Write-Warning -Message ('{0}Failed.' -f $Indentation)
                    }
                }
                else {
                    Write-Warning -Message 'Not found on PowerShellGallery.'
                }
            }
        }
    }

    # End
    End {
    }
}
#endregion Install-ScriptsMissing


#region    Update-ScriptsInstalled
function Update-ScriptsInstalled {
    <#
        .SYNOPSIS
            Updates installed scripts originating from PowerShell Gallery.
    #>

    # Input parameters
    [OutputType([System.Void])]
    Param()

    # Begin
    Begin {
        # Make sure script folder exists by creating "InstalledScriptInfos" if it does not already exist
        $InstalledScriptsInfos = [string] [System.IO.Path]::Combine($Script:ScriptsPath,'InstalledScriptInfos')
        if (-not [System.IO.Directory]::Exists($InstalledScriptsInfos)) {
            $null = [System.IO.Directory]::CreateDirectory($InstalledScriptsInfos)
        }

        # Get installed scripts
        $InstalledScripts = [array](
            $(
                Try {
                    [System.IO.Directory]::GetFiles([System.IO.Path]::Combine($Script:ScriptsPath,'InstalledScriptInfos'),'*.xml').ForEach{
                        [Microsoft.PowerShell.PSResourceGet.UtilClasses.TestHooks]::ReadPSGetResourceInfo(
                            $_
                        )
                    }.Where{$_.'Repository' -eq 'PSGallery' -and $_.'Type' -eq 'Script'}
                }
                Catch {
                }
            )
        )
    }

    # Process
    Process {
        $InstalledScripts.ForEach{
            # Output current script
            Write-Information -MessageData (
                '{0} / {1} "{2}" v{3} by "{4}"' -f (
                    (1 + $InstalledScripts.IndexOf($_)).ToString('0' * $InstalledScripts.'Count'.ToString().'Length'),
                    $InstalledScripts.'Count'.Tostring(),
                    $_.'Name',
                    $_.'Version'.ToString(),
                    $([string[]]($_.'Author',$_.'Entities'.Where{$_.'Role' -eq 'author'}.'Name')).Where{-not[string]::IsNullOrEmpty($_)}[0]
                )
            )
            # Find package on PowerShell Gallery
            $Package = Microsoft.PowerShell.PSResourceGet\Find-PSResource -Repository 'PSGallery' -Type 'Script' -Name $_.'Name' 2>$null
            # Skip if can't be found on PowerShell Gallery
            if ([string]::IsNullOrEmpty($Package.'Name')) {
                Write-Information -MessageData 'Not from PowerShell Gallery, or at least could not be found.'
            }
            else {
                if ($Package.'Version' -gt $_.'Version') {
                    # Install script
                    $null = Microsoft.PowerShell.PSResourceGet\Save-PSResource -Repository 'PSGallery' -TrustRepository `
                        -IncludeXml -SkipDependencyCheck -Path $Script:ScriptsPath -Name $_.'Name'
                    # Move XML file to "InstalledScriptInfos"
                    $Local:Source = [string] [System.IO.Path]::Combine(
                        $Script:ScriptsPath,
                        '{0}_InstalledScriptInfo.xml' -f $_.'Name'
                    )
                    $Local:Destination = [string] [System.IO.Path]::Combine(
                        $InstalledScriptsInfos,
                        '{0}_InstalledScriptInfo.xml' -f $_.'Name'
                    )
                    if ([System.IO.File]::Exists($Local:Source)) {
                        $null = Move-Item -Path $Local:Source -Destination $Local:Destination -Force
                    }
                    # Add to stats
                    $Script:ScriptsUpdated += [string[]]($_.'Name')
                    # Output success
                    Write-Information -MessageData ('{0}Successfully installed.' -f $Indentation)
                }
                else {
                    Write-Information -MessageData 'No newer version available.'
                }
            }
        }
    }

    # End
    End {
    }
}
#endregion Update-ScriptsInstalled
#endregion Scripts



#region    User context PSModulePath
function Set-PSModulePathUserContext {
    [OutputType([System.Void])]
    Param()

    Begin {
    }

    Process {
        # Exit function if running on Linux
        if ($PSVersionTable.'Platform' -eq 'Unix') {
            return
        }

        # Assets
        $PSModulePathWanted = [string] '%LOCALAPPDATA%\Microsoft\PowerShell\Modules'
        $PSModulePathWantedResolved = [string] [System.Environment]::ExpandEnvironmentVariables($PSModulePathWanted)
        $RegistryPath = [string] 'Registry::HKEY_CURRENT_USER\Environment'

        # Create path if it does not exist
        if (-not [System.IO.Directory]::Exists($PSModulePathWantedResolved)) {
            $null = [System.IO.Directory]::CreateDirectory($PSModulePathWantedResolved)
        }

        # Get current value without resolving the path / expanding the environmental variable
        $PSModulePathCurrent = [string](
            (Get-Item -Path $RegistryPath).GetValue(
                'PSModulePath',
                '',
                'DoNotExpandEnvironmentNames'
            )
        )

        # Make current PSModulePath to a string array for easier operations
        $PSModulePathNewAsArray = [string[]](
            $PSModulePathCurrent.Split(
                [System.IO.Path]::PathSeparator
            ).Where{
                -not [string]::IsNullOrEmpty($_)
            }
        )

        # Remove "MyDocuments" if present, as it will resolve to OneDrive if Known Folder Move is enabled
        $PSModulePathNewAsArray = [string[]](
            $PSModulePathNewAsArray.Where{
                $_ -notlike ('{0}\*' -f [System.Environment]::GetFolderPath('MyDocuments'))
            }
        )

        # Add $PSModulePathWanted if not already present
        if ($PSModulePathNewAsArray -notcontains $PSModulePathWanted) {
            $PSModulePathNewAsArray = [string[]](
                [string[]]($PSModulePathWanted) + [string[]]($PSModulePathNewAsArray) | Where-Object -FilterScript {
                    -not [string]::IsNullOrEmpty($_)
                }
            )
        }

        # Convert $PSModulePathNewAsArray to string for easier comparison to existing value
        $PSModulePathNew = [string]($PSModulePathNewAsArray -join [System.IO.Path]::PathSeparator)

        # Set new value if it changed
        if ($PSModulePathNew -ne $PSModulePathCurrent) {
            $null = Set-ItemProperty -Path $RegistryPath -Name 'PSModulePath' -Value $PSModulePathNew -Force -Type ([Microsoft.Win32.RegistryValueKind]::ExpandString)
        }
    }

    End {}
}
#endregion User context PSModulePath



#region    Write-Statistics
function Write-Statistics {
    <#
        .SYNOPSIS
            Outputs statistics after script has ran.
    #>

    # Input parameters
    [CmdletBinding()]
    [OutputType([System.Void])]
    Param ()

    # Begin
    Begin {
    }

    # Process
    Process {
        # Help variables
        $FormatDescriptionLength = [byte]$($Script:StatisticsVariables.ForEach{$_.'Description'.'Length'} | Sort-Object -Descending | Select-Object -First 1)
        $FormatNewLineTab = [string]$("`r`n`t")

        # Output stats
        foreach ($Variable in $Script:StatisticsVariables) {
            $CurrentObject = [string[]]$(Get-Variable -Name $Variable.'VariableName' -Scope 'Script' -ValueOnly -ErrorAction 'SilentlyContinue' | Sort-Object -Unique)
            $CurrentDescription = [string]$($Variable.'Description' + ':' + [string]$(' ' * [byte]$($FormatDescriptionLength - $Variable.'Description'.'Length')))
            Write-Information -MessageData (
                '{0} {1}{2}' -f (
                    $CurrentDescription,
                    $CurrentObject.'Count',
                    [string]$(if ($CurrentObject.'Count' -ge 1){$FormatNewLineTab + [string]$($CurrentObject -join $FormatNewLineTab)})
                )
            )
        }
    }

    # End
    End {
    }
}
#endregion Write-Statistics
#endregion Functions




#region    Main
# Help variables
## Start time
$null = New-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'TimeTotalStart' -Value ([datetime]::Now)


## Random number generator
$null = New-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'Random' -Value ([System.Random]::New())


## Scope
$null = New-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'Scope' -Value (
    [string]$(
        if ($SystemContext) {
            'AllUsers'
        }
        else {
            'CurrentUser'
        }
    )
)


## Modules path
### Root
$null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'PSResourceHomeUser' -Value (
    [string](
        $(
            if ($PSVersionTable.'Platform' -eq 'Unix') {
                [System.IO.Path]::Combine($HOME,'.local','share','powershell')
            }
            else {
                [System.IO.Path]::Combine($env:LOCALAPPDATA,'Microsoft','PowerShell')
            }
        )
    )
)
$null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'PSResourceHomeMachine' -Value (
    [string](
        $(
            if ($PSVersionTable.'Platform' -eq 'Unix') {
                '/usr/local/share/powershell'
            }
            else {
                [System.IO.Path]::Combine(
                    $(
                        if ([System.Environment]::Is64BitOperatingSystem) {
                            $env:ProgramW6432
                        }
                        else {
                            ${env:ProgramFiles(x86)}
                        }
                    ),
                    '{0}PowerShell' -f $(if ($PSVersionTable.'PSEdition' -eq 'Desktop') {'Windows'})
                )
            }
        )
    )
)
$null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'ModulesPathRoot' -Value (
    [string]($(if ($SystemContext){$PSResourceHomeMachine}else{$PSResourceHomeUser}))
)

### Modules path
$null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'ModulesPath' -Value (
    [string] [System.IO.Path]::Combine($Script:ModulesPathRoot,'Modules')
)

### Scripts path
$null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'ScriptsPath' -Value (
    [string] [System.IO.Path]::Combine($Script:ModulesPathRoot,'Scripts')
)

### Create paths if they don't exist
if (-not $SystemContext) {
    $([string[]]($Script:ModulesPath,$Script:ScriptsPath)).ForEach{
        if (-not [System.IO.Directory]::Exists($_)) {
            $null = [System.IO.Directory]::CreateDirectory($_)
        }
    }
}




# Introduce script run
Write-Information -MessageData ('# Script start at {0}' -f $TimeTotalStart.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Information -MessageData (
    'Scope: "{0}" ("{1}").' -f (
        $(if ($SystemContext){'System'}else{'User'}),
        $Script:Scope
    )
)



# Make sure statistics are cleared for each run
foreach ($VariableName in [string[]]($Script:StatisticsVariables | Select-Object -ExpandProperty 'VariableName')) {
    $null = Get-Variable -Name $VariableName -Scope 'Script' -ErrorAction 'SilentlyContinue' | Remove-Variable -Scope 'Script' -Force
}



# Make sure not to install to OneDrive if script is running in user context
if (-not $SystemContext) {
    Set-PSModulePathUserContext
}



# Failproof
## Running as admin if $SystemContext
if ($SystemContext) {
    $IsAdmin = [bool]$([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ((-not($IsAdmin)) -or ([System.Environment]::Is64BitOperatingSystem -and (-not([System.Environment]::Is64BitProcess)))) {
        Write-Information -MessageData ('Run this script with 64 bit PowerShell as Admin!')
        Exit 1
    }
}


## User errors
### Check that same module is not specified in both $ModulesWanted and $ModulesUnwanted
if (($ModulesWanted | Where-Object -FilterScript {$ModulesUnwanted -contains $_}).'Count' -ge 1) {
    Throw ('ERROR - Same module(s) are specified in both $ModulesWanted and $ModulesUnwanted.')
}


## Dependencies for the script to run successfully
### Check that PowerShellGallery is available
if (
    -not $(
        if ($PSVersionTable.'PSEdition' -eq 'Core') {
            Test-Connection -TargetName 'powershellgallery.com' -TCPPort 443 -TimeoutSeconds 2 -IPv4 -Quiet -ErrorAction 'SilentlyContinue'
        }
        else {
            Test-NetConnection -ComputerName 'powershellgallery.com' -Port 443 -InformationLevel 'Quiet' -ErrorAction 'SilentlyContinue'
        }
    )
) {
    Throw ('ERROR - Could not TCP connect to powershellgallery.com:443 within reasonable time. Do you have internet connection, or is the web site down?')
}

### Check that PowerShellGallery is responding within reasonable time
if (-not[bool]$(Try{$null = Invoke-RestMethod -Method 'Get' -Uri 'https://www.powershellgallery.com' -TimeoutSec 3; $?}Catch{$false})) {
    Throw ('ERROR - Could TCP connect to powershellgallery.com:443, but it did not manage to provide a response within reasonable time.')
}



# Set Script Scope Variables
$Script:ModulesInstalledNeedsRefresh = [bool] $true
$Script:StatisticsVariables = [PSCustomObject[]](
    [PSCustomObject]@{'VariableName' = 'ModulesInstalledMissing';   'Description' = 'Modules installed (was missing)'},
    [PSCustomObject]@{'VariableName' = 'ModulesSubInstalledMissing';'Description' = 'Submodules installed (was missing)'},
    [PSCustomObject]@{'VariableName' = 'ModulesUpdated';            'Description' = 'Modules updated'},
    [PSCustomObject]@{'VariableName' = 'ModulesUninstalledOutdated';'Description' = 'Modules uninstalled (outdated)'},
    [PSCustomObject]@{'VariableName' = 'ModulesUninstalledUnwanted';'Description' = 'Modules uninstalled (unwanted)'},
    [PSCustomObject]@{'VariableName' = 'ScriptsMissing';            'Description' = 'Scripts installed (was missing)'},
    [PSCustomObject]@{'VariableName' = 'ScriptsUpdated';            'Description' = 'Scripts updated'}
)



# Set settings
## Disable proxies for speed
[System.Net.WebRequest]::DefaultWebProxy = $null


## Use TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12



# Prerequirement module "Microsoft.PowerShell.PSResourceGet"
## Introduce step
Write-Information -MessageData ('{0}# Prerequirement module "Microsoft.PowerShell.PSResourceGet"' -f ([System.Environment]::NewLine * 2))

## Find conflicting assemblies that are already loaded
$PsrgConflictingAssembly = [object](
    [Threading.Thread]::GetDomain().GetAssemblies().Where{
        -not [string]::IsNullOrEmpty($_.'Location') -and
        $_.'Location'.Contains('PSResourceGet') -and
        $_.'ManifestModule'.'Name' -eq 'Microsoft.PowerShell.PSResourceGet.dll'
    }[0]
)

## If assemblies are already loaded => Use that path
if (-not [string]::IsNullOrEmpty($PsrgConflictingAssembly.'ManifestModule'.'Name')) {
    Write-Information -MessageData 'Found a conflicting assembly, must use that path for importing PSResourceGet.'
    if ((Get-Module -Name 'Microsoft.PowerShell.PSResourceGet').'Count' -le 0) {
        $null = Import-Module -Name $PsrgConflictingAssembly.'ManifestModule'.'Name'
        Write-Information -MessageData (
            'v{0} was successfully imported.' -f (Get-Module -Name 'Microsoft.PowerShell.PSResourceGet').'Version'.ToString()
        )
    }
    else {
        Write-Information -MessageData (
            'v{0} is already imported.' -f (Get-Module -Name 'Microsoft.PowerShell.PSResourceGet').'Version'.ToString()
        )
    }
}


## Else - Use separate version maintained by this script
else {
    # Assets
    $ScriptHomePath = [string] [System.IO.Path]::Combine($PSResourceHomeUser,'PowerShellModulesUpdater')
    $ScriptModulePath = [string] [System.IO.Path]::Combine($ScriptHomePath,'Modules')
    $PsrgParentDirectory = [string] [System.IO.Path]::Combine($ScriptModulePath,'Microsoft.PowerShell.PSResourceGet')

    # Temporarily set PSModulePath to the path managed by this script
    $PSModulePathCurrent = [System.Environment]::GetEnvironmentVariable('PSModulePath','Process')
    $null = [System.Environment]::SetEnvironmentVariable('PSModulePath',$ScriptModulePath,'Process')

    # Check installed version vs. latest version
    ## Get version installed
    $PsrgInstalledVersion = [System.Version](
        $(
            if ([System.IO.Directory]::Exists($PsrgParentDirectory)) {
                [System.IO.Directory]::GetDirectories($PsrgParentDirectory).Where{
                    [System.IO.File]::Exists([System.IO.Path]::Combine($_,'Microsoft.PowerShell.PSResourceGet.psd1'))
                }.ForEach{
                    Try {
                        $_.Split([System.IO.Path]::DirectorySeparatorChar)[-1] -as [System.Version]
                    }
                    Catch {
                        '0.0'
                    }
                } | Sort-Object -Property @{'Expression' ={[System.Version]$_}} -Descending | Select-Object -First 1
            }
            else {
                '0.0'
            }
        )
    )
    ## Get latest GA from PowerShell Gallery
    $PsrgLatestVersion = [System.Version](
        (
            Invoke-RestMethod -Method 'Get' -Uri (
                "https://www.powershellgallery.com/api/v2/Packages?`$filter=IsLatestVersion and IsPrerelease eq false and Id eq 'Microsoft.PowerShell.PSResourceGet'&semVerLevel=1.0.0"
            )
        ).'properties'.'Version'
    )
    ## Create expected install directory based on latest version info
    $PsrgInstallDirectory = [string] [System.IO.Path]::Combine($PsrgParentDirectory,$PsrgLatestVersion.ToString())
    $PsrgImportDll  = [string] [System.IO.Path]::Combine($PsrgInstallDirectory,'Microsoft.PowerShell.PSResourceGet.dll')

    # Install latest version
    if (
        -not [System.IO.Directory]::Exists($PsrgInstallDirectory) -or
        $PsrgInstalledVersion -le [System.Version]('0.0') -or
        $PsrgInstalledVersion -lt $PsrgLatestVersion
    ) {
        if ($PsrgInstalledVersion -le [System.Version]('0.0')) {
            Write-Information -MessageData 'Not installed.'
        }
        else {
            Write-Information -MessageData (
                'Current installed version v{0} is not latest version v{1}.' -f (
                    $PsrgInstalledVersion.ToString(),
                    $PsrgLatestVersion.ToString()
                )
            )
        }
        Write-Information -MessageData 'Installing latest version.'

        # Assets
        $PsrgUri = [string] 'https://www.powershellgallery.com/api/v2/package/Microsoft.PowerShell.PSResourceGet/{0}' -f $PsrgLatestVersion
        $PsrgDownloadFilePath = [string] [System.IO.Path]::Combine(
            $(if ($PSVersionTable.'Platform' -eq 'Unix'){'/tmp'}else{$env:TEMP}),
            'Microsoft.PowerShell.PSResourceGet.{0}.zip' -f $PsrgLatestversion.ToString()
        )

        # Unload module if already loaded
        $null = Get-Module -Name 'Microsoft.PowerShell.PSResourceGet' | Remove-Module

        # Create install path
        if (-not [System.IO.Directory]::Exists($PsrgInstallDirectory)) {
            $null = [System.IO.Directory]::CreateDirectory($PsrgInstallDirectory)
        }

        # Download
        $null = [System.Net.WebClient]::new().DownloadFile($PsrgUri,$PsrgDownloadFilePath)

        # Extract
        $null = Expand-Archive -Path $PsrgDownloadFilePath -DestinationPath $PsrgInstallDirectory

        # Delete download file
        $null = [System.IO.File]::Delete($PsrgDownloadFilePath)

        # Delete older version of the module
        [System.IO.Directory]::GetDirectories($PsrgParentDirectory).Where{
            $_ -ne $PsrgInstallDirectory
        }.ForEach{
            [System.IO.Directory]::Delete($_,$true)
        }
    }

    # Import module
    $PsrgImported = Get-Module -Name 'Microsoft.PowerShell.PSResourceGet'
    if ($PsrgImported.'Path' -eq $PsrgImportDll -and $PsrgImported.'Version' -ge $PsrgLatestVersion) {
        Write-Information -MessageData 'Correct version of "Microsoft.PowerShell.PSResourceGet" is already imported.'
    }
    else {
        if (-not [string]::IsNullOrEmpty($PsrgImported.'Name')) {
            $null = Remove-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Force
        }
        Write-Information -MessageData 'Importing "Microsoft.PowerShell.PSResourceGet".'
        $null = Import-Module -Name $PsrgParentDirectory
    }

    # Set PSModulePath back to what it was
    $null = [System.Environment]::SetEnvironmentVariable('PSModulePath',$PSModulePathCurrent,'Process')
}



# Modules
## Introduce step
Write-Information -MessageData ('{0}# Modules' -f ([System.Environment]::NewLine * 2))

## Uninstall Unwanted Modules
Write-Information -MessageData '## Uninstall unwanted modules'
if ($UninstallUnwantedModules) {
    Uninstall-ModulesUnwanted -ModulesUnwanted $ModulesUnwanted
}
else {
    Write-Information -MessageData ('{0}Uninstall unwanted modules is set to $false.' -f $Indentation)
}

## Update Installed Modules
Write-Information -MessageData ('{0}## Update installed modules' -f [System.Environment]::NewLine)
if ($InstallUpdatedModules) {
    Update-ModulesInstalled
}
else {
    Write-Information -MessageData ('{0}Update installed modules is set to $false.' -f $Indentation)
}

## Install Missing Modules
Write-Information -MessageData ('{0}## Install missing modules' -f [System.Environment]::NewLine)
if ($InstallMissingModules) {
    Install-ModulesMissing -ModulesWanted $ModulesWanted
}
else {
    Write-Information -MessageData ('{0}Install missing modules is set to $false.' -f $Indentation)
}

## Installing Missing Sub Modules
Write-Information -MessageData ('{0}## Install missing sub modules' -f [System.Environment]::NewLine)
if ($InstallMissingSubModules) {
    Install-SubModulesMissing
}
else {
    Write-Information -MessageData ('{0}Install missing sub modules is set to $false.' -f $Indentation)
}

## Remove old modules
Write-Information -MessageData ('{0}## Remove outdated modules' -f [System.Environment]::NewLine)
if ($UninstallOutdatedModules) {
    Uninstall-ModulesOutdated
}
else {
    Write-Information -MessageData ('{0}Remove outdated modules is set to $false.' -f $Indentation)
}



# Scripts
if ($DoScripts) {
    # Introduce step
    Write-Information -MessageData ('{0}# Scripts' -f ([System.Environment]::NewLine * 2))

    # Install missing scripts
    Write-Information -MessageData '## Install missing scripts if any'
    Install-ScriptsMissing -Scripts $Script:ScriptsWanted

    # Update existing scripts if newer are available
    Write-Information -MessageData ('{0}## Update installed scripts' -f [System.Environment]::NewLine)
    Update-ScriptsInstalled
}


# Write Stats
Write-Information -MessageData ('{0}# Finished.' -f ([System.Environment]::NewLine * 2))
Write-Information -MessageData ('## Stats')
Write-Statistics
Write-Information -MessageData ('{0}## Time' -f [System.Environment]::NewLine)
Write-Information -MessageData ('Start time:    {0} ({1}).' -f $Script:TimeTotalStart.ToString('HH\:mm\:ss'), $Script:TimeTotalStart.ToString('o'))
Write-Information -MessageData ('End time:      {0} ({1}).' -f (($Script:TimeTotalEnd = [datetime]::Now).ToString('HH\:mm\:ss'),$Script:TimeTotalEnd.ToString('o')))
Write-Information -MessageData ('Total runtime: {0}.' -f ([string]$([timespan]$($Script:TimeTotalEnd - $Script:TimeTotalStart)).ToString('hh\:mm\:ss')))
#endregion Main
