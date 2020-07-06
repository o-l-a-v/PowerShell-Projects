﻿#Requires -Version 5.1 -RunAsAdministrator
<#
    .NAME
        PowerShellModulesUpdater.ps1

    .SYNOPSIS
        Updates, installs and removes PowerShell modules on your system based on settings in the "Settings & Variables" region.

    .DESCRIPTION
        Updates, installs and removes PowerShell modules on your system based on settings in the "Settings & Variables" region.

        How to use:
            1. Remember to allow script execution!
                Set-ExecutionPolicy -ExecutionPolicy 'Unrestricted' -Scope 'CurrentUser' -Force -Confirm:$false
            2. Some modules requires you to accept a license. This can be passed as a parameter to "Install-Module".
               Set $AcceptLicenses to $true (default) if you want these modules to be installed as well. 
               Some modules that requires $AcceptLicenses to be $true:
                   * Az.ApplicationMonitor

        Requirements
            * Package Providers
                * NuGet              
            * PowerShell Modules
                * PackageManagement  https://www.powershellgallery.com/packages/PackageManagement
                * PowerShellGet      https://www.powershellgallery.com/packages/PowerShellGet          

    .NOTES
        Author:        Olav Rønnestad Birkeland
        Version:       1.8.0.0
        Creation Date: 190310
        Modified Date: 200407    

    .EXAMPLE
        # Run from PowerShell ISE
        & ('{0}\Setup Lighthouse Subscription.ps1' -f ([System.IO.Directory]::GetParent($psISE.'CurrentFile'.'FullPath').'FullName'))

    .EXAMPLE
        # Run from PowerShell ISE, bypass script execution policy
        & powershell.exe -ExecutionPolicy "Bypass" -NoLogo -NonInteractive -NoProfile -WindowStyle "Hidden" -File ('{0}\Setup Lighthouse Subscription.ps1' -f ([System.IO.Directory]::GetParent($psISE.'CurrentFile'.'FullPath').'FullName'))    
#>


# Only continue if running as Admin, and if 64 bit OS => Running 64 bit PowerShell
$IsAdmin = [bool]$([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ((-not($IsAdmin)) -or ([System.Environment]::Is64BitOperatingSystem -and (-not([System.Environment]::Is64BitProcess)))) {
    Write-Output -InputObject ('Run this script with 64 bit PowerShell as Admin!')
}


else {
#region    Settings & Variables
    # Install-Module preferences
    ## Accept licenses - Some modules requires you to accept licenses.
    $AcceptLicenses           = [bool] $true
    ## Skip Publisher Check - Security, checks signing of the module against alleged publisher.
    $SkipPublisherCheck       = [bool] $false

    # Action - What Options Would You Like To Perform
    $InstallMissingModules    = [bool] $true
    $InstallMissingSubModules = [bool] $true
    $InstallUpdatedModules    = [bool] $true
    $UninstallOutdatedModules = [bool] $true
    $UninstallUnwantedModules = [bool] $true
       
    # Settings - PowerShell Output Streams
    $VerbosePreference  = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'

    # List of wanted modules
    $ModulesWanted = [string[]]$(
        'Az',                                     # Microsoft. Used for Azure Resources. Combines and extends functionality from AzureRM and AzureRM.Netcore.
        'Azure',                                  # Microsoft. Used for managing Classic Azure resources/ objects.        
        'AzureAD',                                # Microsoft. Used for managing Azure Active Directory resources/ objects.
        'AzureADPreview',                         # -^-
        'AzureRM',                                # (DEPRECATED, "Az" is it's successor). Microsoft. Used for managing Azure Resource Manager resources/ objects
        'ConfluencePS',                           # Atlassian, for interacting with Atlassian Confluence Rest API.
        'GetBIOS',                                # Damien Van Robaeys. Used for getting BIOS settings for Lenovo, Dell and HP.
        'ExchangeOnlineManagement',               # Microsoft. Used for managing Exchange Online.
        'ImportExcel',                            # dfinke.    Used for import/export to Excel.
        'Intune.USB.Creator',                     # Ben Reader @ powers-hell.com. Used to create Autopilot WinPE.
        'IntuneBackupAndRestore',                 # John Seerden. Uses "MSGraphFunctions" module to backup and restore Intune config.
        'Invokeall',                              # Santhosh Sethumadhavan. Multithread PowerShell commands.
        'ISESteroids',                            # Power The Shell, ISE Steroids. Used for extending PowerShell ISE functionality.
        'JWTDetails',                             # Darren J. Robinson. Used for decoding JWT, JSON Web Tokens.
        'Microsoft.Graph',                        # Microsoft. Works with PowerShell Core.
        'Microsoft.Graph.Intune',                 # Microsoft. Used for managing Intune using PowerShell Graph in the backend.
        'Microsoft.Online.SharePoint.PowerShell', # Microsoft. Used for managing SharePoint Online.
        'Microsoft.PowerShell.ConsoleGuiTools',   # Microsoft.
        'Microsoft.PowerShell.SecretsManagement', # Microsoft. Used for securely storing secrets locally.
        'Microsoft.RDInfra.RDPowerShell',         # Microsoft. Used for managing Windows Virtual Desktop.
        'MicrosoftGraphSecurity',                 # Microsoft. Used for interacting with Microsoft Graph Security API.
        'MSGraphFunctions',                       # John Seerden. Wrapper for Microsoft Graph Rest API.
        'MSOnline',                               # (DEPRECATED, "AzureAD" is it's successor) Microsoft. Used for managing Microsoft Cloud Objects (Users, Groups, Devices, Domains...)
        'newtonsoft.json',                        # Serialize/Deserialize Json using Newtonsoft.json
        'PartnerCenter',                          # Microsoft. Used for interacting with PartnerCenter API.
        'PackageManagement',                      # Microsoft. Used for installing/ uninstalling modules.
        'platyPS',                                # Microsoft. Used for converting markdown to PowerShell XML external help files.
        'PolicyFileEditor',                       # Microsoft. Used for local group policy / gpedit.msc.
        'PoshRSJob',                              # Boe Prox. Used for parallel execution of PowerShell.
        'PowerShellGet',                          # Microsoft. Used for installing updates.
        'PSReadLine',                             # Microsoft. Used for helping when scripting PowerShell.
        'PSScriptAnalyzer',                       # Microsoft. Used to analyze PowerShell scripts to look for common mistakes + give advice.
        'PSWindowsUpdate',                        # Michal Gajda. Used for updating Windows.
        'SetBIOS',                                # Damien Van Robaeys. Used for setting BIOS settings for Lenovo, Dell and HP.
        'WindowsAutoPilotIntune'                  # Michael Niehaus @ Microsoft. Used for Intune AutoPilot stuff.
    )
    
    # List of Unwanted Modules - Will Remove Every Related Module, for AzureRM for instance will also search for AzureRM.*
    $ModulesUnwanted = [string[]]$(
        'PartnerCenterModule'                     # (DEPRECATED, "PartnerCenter" is it's successor). Microsoft. Used for interacting with Partner Center API.
    )

    # List of modules you don't want to get updated
    $ModulesDontUpdate = [string[]]$(
        ''
    )
#endregion Settings & Variables




#region    Functions
    #region    Get-ModuleInstalledVersions
    function Get-ModuleInstalledVersions {
        <#
            .SYNOPSIS
                Gets all available versions of a module name.

            .PARAMETER ModuleName
                String, name of the module you want to check.
        #>
        [CmdletBinding(SupportsPaging=$false)]
        [OutputType([System.Version[]])]            
        Param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string] $ModuleName
        )

        Begin {}

        Process {
            # Return installed versions of $ModuleName
            return [System.Version[]]$(
                $Versions = [System.Version[]]$()
                $Path = [string]$('{0}\WindowsPowerShell\Modules\{1}' -f ([string]$(if([System.Environment]::Is64BitProcess){$env:ProgramW6432}else{${env:ProgramFiles(x86)}}),$ModuleName))
                if ([bool]$(Test-Path -Path $Path -ErrorAction 'SilentlyContinue') -and [bool]$($ModuleName -notin [string[]]$('PackageManagement','PowerShellGet'))) {
                    $Versions = [System.Version[]]$(Get-ChildItem -Path $Path -Depth 0 | Select-Object -ExpandProperty 'FullName').ForEach{$_.Split('\')[-1]}.Where{Try{[System.Version]$($_);$?}Catch{$false}}
                }
                if ($Versions.'Count' -le 0) {
                    $Versions = [System.Version[]]$(Get-InstalledModule -Name $ModuleName -AllVersions | Select-Object -ExpandProperty 'Version')
                }
                $Versions
            )
        }

        End {}
    }
    #endregion Get-ModuleInstalledVersions
    
    
    #region    Get-ModulePublishedVersion
    function Get-ModulePublishedVersion {
        <#
            .SYNAPSIS
                Fetches latest version number of a given module from PowerShellGallery.
            
            .PARAMETER ModuleName
                String, name of the module you want to check.
        #>            
        [CmdletBinding(SupportsPaging=$false)]
        [OutputType([System.Version])]    
        param(
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string] $ModuleName
        )

        Begin {}

        Process {                                
            # Access the main module page, and add a random number to trick proxies
            $Url = [string]('https://www.powershellgallery.com/packages/{0}/?dummy={1}' -f ($ModuleName,$Script:Random.Next(9999)))
            Write-Debug -Message ('URL for module "{0}" = "{1}".' -f ($ModuleName,$Url))
            

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

        End {}
    }
    #endregion Get-ModulePublishedVersion



    #region    Refresh-ModulesInstalled
    function Refresh-ModulesInstalled {
        <#
            .SYNAPSIS
                Gets all currentlyy installed modules.
        #>
        [CmdletBinding(SupportsPaging=$false)]
        [OutputType([PSCustomObject[]])]    
        Param(
            [Parameter(Mandatory = $false)]
            [switch] $AllVersions,

            [Parameter(Mandatory = $false)]
            [switch] $ForceRefresh
        )
            
        Begin {}
            
        Process {
            # Check if variable needs refresh
            if ($ForceRefresh -or $Script:ModulesInstalledNeedsRefresh -or -not [bool]$($null = Get-Variable -Name 'ModulesInstalledNeedsRefresh' -Scope 'Script' -ErrorAction 'SilentlyContinue'; $?)) {
                # Reset Script Scrope Variable "ModulesInstalledNeedsRefresh" to $false
                $null = Set-Variable -Scope 'Script' -Option 'None' -Force -Name 'ModulesInstalledNeedsRefresh' -Value ([bool]$false)
                        
                # Get installed modules                    
                $null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'ModulesInstalled' -Value $(
                    ([PSCustomObject[]]$([array]$(Get-Package -Name '*' -ProviderName 'PowerShellGet' -AllVersions:$AllVersions) | Select-Object -Property 'Name','Version' | Sort-Object -Property 'Name','Version'))
                )
            }
        }

        End {}
    }
    #endregion Refresh-ModulesInstalled



    #region    Update-ModulesInstalled
    function Update-ModulesInstalled {
        <#
            .SYNAPSIS
                Fetches latest version number of a given module from PowerShellGallery.
        #>
        [CmdletBinding(SupportsPaging=$false)]
        [OutputType($null)]
        Param()

        Begin {}

        Process {
            # Refresh Installed Modules variable
            $null = Refresh-ModulesInstalled            

            # Skip if no installed modules was found
            if ($Script:ModulesInstalled.'Count' -le 0) {
                Write-Output -InputObject ('No installed modules where found, no modules to update.')
                Break
            }

            # Help Variables
            $C = [uint16]$(1)
            $CTotal = [string]$($Script:ModulesInstalled.'Count')
            $Digits = [string]$('0' * $CTotal.'Length')
            $ModulesInstalledNames = [string[]]$($Script:ModulesInstalled | Select-Object -ExpandProperty 'Name' | Sort-Object)

            # Update Modules
            :ForEachModule foreach ($ModuleName in $ModulesInstalledNames) {                
                # Get Latest Available Version
                $VersionAvailable = [System.Version]$(Get-ModulePublishedVersion -ModuleName $ModuleName)
                
                # Get Version Installed
                $VersionInstalled = [System.Version]$($Script:ModulesInstalled.Where{$_.'Name' -eq $ModuleName}.'Version')
                
                # Get Version Installed - Get Fresh Version Number if newer version is available and current module is a sub module
                if ([System.Version]($VersionAvailable) -gt [System.Version]$($VersionInstalled) -and $ModuleName -like '*?.?*' -and [string[]]$($ModulesInstalledNames) -contains [string]$($ModuleName.Replace(('.{0}' -f ($ModuleName.Split('.')[-1])),''))) {
                    $VersionInstalled = [System.Version](Get-ModuleInstalledVersions -ModuleName $ModuleName | Sort-Object -Descending | Select-Object -First 1)                        
                }
                
                # Present Current Module
                Write-Output -InputObject ('{0}/{1} {2} v{3}' -f (($C++).ToString($Digits),$CTotal,$ModuleName,$VersionInstalled.ToString()))

                # Compare Version Installed vs Version Available
                if ([System.Version]$($VersionInstalled) -ge [System.Version]$($VersionAvailable)) {
                    Write-Output -InputObject ('{0}Current version is latest version.' -f ("`t"))
                    Continue ForEachModule
                }
                else {               
                    Write-Output -InputObject ('{0}Newer version available, v{1}.' -f ("`t",$VersionAvailable.ToString()))               
                    if ([bool]$($null=Find-Package -Name $ModuleName -Source 'PSGallery' -AllVersions:$false -Force -ErrorAction 'SilentlyContinue';$?)) {
                        if ($ModulesDontUpdate -contains $ModuleName) {
                            Write-Output -InputObject ('{0}{0}Will not update as module is specified in script settings. ($ModulesDontUpdate).' -f ("`t",$Success.ToString))
                        }
                        else {
                            # Install module
                            $Success = [bool]$(Try{$null=PackageManagement\Install-Package -Name $ModuleName -RequiredVersion $VersionAvailable -Scope 'AllUsers' -AllowClobber -AcceptLicense:$AcceptLicenses -SkipPublisherCheck:$SkipPublisherCheck -Confirm:$false -Verbose:$false -Debug:$false -Force 2>$null;$?}Catch{$false})
                                
                            # Doubles check for success
                            if ($Success) {
                                $Success = [bool](Test-Path -Path ('{0}\WindowsPowerShell\Modules\{1}\{2}' -f ($env:ProgramW6432,$ModuleName,$VersionAvailable.ToString())))
                            }
                                
                            # Output success
                            Write-Output -InputObject ('{0}Install success? {1}' -f ([string]$("`t" * 2),$Success.ToString()))
                                
                            # Count as installed if success
                            if ($Success) {
                                # Stats
                                $Script:ModulesUpdated += [string[]]$($ModuleName)
                                # Updated cache of installed modules and version if current module has sub modules
                                if ([uint16]$([string[]]$($ModulesInstalledNames.Where{$_ -like ('{0}.?*' -f ($Module))}).'Count') -ge 1) {
                                    Get-ModulesInstalled
                                }
                                # Else, set flag to update cache of installed modules later
                                else {
                                    $Script:ModulesInstalledNeedsRefresh = $true
                                }
                            }
                            else {
                                Write-Warning -Message ('Failed to install "{0}". Here is the error:{1}"{2}"' -f ($ModuleName,[System.Environment]::NewLine,$Error[0].ToString()))
                            }
                        }
                    }
                    else {
                        Write-Warning -Message ('Did not find "{0}" in PSGallery, probably deprecated, delisted or something similar. Will skip.' -f ($ModuleName))
                    }
                }
            }
        }
            
        End {}
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
        [CmdletBinding(SupportsPaging=$false)]
        [OutputType($null)]
        Param(
            [Parameter(Mandatory = $true)]
            [string[]] $ModulesWanted                
        )

        Begin {}

        Process {
            # Refresh Installed Modules variable
            $null = Refresh-ModulesInstalled
            

            # Help Variables
            $C = [uint16](1)
            $CTotal = [string]($ModulesWanted.'Count')
            $Digits = [string]('0' * $CTotal.'Length')
            $ModulesInstalledNames = [string[]]($Script:ModulesInstalled.'Name' | Sort-Object)


            # Loop each wanted module. If not found in installed modules: Install it
            foreach ($ModuleWanted in $ModulesWanted) {
                Write-Output -InputObject ('{0}/{1} {2}' -f (($C++).ToString($Digits),$CTotal,$ModuleWanted))
                
                # Install if not already installed
                if ($ModulesInstalledNames.Where{$_ -eq $ModuleWanted}.'Count' -ge 1) {
                    Write-Output -InputObject ('{0}Already Installed. Next!' -f ("`t"))
                }
                else {
                    Write-Output -InputObject ('{0}Not already installed. Installing.' -f ("`t"))
                    if ([bool]$($null=Find-Package -Name $ModuleWanted -Source 'PSGallery' -AllVersions:$false -Force -ErrorAction 'SilentlyContinue';$?)) {
                        # Install The Missing Sub Module
                        $Success = [bool]$(Try{$null=PackageManagement\Install-Package -Name $ModuleWanted -Scope 'AllUsers' -AllowClobber -AcceptLicense:$AcceptLicenses -SkipPublisherCheck:$SkipPublisherCheck -Confirm:$false -Verbose:$false -Debug:$false -Force 2>$null;$?}Catch{$false})                                                           
                                
                        # Output success
                        Write-Output -InputObject ('{0}Install success? {1}' -f ([string]$("`t" * 2),$Success.ToString()))
                            
                        # Count if success
                        if ($Success) {
                            # Stats
                            $Script:ModulesInstalledMissing += [string[]]$($ModuleWanted)
                            # Make sure list of installed modules gets refreshed
                            $Script:ModulesInstalledNeedsRefresh = $true                            
                        }
                        else {
                            Write-Warning -Message ('Failed to install "{0}". Here is the error:{1}"{2}"' -f ($ModuleWanted,[System.Environment]::NewLine,$Error[0].ToString()))
                        }
                    }
                    else {
                        Write-Warning -Message ('Did not find "{0}" in PSGallery, probably deprecated, delisted or something similar. Will skip.' -f ($ModuleWanted))
                    }
                }
            }
        }

        End {}
    }
    #endregion Install-ModulesMissing


    #region    Install-SubModulesMissing
    function Install-SubModulesMissing {
        <#
            .SYNAPSIS
                Installs Eventually Missing Submodules

            .PARAMETER ModulesName
                String containing the name of the parent module you want to check for missing submodules.
        #>
        [CmdletBinding(SupportsPaging=$false)]
        [OutputType($null)]
        Param()

        Begin {}

        Process {
            # Refresh Installed Modules variable
            $null = Refresh-ModulesInstalled

            # Skip if no installed modules was found
            if ($Script:ModulesInstalled.'Count' -le 0) {
                Write-Output -InputObject ('No installed modules where found, no modules to check against.')
                Break
            }

            # Help Variables - Both Foreach
            $ModulesInstalledNames       = [string[]]$($Script:ModulesInstalled.'Name' | Sort-Object)
            $ParentModulesInstalledNames = [string[]]$(
                $ModulesInstalledNames.Where{
                    [bool]($_ -in $ModulesWanted) -or
                    [bool]($_ -notlike '*.*') -or 
                    [bool]($_ -like '*.*' -and $_ -notlike '*.*.*' -and [string[]]$($ModulesInstalledNames) -notcontains [string]$($_.Replace(('.{0}' -f ($_.Split('.')[-1])),'')))
                }
            )
            
            # Help Variables - Outer Foreach
            $OC = [uint16]$(1)
            $OCTotal = [string]$($ParentModulesInstalledNames.'Count'.ToString())
            $ODigits = [string]$('0' * $OCTotal.'Length')

            # Loop Through All Installed Modules
            :ForEachModule foreach ($ModuleName in $ParentModulesInstalledNames) {
                # Present Current Module
                Write-Output -InputObject ('{0}/{1} {2}' -f (($OC++).ToString($ODigits),$OCTotal,$ModuleName))
                
                # Get all installed sub modules
                $SubModulesInstalled = [string[]]$($ModulesInstalledNames.Where{$_ -like ('{0}.*' -f ($ModuleName))} | Sort-Object)

                # Get all available sub modules
                $SubModulesAvailable = [string[]]$($([array](Find-Module -Name ('{0}.*' -f ($ModuleName)))).'Name' | Sort-Object)

                # If either $SubModulesAvailable is 0, Continue Outer Foreach
                if ($SubModulesAvailable.'Count' -eq 0) {
                    Write-Output -InputObject ('{0}Found {1} avilable sub module{2}.' -f ("`t",$SubModulesAvailable.'Count'.ToString(),[string]$(if($SubModulesAvailable.'Count' -ne 1){'s'})))
                    Continue ForEachModule
                }

                # Compare objects to see which are missing
                $SubModulesMissing = [string[]]$(
                    if ($SubModulesInstalled.'Count' -eq 0) {
                        $SubModulesAvailable
                    }
                    else {
                        $SubModulesAvailable.Where{$_ -notin $SubModulesInstalled}
                    }
                )
                ## Output result
                Write-Output -InputObject ('{0}Found {1} missing sub module{2}.' -f ("`t",$SubModulesMissing.'Count'.ToString(),[string]$(if($SubModulesMissing.'Count' -ne 1){'s'})))

                # Install missing sub modules
                if ($SubModulesMissing.'Count' -gt 0) {
                    # Help Variables - Inner Foreach
                    $IC = [uint16]$(1)
                    $ICTotal = [string]$($SubModulesMissing.'Count'.ToSTring())
                    $IDigits = [string]$('0' * $ICTotal.'Length')
    
                    # Install Modules
                    :ForEachSubModule foreach ($SubModuleName in $SubModulesMissing) {
                        # Present Current Sub Module
                        Write-Output -InputObject ('{0}{1}/{2} {3}' -f ([string]$("`t" * 2),($IC++).ToString($IDigits),$ICTotal,$SubModuleName))
                            
                        # Make sure package is actually available
                        $PackageAvailable = Find-Package -Name $SubModuleName -Source 'PSGallery' -AllVersions:$false -Force -ErrorAction 'SilentlyContinue'
                            
                        # Install if package available
                        if ($?) {
                            # Install The Missing Sub Module
                            $Success = [bool]$(Try{$null=PackageManagement\Install-Package -Name $SubModuleName -RequiredVersion $PackageAvailable.'Version' -Scope 'AllUsers' -AllowClobber -AcceptLicense:$AcceptLicenses -SkipPublisherCheck:$SkipPublisherCheck -Confirm:$false -Verbose:$false -Debug:$false -Force 2>$null;$?}Catch{$false})
                                
                            # Doubles check for success
                            if ($Success) {
                                $Success = [bool](Test-Path -Path ('{0}\WindowsPowerShell\Modules\{1}' -f ($env:ProgramW6432,$SubModuleName,$PackageAvailable.'Version'.ToString())))
                            }
                                
                            # Output success
                            Write-Output -InputObject ('{0}Install success? {1}' -f ([string]$("`t" * 3),$Success.ToString()))
                                
                            # Count as installed if success
                            if ($Success) {
                                # Stats
                                $Script:ModulesSubInstalledMissing += [string[]]$($SubModuleName)
                                # Make sure list of installed modules gets refreshed
                                $Script:ModulesInstalledNeedsRefresh = $true                                
                            }
                            else {
                                Write-Warning -Message ('Failed to install "{0}". Here is the error:{1}"{2}"' -f ($SubModuleName,[System.Environment]::NewLine,$Error[0].ToString()))
                            }
                        }
                        else {
                            Write-Warning -Message ('Did not find "{0}" in PSGallery, probably deprecated, delisted or something similar. Will skip.' -f ($SubModuleName))
                        }
                    }
                }
            }
        }

        End {}
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
        [OutputType($null)]
        Param (
            # Mandatory
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string] $ModuleName,

            [Parameter(Mandatory = $true)]
            [System.Version] $Version,

            # Optional
            [Parameter(Mandatory = $false)]
            [bool] $SystemContext = $true,

            [Parameter(Mandatory = $false)]
            [bool] $64BitModuleOn64BitOS = $true,

            [Parameter(Mandatory = $false)]
            [string] $ModulePath
        )

        Begin {
            # Check if running as admin if $SystemContext
            if ($SystemContext -and -not $([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Throw 'ERROR - Admin permissions are required if running in system context.'
            }
        }

        Process{
            # Unload module if currently in use
            if ($(Get-Module -Name $ModuleName).Where{[System.Version]$_.'Version' -eq $Version}.'Count' -gt 0) {
                $null = Remove-Module -Name $ModuleName -Force
            }

            # Create path to module
            if (-not $ModulePath -or [string]::IsNullOrEmpty($ModulePath)) {
                $ModulePath = [string](
                    '{0}\{1}\Modules\{2}\{3}' -f (
                        [string[]]$(
                            if ($SystemContext) {
                                if ($64BitModuleOn64BitOS) {
                                    $env:ProgramW6432
                                }
                                else {
                                    ${env:ProgramFiles(x86)}
                                },
                                'WindowsPowerShell'
                            }
                            else {
                                [System.Environment]::GetFolderPath('MyDocuments'),
                                'PowerShell'
                            }
                        ) +
                        $ModuleName,
                        $Version.ToString()
                    )
                )
            }

            # If path does not exist, try to check for versions and create path more dynamically
            if (-not [System.IO.Directory]::Exists($ModulePath)) {
                # Create module path
                $ModulePath = [string]('{0}\WindowsPowerShell\Modules\{1}' -f ($env:ProgramW6432,$ModuleName))
        
                # Get all versions
                $Versions   = [PSCustomObject[]](
                    $([array](Get-ChildItem -Path $ModulePath -Directory -Depth 0)).ForEach{
                        Try {
                            [PSCustomObject]@{
                                'Path'    = $_.'FullName'
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
                        $_.'Version' = [System.Version][string]($_.'Version'.ToString().Split('.')[0..$($([byte]($Version.ToString().Split('.').'Count'))-1)] -join '.')
                    }
                }

                # Set path
                if ([System.Version[]]($Versions.'Version') -contains $Version) {
                    $ModulePath = [string]($Versions.Where{$_.'Version' -eq $Version}.'Path')
                }
            }

            # Delete folder if it exists / Uninstall module version
            if ([bool]$(Try{$null=[System.Version]$ModulePath.Split('\')[-1];$?}Catch{$false}) -and [System.IO.Directory]::Exists($ModulePath)) {
                $null = [System.IO.Directory]::Delete($ModulePath,$true)
                if ($? -and -not [System.IO.Directory]::Exists($ModulePath)) {
                    return
                }
            }

            # Else, uninstall with commandlet
            $null = PackageManagement\Uninstall-Package -Name $ModuleName -RequiredVersion $Version -Scope $(if($SystemContext){'AllUsers'}else{'CurrentUser'}) -Force
        }

        End {}
    }
    #endregion Uninstall-ModuleManually


    #region    Uninstall-ModulesOutdated
    function Uninstall-ModulesOutdated {
        <#
            .SYNAPSIS
                Uninstalls outdated modules / currently installed modules with more than one version.
        #>
        [CmdletBinding(SupportsPaging=$false)]
        [OutputType($null)]
        Param()

        Begin {}

        Process {
            # Refresh Installed Modules variable
            $null = Refresh-ModulesInstalled
            
            # Skip if no installed modules were found
            if ($Script:ModulesInstalled.'Count' -le 0) {
                Write-Output -InputObject ('No installed modules where found, no outdated modules to uninstall.')
                Break
            }

            # Help Variables                
            $ModulesInstalledNames = [string[]]$($Script:ModulesInstalled.'Name' | Sort-Object -Unique)
            $C = [uint16]$(1)
            $CTotal = [string]$($ModulesInstalledNames.'Count')
            $Digits = [string]$('0' * $CTotal.'Length')

            # Get Versions of Installed Main Modules
            :ForEachModuleName foreach ($ModuleName in $ModulesInstalledNames) {
                Write-Output -InputObject ('{0}/{1} {2}' -f (($C++).ToString($Digits),$CTotal,$ModuleName))
                
                # Get all versions installed
                $VersionsAll = [System.Version[]]$(Get-ModuleInstalledVersions -ModuleName $ModuleName | Sort-Object -Descending)                    
                Write-Output -InputObject ('{0}{1} got {2} installed version{3} ({4}).' -f ("`t",$ModuleName,$VersionsAll.'Count'.ToString(),[string]$(if($VersionsAll.'Count' -gt 1){'s'}),[string]($VersionsAll -join ', ')))
            
                # Remove old versions if more than 1 versions
                if ($VersionsAll.'Count' -gt 1) {
                    # Find newest version
                    $VersionNewest        = [System.Version]($VersionsAll | Sort-Object -Descending | Select-Object -First 1)
                    $VersionsAllButNewest = [System.Version[]]($VersionsAll.Where{$_ -ne $VersionNewest})
                    
                    # If current module is PackageManagement
                    if ($ModuleName -eq 'PackageManagement') {
                        # Assets
                        $VersionsInUse = [System.Version[]]($(Get-Module -Name 'PackageManagement').'Version')

                        # Continue if any version but the newest are currently in use
                        if ([byte]$($([array]($([System.Version[]]($VersionsAllButNewest)).Where{$VersionsInUse -contains $_})).'Count') -gt 0) {
                            Write-Warning -Message ('{0}{0}"PackageManagement" was updated during this session.' -f ("`t")) -WarningAction 'Continue'
                            Write-Warning -Message ('{0}{0}Current PowerShell session must be closed (EXIT) before attempting to remove outdated versions.' -f ("`t")) -WarningAction 'Continue'
                            Continue ForEachModuleName
                        }
                    }

                    # Uninstall all but newest
                    foreach ($Version in $VersionsAllButNewest) {
                        Write-Output -InputObject ('{0}{0}Uninstalling module "{1}" version "{2}".' -f ("`t",$ModuleName,$Version.ToString()))                        
                            
                        # Uninstall
                        $Success = [bool]$(
                            Try {
                                $null = Uninstall-ModuleManually -ModuleName $ModuleName -Version $Version -WarningAction 'SilentlyContinue' -ErrorAction 'SilentlyContinue'
                                $?
                            }
                            Catch {
                                $false
                            }
                        )
                            
                        # Check for success
                        $Success = [bool]$(
                            if ($Success) {
                                $([array](PackageManagement\Get-Package -Name $ModuleName -RequiredVersion $Version -ErrorAction 'SilentlyContinue')).'Count' -eq 0
                            }
                            else {
                                $Success
                            }
                        )
                            
                        # Output
                        Write-Output -InputObject ('{0}{0}{0}Success? {1}.' -f ("`t",$Success.ToString()))
                        if ($Success) {
                            # Stats
                            $Script:ModulesUninstalledOutdated += [string[]]($ModuleName)
                            # Make sure list of installed modules gets refreshed
                            $Script:ModulesInstalledNeedsRefresh = $true                                
                        }
                        else {
                            # Special case for module "PackageManagement"
                            if ($ModuleName -eq 'PackageManagement') {
                                Write-Warning -Message ('{0}{0}{0}{0}"PackageManagement" often can`t be removed before exiting current PowerShell session.' -f ("`t")) -WarningAction 'Continue'
                            }
                            Continue ForEachModuleName
                        }
                    }
                }
            }
        }

        End {}
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
        [CmdletBinding(SupportsPaging=$false)]
        [OutputType($null)]
        Param(
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string[]] $ModulesUnwanted
        )

        Begin {}

        Process {
            # Refresh Installed Modules variable
            $null = Refresh-ModulesInstalled
            
            # Skip if no installed modules was found
            if ($Script:ModulesInstalled.'Count' -le 0) {
                Write-Output -InputObject ('No installed modules where found, no modules to uninstall.')
                Break
            }
                      
            # Find out if we got unwated modules installed based on input parameter $ModulesUnwanted vs $InstalledModulesAll
            $ModulesToRemove = [string[]]$(
                :ForEachInstalledModule foreach ($ModuleInstalled in [string[]]$($Script:ModulesInstalled.'Name')) {
                    :ForEachUnwantedModule foreach ($ModuleUnwanted in $ModulesUnwanted) {
                        if ($ModuleInstalled -eq $ModuleUnwanted -or $ModuleInstalled -like ('{0}.*' -f ($ModuleUnwanted))) {
                            $ModuleInstalled
                            Continue ForEachInstalledModule
                        }
                    }
                }
            ) | Sort-Object

            # Write Out How Many Unwanted Modules Was Found
            Write-Output -InputObject ('Found {0} unwanted module{1}.{2}' -f ($ModulesToRemove.'Count',$(if($ModulesToRemove.'Count' -ne 1){'s'}),$(if($ModulesToRemove.'Count' -gt 0){' Will proceed to uninstall them.'})))    

            # Uninstall Unwanted Modules If More Than 0 Was Found
            if ([uint16]$($ModulesToRemove.'Count') -gt 0) {
                $C      = [uint16]$(1)
                $CTotal = [string]$($ModulesToRemove.'Count'.ToString())
                $Digits = [string]$('0' * $CTotal.'Length')
                $LengthModuleLongestName = [byte]($([byte[]]($([string[]]($ModulesToRemove)).ForEach{$_.'Length'} | Sort-Object)) | Select-Object -Last 1)

                foreach ($Module in $ModulesToRemove) {
                    Write-Output -InputObject ('{0}{1}/{2} {3}{4}' -f ("`t",($C++).ToString($Digits),$CTotal,$Module,(' ' * [byte]$([byte]$($LengthModuleLongestName) - [byte]$($Module.'Length')))))
                    
                    # Do not uninstall "PowerShellGet"
                    if ($Module -eq 'PowerShellGet') {
                        Write-Output -InputObject ('{0}{0}Will not uninstall "PowerShellGet" as it`s a requirement for this script.' -f ("`t"))
                        Continue
                    }
                    
                    # Remove Current Module
                    Uninstall-Module -Name $Module -Confirm:$false -AllVersions -Force -ErrorAction 'SilentlyContinue'
                    $Success = [bool]$($?)

                    # Write Out Success
                    Write-Output -InputObject ('{0}{0}Success? {1}.' -f ("`t",$Success.ToString()))
                    if ($Success) {
                        # Stats
                        $Script:ModulesUninstalledUnwanted += [string[]]($Module)
                        # Make sure list of installed modules gets refreshed
                        $Script:ModulesInstalledNeedsRefresh = $true
                    }
                }
            }
        }

        End {}
    }
    #endregion Uninstall-ModulesUnwanted


    #region    Output-Statistics
    function Output-Statistics {
        [CmdletBinding()]
        [OutputType($null)]
        Param ()
            
        Begin {}

        Process {
            # Help variables
            $FormatDescriptionLength = [byte]$($Script:StatisticsVariables.ForEach{$_.'Description'.'Length'} | Sort-Object -Descending | Select-Object -First 1)
            $FormatNewLineTab        = [string]$("`r`n`t")

            # Output stats
            foreach ($Variable in $Script:StatisticsVariables) {
                $CurrentObject      = [string[]]$(Get-Variable -Name $Variable.'VariableName' -Scope 'Script' -ValueOnly -ErrorAction 'SilentlyContinue')
                $CurrentDescription = [string]$($Variable.'Description'+':'+[string]$(' '*[byte]$($FormatDescriptionLength-$Variable.'Description'.'Length')))
                Write-Output -InputObject (
                    '{0} {1}{2}' -f (
                        $CurrentDescription,
                        $CurrentObject.'Count',
                        [string]$(if($CurrentObject.'Count' -ge 1){$FormatNewLineTab+[string]$($CurrentObject -join $FormatNewLineTab)})
                    )
                )
            }
        }

        End {}
    }
    #endregion Output-Statistics
#endregion Functions




#region    Main
    # Help variables
    ## Start time
    $null = New-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'TimeTotalStart' -Value ([datetime]$([datetime]::Now))    
    ## Random generator object
    $null = New-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'Random' -Value ([System.Random]::New())

    # Make sure statistics are cleared for each run
    foreach ($VariableName in [string[]]($Script:StatisticsVariables | Select-Object -ExpandProperty 'VariableName')) {
        $null = Get-Variable -Name $VariableName -Scope 'Script' -ErrorAction 'SilentlyContinue' | Remove-Variable -Scope 'Script' -Force
    }

    # Check that same module is not specified in both $ModulesWanted and $ModulesUnwanted
    if (($ModulesWanted | Where-Object -FilterScript {$ModulesUnwanted -contains $_}).'Count' -ge 1) {
        Throw ('ERROR - Same module(s) are specified in both $ModulesWanted and $ModulesUnwanted.')
    }

    # Check that neccesary modules are not specified in $ModulesUnwanted
    if ([byte]($([string[]]($([string[]]('PackageManagement','PowerShellGet')).Where{$ModulesUnwanted -contains $_})).'Count') -gt 0) {
        Throw ('ERROR - Either "PackageManagement" or "PowerShellGet" was specified in $ModulesUnwanted. This is not supported as the script would not work correctly without them.')
    }

    # Set Script Scope Variables
    $Script:ModulesInstalledNeedsRefresh = [bool]$($true)
    $Script:StatisticsVariables          = [PSCustomObject[]](
        [PSCustomObject]@{'VariableName'='ModulesInstalledMissing';   'Description'='Installed (was missing)'},
        [PSCustomObject]@{'VariableName'='ModulesSubInstalledMissing';'Description'='Installed submodule (was missing)'},
        [PSCustomObject]@{'VariableName'='ModulesUpdated';            'Description'='Updated'},
        [PSCustomObject]@{'VariableName'='ModulesUninstalledOutdated';'Description'='Uninstalled (outdated)'},
        [PSCustomObject]@{'VariableName'='ModulesUninstalledUnwanted';'Description'='Uninstalled (unwanted)'}
    )

    # Set settings
    ## Disable proxies for speed
    [System.Net.WebRequest]::DefaultWebProxy = $null
    ## Use TLS 1.2
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12


    # Prerequirements
    Write-Output -InputObject ('### Install Prerequirements')
    if ([byte]$([string[]]$(Get-Module -Name 'PowerShellGet','PackageManagement' -ListAvailable -ErrorAction 'SilentlyContinue' | Where-Object -Property 'ModuleType' -eq 'Script' | Select-Object -ExpandProperty 'Name' -Unique).'Count') -lt 2) {        
        Write-Output -InputObject ('Prerequirements were not met.')


        # Prerequirement - NuGet (Package Provider)
        Write-Output -InputObject ('# Prerequirement - Package Provider - "NuGet"' -f ("`r`n`r`n"))
        $VersionNuGetMinimum   = [System.Version]$(Find-PackageProvider -Name 'NuGet' -Force -Verbose:$false -Debug:$false | Select-Object -ExpandProperty 'Version')
        $VersionNuGetInstalled = [System.Version]$(Get-PackageProvider -ListAvailable -Name 'NuGet' -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty 'Version' | Sort-Object -Descending | Select-Object -First 1) 
        if ((-not($VersionNuGetInstalled)) -or $VersionNuGetInstalled -lt $VersionNuGetMinimum) {        
            $null = Install-PackageProvider 'NuGet' –Force -Verbose:$false -Debug:$false -ErrorAction 'Stop'
            Write-Output -InputObject ('{0}Not installed, or newer version available. Installing... Success? {1}' -f ("`t",$?.ToString()))
        }
        else {
            Write-Output -InputObject ('{0}NuGet (Package Provider) is already installed.' -f ("`t"))
        }


        # Prerequirement - PowerShellGet (PowerShell Module)
        Write-Output -InputObject ('{0}# Prerequirement - PowerShell Modules - "PackageManagement" and "PowerShellGet"' -f ("`r`n"))
        $ModulesRequired = [string[]]@('PackageManagement','PowerShellGet')
        foreach ($ModuleName in $ModulesRequired) {
            Write-Output -InputObject ('{0}' -f ($ModuleName))
            $VersionModuleAvailable = [System.Version]$(Get-ModulePublishedVersion -ModuleName $ModuleName)
            $VersionModuleInstalled = [System.Version]$(Get-InstalledModule -Name $ModuleName -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty 'Version')
            if ((-not($VersionModuleInstalled)) -or $VersionModuleInstalled -lt $VersionModuleAvailable) {           
                Write-Output -InputObject ('{0}Not installed, or newer version available. Installing...' -f ("`t"))
                $null = Install-Module -Name $ModuleName -Scope 'AllUsers' -Verbose:$false -Debug:$false -Confirm:$false -Force -ErrorAction 'Stop'
                $Success = [bool]$($?)
                Write-Output -InputObject ('{0}{0}Success? {1}' -f ("`t",$Success.ToString()))
                if ($Success) {
                    $null = Import-Module -Name $ModuleName -RequiredVersion $VersionModuleAvailable -Force -ErrorAction 'Stop'
                }
            }
            else {
                Write-Output -InputObject ('{0}"{1}" (PowerShell Module) is already installed.' -f ("`t",$ModuleName))
            }
        }


        # Exit script, tell user to restart PowerShell session
        Write-Output -InputObject 'Prerequirements are now installed. Please restart current PowerShell session before running script again.'
        Exit 0
    }
    else {
        Write-Output -InputObject ('Prerequirements were met.')
    }



    # Only continue if PowerShellGet is installed and can be imported successfully
    if (-not([bool]$($null = Import-Module -Name 'PowerShellGet' -Force -ErrorAction 'SilentlyContinue';$?))) {
        Throw 'ERROR: PowerShell module "PowerShellGet" is required to continue.'
    }



    # Uninstall Unwanted Modules
    Write-Output -InputObject ('{0}### Uninstall Unwanted Modules' -f ("`r`n`r`n"))
    if ($UninstallUnwantedModules) {
        Uninstall-ModulesUnwanted -ModulesUnwanted $ModulesUnwanted
    }
    else {
        Write-Output -InputObject ('{0}Uninstall Unwanted Modules is set to $false.' -f ("`t"))
    }
    


    # Update Installed Modules
    Write-Output -InputObject ('{0}### Update Installed Modules' -f ("`r`n`r`n"))
    if ($InstallUpdatedModules) {
        Update-ModulesInstalled
    }
    else {
        Write-Output -InputObject ('{0}Update Installed Modules is set to $false.' -f ("`t"))
    }



    # Install Missing Modules
    Write-Output -InputObject ('{0}### Install Missing Modules' -f ("`r`n`r`n"))
    if ($InstallMissingModules) {
        Install-ModulesMissing -ModulesWanted $ModulesWanted
    }
    else {
        Write-Output -InputObject ('{0}Install Missing Modules is set to $false.' -f ("`t"))
    }



    # Installing Missing Sub Modules
    Write-Output -InputObject ('{0}### Install Missing Sub Modules' -f ("`r`n`r`n"))
    if ($InstallMissingSubModules) {
        Install-SubModulesMissing
    }
    else {
        Write-Output -InputObject ('{0}Install Missing Sub Modules is set to $false.' -f ("`t"))
    }

    

    # Remove old modules
    Write-Output -InputObject ('{0}### Remove Outdated Modules' -f ("`r`n`r`n"))
    if ($UninstallOutdatedModules) {
        Uninstall-ModulesOutdated
    }
    else {
        Write-Output -InputObject ('{0}Remove Outdated Modules is set to $false.' -f ("`t"))
    }



    # Write Stats
    Write-Output -InputObject ('{0}### Finished.' -f ("`r`n"))
    Write-Output -InputObject ('#### Stats')
    Output-Statistics
    Write-Output -InputObject ('{0}#### Time' -f ("`r`n"))
    Write-Output -InputObject ('Start Time:    {0} ({1}).' -f ($Script:TimeTotalStart.ToString('HH\:mm\:ss'),$Script:TimeTotalStart.ToString('o')))
    Write-Output -InputObject ('End Time:      {0} ({1}).' -f (($Script:TimeTotalEnd = [datetime]::Now).ToString('HH\:mm\:ss'),$Script:TimeTotalEnd.ToString('o')))
    Write-Output -InputObject ('Total Runtime: {0}.' -f ([string]$([timespan]$($Script:TimeTotalEnd - $Script:TimeTotalStart)).ToString('hh\:mm\:ss')))
#endregion Main
}