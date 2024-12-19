#Requires -Version 5.1
<#
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
        Created:  2019-03-10
        Modified: 2024-12-19

    .EXAMPLE
        # Run from PowerShell ISE or Visual Studio Code, user context
        Set-ExecutionPolicy -Scope 'Process' -ExecutionPolicy 'Bypass' -Force
        & $(Try{$psEditor.GetEditorContext().CurrentFile.Path}Catch{$psISE.CurrentFile.FullPath}) -SystemContext $false -DevDrive 'D'

    .EXAMPLE
        # Run from PowerShell ISE, system context, bypass script execution policy
        Set-ExecutionPolicy -Scope 'Process' -ExecutionPolicy 'Bypass' -Force
        & $(Try{$psEditor.GetEditorContext().CurrentFile.Path}Catch{$psISE.CurrentFile.FullPath}) -SystemContext $true
#>



# Input parameters and expected output
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'SkipAuthenticodeCheck', Justification = 'False positive.')]
[OutputType([System.Void])]
Param (
    [Parameter(HelpMessage = 'System/device context, else current user only.')]
    [bool] $SystemContext = $false,

    [Parameter(HelpMessage = 'Whether to accept licenses when installing modules that requires it.')]
    [bool] $AcceptLicenses = $true,

    [Alias('SkipPublisherCheck')]
    [Parameter(HelpMessage = 'Security, whether to skip checking signing of the module against alleged publisher.')]
    [bool] $SkipAuthenticodeCheck = $false,

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
    [bool] $DoScripts = $true,

    [Parameter(HelpMessage = 'Windows only! Whether to use DevDrive as temp, if yes provide label, ex "D".')]
    [ValidateScript({
            $PSVersionTable.'Platform' -ne 'Unix' -and
            $_.'Length' -eq 1 -and
            $_ -match '(?i)^[a-z]{1}$'
            [System.IO.Directory]::Exists(($_+':\')) -and
            (Get-Volume -DriveLetter $_).'FileSystem' -eq 'ReFS'
        })
    ]
    [string] $DevDrive
)



#region    Settings & Variables
# List of modules
## Modules you want to install and keep installed
$ModulesWanted = [string[]](
    'AIPService',                             # Microsoft. Used for managing Microsoft Azure Information Protection (AIP).
    'Az',                                     # Microsoft. Used for Azure Resources. Combines and extends functionality from AzureRM and AzureRM.Netcore.
    'AzPolicyTest',                           # Tao Yang. Used for validating Azure policies with Pester.
    'AzSK',                                   # Microsoft. Azure Secure DevOps Kit. https://azsk.azurewebsites.net/00a-Setup/Readme.html
    'AzViz',                                  # Prateek Singh. Used for visualizing a Azure environment.
    'AWSPowerShell.NetCore',                  # Amazon AWS.
    'Bicep',                                  # PSBicep. Provides the same functionality as Bicep CLI, plus some additional features to simplify the Bicep authoring experience.
    'BuildHelpers',                           # Warren Frame. Helper functions for PowerShell CI/CD scenarios.
    'ConfluencePS',                           # Atlassian, for interacting with Atlassian Confluence Rest API.
    'ConvertToSARIF',                         # Microsoft. A cmdlet for converting PSScriptAnalyzer output to the SARIF format.
    'DefenderMAPS',                           # Alex Verboon, for testing connectivity to "MAPS connection for Microsoft Windows Defender".
    'Evergreen',                              # By Aaron Parker @ Stealth Puppy. For getting URL etc. to latest version of various applications on Windows.
    'EnterprisePolicyAsCode',                 # Microsoft. EPAC, Enterprise Azure Policy as Code.
    'ExchangeOnlineManagement',               # Microsoft. Used for managing Exchange Online.
    'GetBIOS',                                # Damien Van Robaeys. Used for getting BIOS settings for Lenovo, Dell and HP.
    'ImportExcel',                            # dfinke.    Used for import/export to Excel.
    'Intune.USB.Creator',                     # Ben Reader @ powers-hell.com. Used to create Autopilot WinPE.
    'IntuneBackupAndRestore',                 # John Seerden. Uses "MSGraphFunctions" module to backup and restore Intune config.
    'Invokeall',                              # Santhosh Sethumadhavan. Multithread PowerShell commands.
    'JWTDetails',                             # Darren J. Robinson. Used for decoding JWT, JSON Web Tokens.
    'Mailozaurr',                             # Przemyslaw Klys. Used for various email related tasks.
    'MarkdownPS',                             # Alex Sarafian. Generate Markdown.
    'Microsoft.Graph',                        # Microsoft. Works with PowerShell Core.
    'Microsoft.Graph.Beta',                   # Microsoft. Works with PowerShell Core.
    'Microsoft.Online.SharePoint.PowerShell', # Microsoft. Used for managing SharePoint Online.
    'Microsoft.PowerShell.ConsoleGuiTools',   # Microsoft.
    'Microsoft.PowerShell.PSResourceGet',     # Microsoft, successor to PowerShellGet and PackageManagement.
    'Microsoft.PowerShell.SecretManagement',  # Microsoft. Used for securely managing secrets.
    'Microsoft.PowerShell.SecretStore',       # Microsoft. Used for securely storing secrets locally.
    'Microsoft.PowerShell.ThreadJob',         # Microsoft. Successfor of "ThreadJob" originally by Paul Higinbotham.
    'Microsoft.RDInfra.RDPowerShell',         # Microsoft. Used for managing Windows Virtual Desktop.
    'Microsoft.WinGet.Client',                # Microsoft. Used for running WinGet commands with a dedicated PowerShell module, vs. using the CLI.
    #'Microsoft.WinGet.Configuration',         # Microsoft. Used to configure Winget.
    'MicrosoftGraphSecurity',                 # Microsoft. Used for interacting with Microsoft Graph Security API.
    'MicrosoftPowerBIMgmt',                   # Microsoft. Used for managing Power BI.
    'MicrosoftTeams',                         # Microsoft. Used for managing Microsoft Teams.
    'ModuleFast',                             # Justin Grote. Module to install PowerShell modules fast.
    'MSAL.PS',                                # Microsoft. Used for MSAL authentication.
    'MSGraphFunctions',                       # John Seerden. Wrapper for Microsoft Graph Rest API.
    'MSOnline',                               # (DEPRECATED, "AzureAD" is it's successor) Microsoft. Used for managing Microsoft Cloud Objects (Users, Groups, Devices, Domains...)
    'Nevergreen',                             # Dan Gough. Evergreen alternative that scrapes websites for getting latest version and URL to a package.
    'newtonsoft.json',                        # Serialize/Deserialize Json using Newtonsoft.json
    'Office365DnsChecker',                    # Colin Cogle. Checks a domain's Office 365 DNS records for correctness.
    'Optimized.Mga',                          # Bas Wijdenes. Microsoft Graph batch operations.
    'PartnerCenter',                          # Microsoft. Used for interacting with PartnerCenter API.
    'Pester',                                 # Pester. Ubiquitous test and mock framework for PowerShell.
    'platyPS',                                # Microsoft. Used for converting markdown to PowerShell XML external help files.
    'PnP.PowerShell',                         # Microsoft. Used for managing SharePoint Online.
    'PolicyFileEditor',                       # Microsoft. Used for local group policy / gpedit.msc.
    'PoshRSJob',                              # Boe Prox. Used for parallel execution of PowerShell.
    'powershell-yaml',                        # Cloudbase. Serialize and deserialize YAML, using YamlDotNet.
    'PSIntuneAuth',                           # Nickolaj Andersen. Get auth token to Intune.
    'PSPackageProject',                       # Microsoft. Help with building and publishing PowerShell packages.
    'PSPKI',                                  # Vadims Podans. Used for infrastructure and certificate management.
    'PSReadLine',                             # Microsoft. Used for helping when scripting PowerShell.
    'PSRule',                                 # Microsoft. Validate infrastructure as code (IaC) and objects using PowerShell rules.
    'PSRule.Rules.Azure',                     # Microsoft. PSRule rules for Azure.
    'PSScriptAnalyzer',                       # Microsoft. Used to analyze PowerShell scripts to look for common mistakes + give advice.
    'PSSemVer',                               # Marius Storhaug. Semantic Versioning in PowerShell using a class.
    'PSSendGrid',                             # Barbara Forbes. Used to send emails with Send Grid.
    'PSWindowsUpdate',                        # Michal Gajda. Used for updating Windows.
    'RunAsUser',                              # Kelvin Tegelaar. Allows running as current user while running as SYSTEM using impersonation.
    'SCEPman',                                # SCEPman. Used for managing SCEPman.
    'Scoop',                                  # Thomas Nieto. PowerShell module for Scoop.
    'SetBIOS',                                # Damien Van Robaeys. Used for setting BIOS settings for Lenovo, Dell and HP.
    'SharePointPnPPowerShellOnline',          # Microsoft. Used for managing SharePoint Online.
    'SpeculationControl',                     # Microsoft, by Matt Miller. To query speculation control settings (Meltdown, Spectr).
    'TaskJob',                                # Justin Grote. Enables you to take any .NET Task object and treat it like a PowerShell Job.
    'VcRedist',                               # Aaron Parker. Lifecycle management for the Microsoft Visual C++ Redistributables.
    'VSTeam',                                 # Donovan Brown. Adds functionality for working with Azure DevOps and Team Foundation Server.
    'WindowsAutoPilotIntune',                 # Michael Niehaus @ Microsoft. Used for Intune AutoPilot stuff.
    'WinSCP'                                  # Thomas Malkewitz. WinSCP PowerShell Wrapper Module.
)

## Modules you don't want - Will Remove Every Related Module, for AzureRM for instance will also search for AzureRM.*
$ModulesUnwanted = [string[]](
    'AnyPackage',                             # AnyPackage / Thomas Nieto. Spiritual successor to OneGet / PackageManagement.
    'AnyPackage.PSResourceGet',               # AnyPackage / Thomas Nieto. PSResourceGet for AnyPackage.
    'Az.Insights',                            # Name changed to "Az.Monitor": https://learn.microsoft.com/en-us/powershell/azure/migrate-az-1.0.0#module-name-changes
    'Az.Profile',                             # Name changed to "Az.Accounts": https://learn.microsoft.com/en-us/powershell/azure/migrate-az-1.0.0#module-name-changes
    'Az.Tags',                                # Functionality merged into "Az.Resources": https://learn.microsoft.com/en-us/powershell/azure/migrate-az-1.0.0#module-name-changes
    'Az.Tools',                               # I don't use any of them, and Az.Tools.Predictor randomly crashed vscode-powershell terminal.
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
    #'Az.Accounts'
)

## Module versions you don't want removed
$ModulesVersionsDontRemove = [ordered]@{
    #'Az.Accounts' = [System.Version[]]('2.19.0')
}



# List of wanted scripts
$ScriptsWanted = [string[]](
    'Get-WindowsAutoPilotInfo',                # Microsoft, Michael Niehaus. Get Windows AutoPilot info.
    'Get-AutopilotDiagnostics',                # Microsoft, Michael Niehaus. Display diagnostics information.
    'Upload-WindowsAutopilotDeviceInfo',       # Nickolaj Andersen. Upload autopilot hash straight to Intune.
    'winfetch'                                 # Rashil Gandhi. Like neofetch, just for Windows and PowerShell.
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
$null = Set-Variable -Option 'AllScope', 'ReadOnly' -Force -Name 'Indentation' -Value ([string]'  ')
#endregion Settings & Variables




#region    Functions
#region    PowerShell Gallery
#region    Find-PSGalleryPackageLatestVersionUsingApiInBatch
function Find-PSGalleryPackageLatestVersionUsingApiInBatch {
    <#
        .SYNOPSIS
            Get PowerShell Gallery package version info for one or multiple packages using the API directly.

        .NOTES
            Author:   Olav Rønnestad Birkeland | github.com/o-l-a-v
            Created:  240313
            Modified: 240330

        .EXAMPLE
            Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds 'Az.Accounts','Az.Resources' -Verbose
            Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds 'Az.CosmosDB' -IncludePrerelease -Verbose -As
            Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds 'Az.*' -MinimalInfo -Verbose
            Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds 'Az.*' -AsHashtable -Verbose
            Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds 'Az.*' -AsHashtable -MinimalInfo -Verbose
    #>

    # Input parameters and expected output
    [CmdletBinding(DefaultParameterSetName='Stable')]
    [CmdletBinding()]
    [OutputType([hashtable],[PSCustomObject])]
    Param(
        [Parameter(Mandatory, ParameterSetName = 'Absolute')]
        [Parameter(Mandatory, ParameterSetName = 'Stable')]
        [Parameter(Mandatory)]
        $PackageIds,

        [Alias('Absolute','Prerelease')]
        [Parameter(HelpMessage = 'Include prerelease.', ParameterSetName = 'Absolute')]
        [switch] $IncludePrerelease,

        [Parameter(HelpMessage = 'Optionally return result as hashtable.', ParameterSetName = 'Absolute')]
        [Parameter(HelpMessage = 'Optionally return result as hashtable.', ParameterSetName = 'Stable')]
        [switch] $AsHashtable,

        [Parameter(HelpMessage = 'Strip some info to speed up interaction with PowerShell Gallery API.', ParameterSetName = 'Absolute')]
        [Parameter(HelpMessage = 'Strip some info to speed up interaction with PowerShell Gallery API.', ParameterSetName = 'Stable')]
        [switch] $MinimalInfo
    )

    # Process
    Process {
        # Parse and validate input
        ## Is of expected type - String or string array
        if ($PackageIds -is [string] -or $PackageIds.ForEach{$_ -is [string]} -notcontains $false) {
            Write-Verbose -Message 'Input seems legit.'
        }
        else {
            Throw [ArgumentException]::new('Input is not a string or string array.')
        }

        ## Cast input to string array, remove duplicates and sort alphabetically
        $PackageIdsFiltered = [string[]]($PackageIds -as [string[]] | Where-Object -FilterScript {-not [string]::IsNullOrEmpty($_)} | Sort-Object -Unique)

        ## Verify we still have any input
        if ($PackageIdsFiltered.Where{-not[string]::IsNullOrEmpty($_)}.'Count' -le 0) {
            Throw [ArgumentException]::new('Input seems to be empty.')
        }

        ## Verify that input is as expected
        if ($PackageIdsFiltered.Where({$_.'Length' -gt 90 -or $_ -notmatch '^\*$|^(\*?)([a-zA-Z0-9.\-_]{1,})(\*?)$'},'First').'Count' -gt 0) {
            Throw [ArgumentException]::new('Input contains unexpected characters or is too long (>90ch).')
        }

        ## Output warning if duplicates was found
        if ($PackageIdsFiltered.'Count' -lt $PackageIds.'Count') {
            Write-Warning -Message 'Input $PackageIds had duplicates / multiple items with the same value.'
        }


        # Create API filter based on parameter set name
        $VersionFilter = [string](
            $(
                if ($IncludePrerelease.'IsPresent') {
                    'IsAbsoluteLatestVersion'
                }
                else {
                    'IsLatestVersion and not IsPrerelease'
                }
            )
        )
        if ($MinimalInfo.'IsPresent') {
            $Select = [string] '&$select=Authors,Published,RequireLicenseAcceptance,Tags,Version' -f $Filter
        }


        # Skip batching if input is only one module
        if ($PackageIdsFiltered.'Count' -eq 1 -and $PackageIdsFiltered.Where{$_.Contains('*')}.'Count' -le 0) {
            Write-Verbose -Message 'Skipping batching because input is one specific package ID.'
            $Results = [System.Xml.XmlElement[]](
                $PackageIdsFiltered.ForEach{
                    $Uri = [string](
                        "https://www.powershellgallery.com/api/v2/FindPackagesById()?id='{0}'&`$filter={1}{2}&semVerLevel=1.0.0" -f (
                            $_,
                            $VersionFilter,
                            $Select
                        )
                    )
                    Write-Verbose -Message $Uri
                    Invoke-RestMethod -Method 'Get' -Uri $Uri
                }
            )
        }


        # Else - Do batching, paging etc.
        else {
            # Assets
            ## PowerShell version - To enable version specific features
            $IsPowerShell72 = [bool]($PSVersionTable.'PSVersion' -ge $([System.Version]('7.2')))
            $IsPowerShell74 = [bool]($PSVersionTable.'PSVersion' -ge $([System.Version]('7.4')))

            ## PowerShell Gallery API
            $ApiPageSize = [byte] 100
            $Headers = [ordered]@{
                'Accept'          = [string] 'application/atom+xml;charset=UTF-8'
                'Accept-Encoding' = [string] 'gzip, deflate'
            }
            if ($IsPowerShell74) {
                $Headers.'Accept-Encoding' = [string] '{0}, br' -f $Headers.'Accept-Encoding'
            }

            ## Batching of input
            $ArrayLastIndex = [uint16]($PackageIdsFiltered.'Count'-1)
            $ArrayPage = [byte] 0
            $ArrayPageSize = [byte] 30

            ## Results
            $Results = [System.Collections.Generic.List[System.Xml.XmlElement]]::new()


            # Find packages
            Write-Verbose -Message ('Headers:{0}' -f (ConvertTo-Json -Depth 1 -InputObject $Headers -Compress))
            do {
                $ArrayCurrentLastIndex = [uint16]($([uint16[]]($ArrayLastIndex,(($ArrayPage*$ArrayPageSize)+$ArrayPageSize-1) | Sort-Object))[0])
                $ApiPage = [byte] 0
                do {
                    # Create variable to splat into Invoke-RestMethod
                    $Splat = [ordered]@{
                        'Headers' = $Headers
                        'Method'  = [string] 'Get'
                        'Uri'     = [string](
                            ('https://www.powershellgallery.com/api/v2/Packages?$filter={0} and (' -f $VersionFilter) +
                            (
                                $PackageIdsFiltered[($ArrayPage*$ArrayPageSize) .. $ArrayCurrentLastIndex].ForEach{
                                    if ($_.EndsWith('*') -and $_.StartsWith('*')) {
                                        "substringof('{0}',Id)" -f $_.Replace('*','')
                                    }
                                    elseif ($_.EndsWith('*')) {
                                        "startswith(Id,'{0}')" -f $_.Replace('*','')
                                    }
                                    elseif ($_.StartsWith('*')) {
                                        "endswith(Id,'{0}')" -f $_.Replace('*','')
                                    }
                                    else {
                                        "Id eq '{0}'" -f $_
                                    }
                                } -join ' or '
                            ) + ('){0}&semVerLevel=1.0.0$inlinecount=allpages&$skip={1}&$top={2}' -f $Select, ($ApiPage*$ApiPageSize), $ApiPageSize)
                        )
                    }
                    # Use retry and WebSession if PowerShell >= 7.2
                    if ($IsPowerShell72) {
                        if ($ApiPage -le 0 -and $ArrayPage -le 0) {
                            $Splat.'SessionVariable' = [string] 'WebSession'
                        }
                        else {
                            $Splat.'WebSession' = $WebSession
                        }
                        $Splat.'HttpVersion'       = [System.Version] '2.0'
                        $Splat.'MaximumRetryCount' = [int32] 2
                        $Splat.'RetryIntervalSec'  = [int32] 1
                    }
                    # Do the request
                    Write-Verbose -Message ('{0}:{1}' -f ($ArrayPage+1).ToString('00'), $Splat.'Uri')
                    $Response = [System.Xml.XmlElement[]](Invoke-RestMethod @Splat)
                    if ($Response.'Count' -gt 0) {
                        $Results.AddRange($Response)
                    }
                    $ApiPage++
                }
                while ($Response.'Count' -gt 0 -and $Response.'Count' % $ApiPageSize -eq 0)
                $ArrayPage++
            }
            while ($ArrayCurrentLastIndex -lt $ArrayLastIndex)
        }


        # Parse and create output
        Write-Verbose -Message 'Parse and create output'
        $OutputAsPSCustomObjectArray = [PSCustomObject[]](
            $(
                $Results.Where{
                    -not [string]::IsNullOrEmpty($_.'Id')
                }.ForEach{
                    [PSCustomObject]@{
                        'Name'                     = [string] $_.'title'.'#text'
                        'Author'                   = [string] $_.'author'.'name'
                        'Depencencies'             = [string[]](
                            $(
                                if (-not [string]::IsNullOrEmpty($_.'properties'.'Dependencies')) {
                                    $_.'properties'.'Dependencies'.Split(
                                        '|', [StringSplitOptions]::RemoveEmptyEntries
                                    ).ForEach{
                                        $_.Split(':', [StringSplitOptions]::RemoveEmptyEntries)[0]
                                    } | Sort-Object -Unique
                                }
                            )
                        )
                        'Owners'                   = [string] $_.'properties'.'owners'
                        'PublishedDate'            = [nullable[datetime]] $_.'properties'.'Published'.'#text'
                        'RequireLicenseAcceptance' = [bool]($_.'properties'.'RequireLicenseAcceptance'.'#text' -eq 'true')
                        'Tags'                     = [string[]](
                            $(
                                if ($_.'properties'.'Tags' -is [System.Xml.XmlLinkedNode]) {
                                    $_.'properties'.'Tags'.'#text'
                                }
                                else {
                                    $_.'properties'.'Tags'
                                }
                            ).Split(' ', [StringSplitOptions]::RemoveEmptyEntries) | Sort-Object -Unique)
                        'Unlisted'                 = [bool]$(
                            [string]::IsNullOrEmpty($_.'properties'.'Published'.'#text') -or
                            $([datetime]($_.'properties'.'Published'.'#text')).'Year' -le 1900
                        )
                        'Version'                  = [System.Version] $_.'properties'.'Version'.Split('-')[0]
                        'Prerelease'               = [string]$(if($_.'properties'.'Version'.Contains('-')){$_.'properties'.'Version'.Split('-')[-1]})
                        'VersionAsString'          = [string] $_.'properties'.'Version'
                    }
                }
            ) | Sort-Object -Property 'Name'
        )

        # Return
        if ($AsHashtable) {
            Write-Verbose -Message 'Return output as hashtable'
            $OutputAsHashtable = [hashtable]@{}
            $OutputAsPSCustomObjectArray.ForEach{
                $OutputAsHashtable.Add($_.'Name',$_)
            }
            $OutputAsHashtable
        }
        else {
            Write-Verbose -Message 'Return output as PSCustomObject'
            $OutputAsPSCustomObjectArray.ForEach{
                $_
            }
        }
    }
}
#endregion Find-PSGalleryPackageLatestVersionUsingApiInBatch

#region    Save-PSResourceInParallel
function Save-PSResourceInParallel {
    <#
        .SYNOPSIS
            Speed up PSResourceGet\Save-PSResource by parallizing using PowerShell native runspace factory.

        .NOTES
            Author:   Olav Rønnestad Birkeland | github.com/o-l-a-v
            Created:  2023-11-16
            Modified: 2024-06-28

        .EXAMPLE
            . $psEditor.GetEditorContext().CurrentFile.Path
            # All Az modules
            Save-PSResourceInParallel `
                -Name (Find-PSResource -Repository 'PSGallery' -Type 'Module' -Name 'Az').'Dependencies'.'Name' `
                -Path ([System.IO.Path]::Combine(([System.Environment]::GetFolderPath('Desktop')),'TestPS'))
            # AWSPowerShell.NetCore
            Save-PSResourceInParallel `
                -Name 'AWSPowerShell.NetCore' `
                -Path ([System.IO.Path]::Combine(([System.Environment]::GetFolderPath('Desktop')),'TestPS'))
    #>

    # Input parameters and expected output
    [CmdletBinding()]
    [OutputType([Microsoft.PowerShell.PSResourceGet.UtilClasses.PSResourceInfo[]])]
    Param(
        [Parameter(Mandatory, HelpMessage = 'Name of the resource(s) you want saved.')]
        [string[]] $Name,

        [Parameter(HelpMessage = 'Whether to include XML.')]
        [bool] $IncludeXml = [bool] $true,

        [Parameter(Mandatory, HelpMessage = 'Where to save the resource to.')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({(Test-Path -Path $_ -IsValid -PathType 'Container') -and [System.IO.Directory]::Exists($_)})]
        [string] $Path,

        [Parameter(HelpMessage = 'Where to import module "Microsoft.PowerShell.PSResourceGet" from.')]
        [ValidateNotNullOrEmpty()]
        [string] $PSResourceGetPath = (Get-Module -Name 'Microsoft.PowerShell.PSResourceGet' | Select-Object -First 1 -ExpandProperty 'Path'),

        [Parameter(HelpMessage = 'What PSResourceGet repository to use.')]
        [ValidateNotNullOrEmpty()]
        [string] $Repository = 'PSGallery',

        [Parameter(HelpMessage = 'Whether to skip dependency check.')]
        [bool] $SkipDependencyCheck = [bool] $true,

        [Parameter(HelpMessage = 'Override temporary path.')]
        [string] $TemporaryPath,

        [Parameter(HelpMessage = 'Maximum concurrent jobs to run simultaneously.')]
        [byte] $ThrottleLimit = 10,

        [Parameter()]
        [bool] $TrustRepository = [bool] $true
    )


    # Begin
    Begin {
        # Assets
        $ScriptBlock = [scriptblock]{
            [OutputType([System.Void])]
            Param(
                [Parameter()]
                [bool] $IncludeXml = $true,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string] $Name,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string] $Path,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string] $PSResourceGetPath,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string] $Repository,

                [Parameter()]
                [bool] $SkipDependencyCheck = $true,

                [Parameter()]
                [string] $TemporaryPath,

                [Parameter()]
                [bool] $TrustRepository = $true
            )
            $ErrorActionPreference = 'Stop'
            $null = Import-Module -Name $PSResourceGetPath
            $Splat = [ordered]@{
                'AuthenticodeCheck'   = [bool] -not $SkipAuthenticodeCheck
                'IncludeXml'          = [bool] $IncludeXml
                'Name'                = [string] $Name
                'Repository'          = [string] $Repository
                'Path'                = [string] $Path
                'SkipDependencyCheck' = [bool] $SkipDependencyCheck
                'TrustRepository'     = [bool] $TrustRepository
            }
            if ($TemporaryPath) {
                $Splat.'TemporaryPath' = [string] $TemporaryPath
            }
            Microsoft.PowerShell.PSResourceGet\Save-PSResource @Splat
        }

        # Initilize runspace pool
        $RunspacePool = [runspacefactory]::CreateRunspacePool(1,$ThrottleLimit)
        $RunspacePool.Open()
    }


    # Process
    Process {
        # Start jobs in the runspace pool
        $RunspacePoolJobs = [PSCustomObject[]](
            $(
                foreach ($ModuleName in $Name) {
                    $Parameters = [ordered]@{
                        'IncludeXml'          = [bool] $IncludeXml
                        'Name'                = [string] $ModuleName
                        'Path'                = [string] $Path
                        'PSResourceGetPath'   = [string] $PSResourceGetPath
                        'Repository'          = [string] $Repository
                        'SkipDependencyCheck' = [bool] $SkipDependencyCheck
                        'TrustRepository'     = [bool] $TrustRepository
                    }
                    if ($TemporaryPath) {
                        $Parameters.'TemporaryPath' = [string] $TemporaryPath
                    }
                    $PowerShellObject = [powershell]::Create().AddScript($ScriptBlock).AddParameters($Parameters)
                    $PowerShellObject.'RunspacePool' = $RunspacePool
                    [PSCustomObject]@{
                        'ModuleName' = $ModuleName
                        'Instance'   = $PowerShellObject
                        'Result'     = $PowerShellObject.BeginInvoke()
                    }
                }
            )
        )

        # Wait for jobs to finish
        $PrettyPrint = [string]('0'*$RunspacePoolJobs.'Count'.ToString().'Length')
        $WaitIterations = $Completed = $CompletedOld = $Succeeded = [uint16]::MinValue
        while ($RunspacePoolJobs.Where({-not $_.'Result'.'IsCompleted'},'First').'Count' -gt 0) {
            if ($WaitIterations % 5 -eq 0) {
                $Completed = [uint16]($RunspacePoolJobs.Where{$_.'Result'.'IsCompleted'}.'Count')
                $Succeeded = [uint16]($RunspacePoolJobs.Where{$_.'Result'.'IsCompleted' -and -not $_.'Instance'.'HadErrors'}.'Count')
                $StatusChanged = [bool]($Completed -gt $CompletedOld)
                $Message = [string](
                    '{0} / {1} jobs finished, {2} / {0} was successfull.' -f (
                        $Completed.ToString($PrettyPrint),
                        $RunspacePoolJobs.'Count'.ToString(),
                        $Succeeded.ToString($PrettyPrint)
                    )
                )
                if ($PSBoundParameters.'Keys'.Contains('Verbose')) {
                    Write-Verbose -Message $Message
                }
                elseif ($WaitIterations -le 0 -or $StatusChanged) {
                    $CompletedOld = [uint16] $Completed
                    Write-Information -MessageData $Message
                }
            }
            $WaitIterations++
            Start-Sleep -Milliseconds 100
        }

        # Get success state of jobs
        Write-Verbose -Message (
            $RunspacePoolJobs.ForEach{
                [PSCustomObject]@{
                    'Name'        = [string] $_.'ModuleName'
                    'IsCompleted' = [bool] $_.'Result'.'IsCompleted'
                    'HadErrors'   = [bool] $_.'Instance'.'HadErrors'
                }
            } | Sort-Object -Property 'Name' | Format-Table | Out-String
        )

        # Collect results
        $Results = [Microsoft.PowerShell.PSResourceGet.UtilClasses.PSResourceInfo[]](
            $RunspacePoolJobs.ForEach{
                $_.'Instance'.EndInvoke($_.'Result')
            }
        )
    }


    # End
    End {
        # Terminate runspace pool
        $RunspacePool.Close()
        $RunspacePool.Dispose()

        # Output results
        $Results
    }
}
#endregion Save-PSResourceInParallel
#endregion PowerShell Gallery



#region    Modules
#region    Get-ModuleInstalledVersion
function Get-ModuleInstalledVersion {
    <#
        .SYNOPSIS
            Gets all installed versions of a module.

        .DESCRIPTION
            Gets all installed versions of a module.
            * Includes workaround to handle versions that can't be parsed as [System.Version]

        .PARAMETER ModuleName
            String, name of the module you want to check.
    #>
    [CmdletBinding()]
    [OutputType([System.Version],[System.Version[]])]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $ModuleName,

        [Parameter()]
        [switch] $Latest
    )

    # Begin
    Begin {}

    # Process
    Process {
        # Assets
        $Path = [string] [System.IO.Path]::Combine($Script:ModulesPath,$ModuleName)

        # Get installed versions of $ModuleName
        $AllVersions = [System.Version[]](
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
            } | Sort-Object -Property @{'Expression' = {[System.Version]$_}} -Descending
        )

        # Return
        if ($Latest.'IsPresent' -and $AllVersions.'Count' -ge 1) {
            $AllVersions[0]
        }
        else {
            $AllVersions
        }
    }

    # End
    End {
    }
}
#endregion Get-ModuleInstalledVersion



#region    Get-ModulesInstalled
function Get-ModulesInstalled {
    <#
        .SYNOPSIS
            Gets all currently installed modules.

        .NOTES
            * Includes some workarounds to handle versions that can't be parsed as [System.Version], like beta/ pre-release.

        .EXAMPLE
            Get-ModulesInstalled -ForceRefresh
    #>

    # Input parameters and expected output
    [CmdletBinding()]
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
                Microsoft.PowerShell.PSResourceGet\Get-InstalledPSResource -Path $Script:ModulesPath -ErrorAction 'Ignore' |
                    Where-Object -FilterScript {
                        $_.'Type' -eq 'Module' -and
                        $_.'Repository' -eq 'PSGallery' -and
                        -not ($IncludePreReleaseVersions -and $_.'IsPrerelease')
                    } | ForEach-Object -Process {
                        Add-Member -InputObject $_ -MemberType 'AliasProperty' -Name 'Path' -Value 'InstalledLocation' -PassThru
                    } | Group-Object -Property 'Name' | Select-Object -Property 'Name',@{'Name'='Versions';'Expression'='Group'}
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
        .SYNOPSIS
            Fetches latest version number of a given module from PowerShellGallery.
    #>

    # Input parameters and expected output
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([System.Void])]
    Param()

    # Begin
    Begin {
    }

    # Process
    Process {
        # Refresh Installed Modules variable
        $null = Get-ModulesInstalled

        # Skip if no installed modules was found
        if ($Script:ModulesInstalled.'Count' -le 0) {
            Write-Information -MessageData ('No installed modules where found, no modules to update.')
            return
        }

        # Assets
        Write-Information -MessageData ('Get newest available version for all installed module(s).' -f $Script:ModulesInstalled.'Count'.ToString())
        $ModulesInstalledNewestVersion = Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds $Script:ModulesInstalled.'Name'.Where{$_ -notin $ModulesDontUpdate} -AsHashtable -MinimalInfo
        $ModulesInstalledWithNewerVersionAvailable = [PSCustomObject[]](
            $Script:ModulesInstalled.Where{
                $([System.Version]$_.'Version') -lt $([System.Version]$ModulesInstalledNewestVersion.$($_.'Name').'Version')
            }.ForEach{
                $null = Add-Member -InputObject $_ -MemberType 'NoteProperty' -Force -Name 'VersionAvailable' -Value (
                    $ModulesInstalledNewestVersion.$($_.'Name').'Version' -as [System.Version]
                )
                $_
            } | Sort-Object -Property 'Name'
        )
        Write-Information -MessageData ('{0} module(s) has a newer version available.' -f $ModulesInstalledWithNewerVersionAvailable.'Count'.ToString())
        $ModulesInstalledNotFoundInPSGallery = [PSCustomObject[]]($Script:ModulesInstalled.Where{$_.'Name' -notin $ModulesInstalledNewestVersion.'Keys'})
        if ($ModulesInstalledNotFoundInPSGallery.'Count' -gt 0) {
            Write-Warning -WarningAction 'Continue' -Message ('Following {0} installed module(s) was not found in PSGallery:' -f $ModulesInstalledNotFoundInPSGallery.'Count'.ToString())
            Write-Warning -WarningAction 'Continue' -Message ($Indentation + ($ModulesInstalledNotFoundInPSGallery.'Name' | Sort-Object | Join-String -Separator ', '))
        }

        # Don't continue if no modules requires update
        if ($ModulesInstalledWithNewerVersionAvailable.'Count' -le 0) {
            return
        }

        # Update outdated modules
        if ($PSCmdlet.ShouldProcess(('{0} modules' -f $ModulesInstalledWithNewerVersionAvailable.'Count'.ToString()), 'update')) {
            Write-Information -MessageData ('Updating {0} outdated module(s) in parallel.' -f $ModulesInstalledWithNewerVersionAvailable.'Count'.ToString())
            $Splat = [ordered]@{
                'IncludeXml'          = [bool] $true
                'Name'                = [string[]] $ModulesInstalledWithNewerVersionAvailable.'Name'
                'Path'                = [string] $Script:ModulesPath
                'Repository'          = [string] 'PSGallery'
                'SkipDependencyCheck' = [bool] $true
                'TrustRepository'     = [bool] $true
                'ThrottleLimit'       = [byte] 16
            }
            if (-not [string]::IsNullOrEmpty($Script:TemporaryPath)) {
                $Splat.'TemporaryPath' = [string] $Script:TemporaryPath
            }
            $null = Save-PSResourceInParallel @Splat

            # Check success
            $ModulesInstalledWithNewerVersionAvailable.ForEach{
                $null = Add-Member -InputObject $_ -MemberType 'NoteProperty' -Force -Name 'Success' -Value (
                    [System.IO.Directory]::Exists(
                        [System.IO.Path]::Combine($Script:ModulesPath, $_.'Name', $_.'Version'.ToString())
                    )
                )
            }
            $Success = [bool]($ModulesInstalledWithNewerVersionAvailable.Where{-not $_.'Success'}.'Count' -le 0)
            Write-Information -MessageData (
                'Success? {0}. {1} of {2} installed successfully.' -f (
                    $Success.ToString(),
                    $ModulesInstalledWithNewerVersionAvailable.Where{$_.'Success'}.'Count'.ToString(),
                    $ModulesInstalledWithNewerVersionAvailable.'Count'.ToString()
                )
            )

            # Output modules not installed
            if (-not $Success) {
                $FailedToInstall = [PSCustomObject[]](
                    $ModulesInstalledWithNewerVersionAvailable.Where{-not $_.'Success'}
                )
                Write-Warning -Message ('Following {0} submodule(s) failed to install' -f $FailedToInstall.'Count'.ToString())
                Write-Warning -Message ($FailedToInstall.'Name' | Sort-Object -Unique | Join-String -Separator ', ')
            }

            # If at least one module was successfully installed
            if ($ModulesInstalledWithNewerVersionAvailable.Where{$_.'Success'}.'Count' -gt 0) {
                # Stats
                $Script:ModulesUpdated += [string[]]$($ModulesInstalledWithNewerVersionAvailable.Where{$_.'Success'}.'Name')
                # Make sure list of installed modules gets refreshed
                $Script:ModulesInstalledNeedsRefresh = $true
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
        .SYNOPSIS
            Installs missing modules by comparing installed modules vs input parameter $ModulesWanted.

        .PARAMETER ModulesWanted
            A string array containing names of wanted modules.
    #>

    # Input parameters and expected output
    [CmdletBinding()]
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

        # Find missing modules
        $ModulesMissing = [string[]]($ModulesWanted.Where{$_ -notin $Script:ModulesInstalled.'Name'})
        Write-Information -MessageData ('Found {0} missing module(s).' -f $ModulesMissing.'Count'.ToString())
        if ($ModulesMissing.'Count' -le 0) {
            return
        }

        # Find missing modules in the PowerShell Gallery
        Write-Information -MessageData 'Find info on all missing modules and dependencies.'
        $ModulesMissingPsgInfo = [PSCustomObject[]](
            Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds $ModulesMissing
        )

        # Failproof
        ## None of the missing modules where found in the PowerShell Gallery = Crash
        if ($ModulesMissingPsgInfo.'Count' -le 0) {
            Write-Error -Message ('Did not find any of the following missing modules in PowerShell Gallery: "{0}"-' -f ($ModulesMissingNotFound -join '", "'))
        }
        ## Some of the missing modules where not found in the PowerShell Gallery = Warn
        $ModulesMissingNotFound = [string[]](
            $ModulesMissing.Where{$_ -notin $ModulesMissingPsgInfo.'Name'} | Sort-Object
        )
        if ($ModulesMissingNotFound.'Count' -gt 0) {
            Write-Warning -Message ('Did not find following missing modules in PowerShell Gallery: "{0}".' -f ($ModulesMissingNotFound -join '", "'))
        }

        # Get more info on the missing modules
        $ModulesMissingPsgInfo = [PSCustomObject[]](
            Find-PSGalleryPackageLatestVersionUsingApiInBatch -MinimalInfo -PackageIds (
                $(
                    [string[]](
                        $ModulesMissingPsgInfo.'Name' + $(
                            foreach ($Module in $ModulesMissingPsgInfo) {
                                $Module.'Depencencies'.Where{$_.StartsWith('{0}.' -f $Module.'Name')}
                            }
                        )
                    ) | Sort-Object -Unique | Where-Object -FilterScript {$_ -notin $Script:ModulesInstalled.'Name'}
                )
            )
        )
        Write-Information -MessageData (
            'Found a total of {0} module(s) that is missing and will be installed.' -f (
                $ModulesMissingPsgInfo.Where{-not $_.'Installed'}.'Count'.ToString()
            )
        )

        # Install missing modules
        $Splat = [ordered]@{
            'IncludeXml'          = [bool] $true
            'Name'                = [string[]] $ModulesMissingPsgInfo.'Name'
            'Path'                = [string] $Script:ModulesPath
            'Repository'          = [string] 'PSGallery'
            'SkipDependencyCheck' = [bool] $true
            'TrustRepository'     = [bool] $true
            'ThrottleLimit'       = [byte] 16
        }
        if (-not [string]::IsNullOrEmpty($Script:TemporaryPath)) {
            $Splat.'TemporaryPath' = [string] $Script:TemporaryPath
        }
        $null = Save-PSResourceInParallel @Splat

        # Check success
        $ModulesMissingPsgInfo.ForEach{
            $null = Add-Member -InputObject $_ -MemberType 'NoteProperty' -Force -Name 'Success' -Value (
                [System.IO.Directory]::Exists(
                    [System.IO.Path]::Combine($Script:ModulesPath, $_.'Name', $_.'Version'.ToString())
                )
            )
        }
        $Success = [bool]($ModulesMissingPsgInfo.Where{-not $_.'Success'}.'Count' -le 0)
        Write-Information -MessageData (
            'Success? {0}. {1} of {2} installed successfully.' -f (
                $Success.ToString(),
                $ModulesMissingPsgInfo.Where{$_.'Success'}.'Count'.ToString(),
                $ModulesMissingPsgInfo.'Count'.ToString()
            )
        )

        # Output modules not installed
        if (-not $Success) {
            $FailedToInstall = [PSCustomObject[]](
                $ModulesMissingPsgInfo.Where{-not $_.'Success'}
            )
            Write-Warning -Message ('Following {0} submodule(s) failed to install' -f $FailedToInstall.'Count'.ToString())
            Write-Warning -Message ($FailedToInstall.'Name' | Sort-Object -Unique | Join-String -Separator ', ')
        }

        # If at least one module was successfully installed
        if ($ModulesMissingPsgInfo.Where{$_.'Success'}.'Count' -gt 0) {
            # Stats
            $Script:ModulesInstalledMissing += [string[]]$($ModulesMissingPsgInfo.Where{$_.'Success'}.'Name')
            # Make sure list of installed modules gets refreshed
            $Script:ModulesInstalledNeedsRefresh = $true
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
        .SYNOPSIS
            Installs eventually missing submodules
    #>

    # Input parameters and expected output
    [CmdletBinding()]
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

        # Find parent modules
        Write-Information -MessageData 'Find parent modules.'
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
        $ParentModulesInstalled = [array](
            $ParentModulesInstalled.Where{
                -not $_.'Name'.Contains('.') -or
                $_.'Name'.Replace(
                    ('.{0}' -f $_.'Name'.Split('.',[System.StringSplitOptions]::RemoveEmptyEntries)[-1]),
                    ''
                ) -notin $ParentModulesInstalled.'Name'
            }
        )

        # Find availble sub modules
        Write-Information -MessageData 'Find available submodules.'
        $AvailableSubModules = [PSCustomObject[]](Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds ($ParentModulesInstalled.'Name'.ForEach{'{0}.*' -f $_}))
        foreach ($ParentModule in $ParentModulesInstalled) {
            # Add available submodules
            $null = Add-Member -InputObject $ParentModule -MemberType 'NoteProperty' -Force -Name 'AvailableSubModules' -Value (
                [PSCustomObject[]](
                    $AvailableSubModules.Where{
                        $_.'Name'.StartsWith('{0}.' -f $ParentModule.'Name') -and
                        $_.'Author' -eq $ParentModule.'Author' -and
                        $(
                            foreach ($ModuleUnwanted in $Script:ModulesUnwanted) {
                                $_.'Name' -eq $ModuleUnwanted -or
                                $_.'Name'.StartsWith($ModuleUnwanted+'.')
                            }
                        ) -notcontains $true
                    } | Select-Object -Property 'Name', 'Version'
                )
            )
            # Check that all available submodules is installed
            $ParentModule.'AvailableSubModules'.ForEach{
                $null = Add-Member -InputObject $_ -MemberType 'NoteProperty' -Force -Name 'Installed' -Value (
                    $_.'Name' -in $ModulesInstalled.'Name'
                )
            }
        }

        # Parent modules with available submodules that is not installed
        $ParentModulesInstalledWithAvailableSubmodules = [PSCustomObject[]](
            $ParentModulesInstalled.Where{$_.'AvailableSubModules'.Where{-not $_.'Installed'}.'Count' -gt 0}
        )
        Write-Information -MessageData (
            'Found {0} parent module(s) with available submodule(s) that is not already installed.' -f $ParentModulesInstalledWithAvailableSubmodules.'Count'.ToString()
        )
        $SubModulesMissing = [PSCustomObject[]](
            $ParentModulesInstalledWithAvailableSubmodules.'AvailableSubModules'.Where{-not $_.'Installed'} | Sort-Object -Property 'Name' -Unique
        )

        # Don't continue if no missing submodules were found
        if ($SubModulesMissing.'Count' -le 0) {
            return
        }

        # Install missing submodules in parallel
        Write-Information -MessageData ('Installing {0} missing submodule(s) in parallel.' -f $SubModulesMissing.'Count'.ToString())
        $Splat = [ordered]@{
            'IncludeXml'          = [bool] $true
            'Name'                = [string[]] $SubModulesMissing.'Name'
            'Path'                = [string] $Script:ModulesPath
            'Repository'          = [string] 'PSGallery'
            'SkipDependencyCheck' = [bool] $true
            'TrustRepository'     = [bool] $true
            'ThrottleLimit'       = [byte] 16
        }
        if (-not [string]::IsNullOrEmpty($Script:TemporaryPath)) {
            $Splat.'TemporaryPath' = [string] $Script:TemporaryPath
        }
        $null = Save-PSResourceInParallel @Splat

        # Check success
        $SubModulesMissing.ForEach{
            $null = Add-Member -InputObject $_ -MemberType 'NoteProperty' -Force -Name 'Success' -Value (
                [System.IO.Directory]::Exists(
                    [System.IO.Path]::Combine($Script:ModulesPath, $_.'Name', $_.'Version'.ToString())
                )
            )
        }
        $Success = [bool]($SubModulesMissing.Where{-not $_.'Success'}.'Count' -le 0)
        Write-Information -MessageData (
            'Success? {0}. {1} of {2} installed successfully.' -f (
                $Success.ToString(),
                $SubModulesMissing.Where{$_.'Success'}.'Count'.ToString(),
                $SubModulesMissing.'Count'.ToString()
            )
        )

        # Output modules not installed
        if (-not $Success) {
            $FailedToInstall = [PSCustomObject[]](
                $SubModulesMissing.Where{-not $_.'Success'}
            )
            Write-Warning -Message ('Following {0} submodule(s) failed to install' -f $FailedToInstall.'Count'.ToString())
            Write-Warning -Message ($FailedToInstall.'Name' | Sort-Object -Unique | Join-String -Separator ', ')
        }

        # If at least one module was successfully installed
        if ($SubModulesMissing.Where{$_.'Success'}.'Count' -gt 0) {
            # Stats
            $Script:ModulesSubInstalledMissing += [string[]]$($SubModulesMissing.Where{$_.'Success'}.'Name')
            # Make sure list of installed modules gets refreshed
            $Script:ModulesInstalledNeedsRefresh = $true
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
                $([array](Get-ChildItem -Path $ModulePath -Directory -Depth 0)).Where{
                    [System.Version]::TryParse($_.'Name',[ref]$null)
                }.ForEach{
                    [PSCustomObject]@{
                        'Path'    = [string] $_.'FullName'
                        'Version' = [System.Version] $_.'Name'
                    }
                }
            )

            # Cut ending 0 if more version numbers than input $Version
            $Versions.ForEach{
                if (
                    $_.'Version'.ToString().Split('.').'Count' -gt $Version.ToString().Split('.').'Count' -and
                    $([uint16]$_.'Version'.ToString().Split('.')[-1]) -eq 0
                ) {
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

    # Input parameters and expected output
    [CmdletBinding()]
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
                    [string]$(if($VersionsAll.'Count' -gt 1){'s'}),
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
        .SYNOPSIS
            Uninstalls installed modules that matches any value in the input parameter $ModulesUnwanted.

        .PARAMETER ModulesUnwanted
            String Array containig modules you don't want to be installed on your system.
    #>
    [CmdletBinding()]
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
        $null = Get-ModulesInstalled

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
                $(if($ModulesToRemove.'Count' -ne 1){'s'}),
                $(if($ModulesToRemove.'Count' -gt 0){
                        ' Will proceed to uninstall {0}.' -f $(if($ModulesToRemove.'Count' -eq 1){'it'}else{'them'})
                    }
                )
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

    # Input parameters and expected output
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
            [System.IO.Directory]::GetFiles(
                [System.IO.Path]::Combine($Script:ScriptsPath,'InstalledScriptInfos'),
                '*.xml'
            ).ForEach{
                Try {
                    [Microsoft.PowerShell.PSResourceGet.UtilClasses.TestHooks]::ReadPSGetResourceInfo(
                        $_
                    )
                }
                Catch {
                    $null
                }
            }.Where{
                $_.'Repository' -eq 'PSGallery' -and $_.'Type' -eq 'Script'
            }
        )

        # Find missing scripts
        $MissingScripts = [string[]](
            $Scripts.Where{$_ -notin $InstalledScripts.'Name'} | Sort-Object
        )
    }

    # Process
    Process {
        # Don't continue if no missing scripts were found
        if ($MissingScripts.'Count' -le 0) {
            Write-Information -MessageData 'Found 0 missing scripts.'
            return
        }

        # Install missing scripts
        foreach ($Script in $MissingScripts) {
            # Introduce current item
            Write-Information -MessageData (
                '{0} / {1} "{2}"' -f (
                    (1 + $MissingScripts.IndexOf($Script)).ToString('0' * $MissingScripts.'Count'.ToString().'Length'),
                    $MissingScripts.'Count'.Tostring(),
                    $Script
                )
            )

            # Find script in PowerShell Gallery
            $PSResource = Microsoft.PowerShell.PSResourceGet\Find-PSResource -Type 'Script' -Repository 'PSGallery' -Name $Script -ErrorAction 'SilentlyContinue'

            # Install it if it was found
            if ($? -and -not [string]::IsNullOrEmpty($PSResource.'Name')) {
                # Special case if Unix
                if (
                    $PSVersionTable.'Platform' -eq 'Unix' -and
                    $PSResource.'Tags'.Where({$_ -in 'Linux','Mac','MacOS','PSEdition_Core'},'First').'Count' -le 0
                ) {
                    Write-Information -MessageData ('{0}This script does not seem to support Unix, thus skipping it.' -f $Indentation*2)
                    Continue
                }
                # Install script
                $Splat = [ordered]@{
                    'AuthenticodeCheck'   = [bool] -not $SkipAuthenticodeCheck
                    'IncludeXml'          = [bool] $true
                    'Name'                = [string] $Script
                    'Path'                = [string] $Script:ScriptsPath
                    'Repository'          = [string] 'PSGallery'
                    'SkipDependencyCheck' = [bool] $true
                    'TrustRepository'     = [bool] $true
                }
                $null = Microsoft.PowerShell.PSResourceGet\Save-PSResource @Splat
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

    # Input parameters and expected output
    [CmdletBinding(SupportsShouldProcess)]
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
            [System.IO.Directory]::GetFiles(
                [System.IO.Path]::Combine($Script:ScriptsPath,'InstalledScriptInfos'),
                '*.xml'
            ).ForEach{
                Try {
                    [Microsoft.PowerShell.PSResourceGet.UtilClasses.TestHooks]::ReadPSGetResourceInfo(
                        $_
                    )
                }
                Catch {
                    $null
                }
            }.Where{
                $_.'Repository' -eq 'PSGallery' -and $_.'Type' -eq 'Script'
            }
        )
    }

    # Process
    Process {
        # Get newest version of installed scripts
        $InstalledScriptsAvailableVersion = [hashtable](Find-PSGalleryPackageLatestVersionUsingApiInBatch -PackageIds $InstalledScripts.'Name' -AsHashtable -MinimalInfo)
        $InstalledScriptsWithNewerVersionAvailable = [PSCustomObject[]](
            $InstalledScripts.Where{
                $_.'Version' -lt $InstalledScriptsAvailableVersion.$($_.'Name').'Version'
            }.ForEach{
                $null = Add-Member -InputObject $_ -MemberType 'NoteProperty' -Force -Name 'VersionAvailable' -Value (
                    $InstalledScriptsAvailableVersion.$($_.'Name').'Version'
                )
            }
        )
        Write-Information -MessageData ('{0} installed script(s) has a newer version available.' -f $InstalledScriptsWithNewerVersionAvailable.'Count'.ToString())
        $InstalledScriptsNotFoundInPsGallery = [PSCustomObject[]]($InstalledScripts.Where{$_.'Name' -notin $InstalledScriptsAvailableVersion.'Keys'})
        if ($InstalledScriptsNotFoundInPsGallery.'Count' -gt 0) {
            Write-Warning -Message 'Following {0} installed script(s) was not found in the PowerShell Gallery:'
            Write-Warning -Message ($Indentation + ($InstalledScriptsNotFoundInPsGallery.'Name' | Sort-Object | Join-String -Separator ', '))
        }

        # Update if any has a newer version
        if ($PSCmdlet.ShouldProcess(('{0} scripts' -f $InstalledScriptsWithNewerVersionAvailable.'Count'.ToString()),'update')) {
            $InstalledScriptsWithNewerVersionAvailable.ForEach{
                # Output current script
                Write-Information -MessageData (
                    '{0} / {1} "{2}" v{3} by "{4}" to v{5}' -f (
                        (1 + $InstalledScriptsWithNewerVersionAvailable.IndexOf($_)).ToString('0' * $InstalledScriptsWithNewerVersionAvailable.'Count'.ToString().'Length'),
                        $InstalledScriptsWithNewerVersionAvailable.'Count'.Tostring(),
                        $_.'Name',
                        $_.'Version'.ToString(),
                        $([string[]]($_.'Author',$_.'Entities'.Where{$_.'Role' -eq 'author'}.'Name')).Where{-not[string]::IsNullOrEmpty($_)}[0],
                        $_.'VersionAvailable'.ToString()
                    )
                )

                # Install script
                $Splat = [ordered]@{
                    'AuthenticodeCheck'   = [bool] -not $SkipAuthenticodeCheck
                    'IncludeXml'          = [bool] $true
                    'Name'                = [string] $_.'Name'
                    'Path'                = [string] $Script:ScriptsPath
                    'Repository'          = [string] 'PSGallery'
                    'SkipDependencyCheck' = [bool] $true
                    'TrustRepository'     = [bool] $true
                    'Version'             = [bool] $_.'VersionAvailable'
                }
                $null = Microsoft.PowerShell.PSResourceGet\Save-PSResource @Splat

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
    [CmdletBinding(SupportsShouldProcess)]
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
        $PSResourcePathWanted = [string] '%LOCALAPPDATA%\Microsoft\PowerShell'
        $RegistryPath = [string] 'Registry::HKEY_CURRENT_USER\Environment'
        $Paths = [PSCustomObject[]](
            [PSCustomObject]@{
                'Type'        = [string] 'Modules'
                'EnvVariable' = [string] 'PSModulePath'
                'Unresolved'  = [string] '{0}\Modules' -f $PSResourcePathWanted
            },
            [PSCustomObject]@{
                'Type'        = [string] 'Scripts'
                'EnvVariable' = [string] 'Path'
                'Unresolved'  = [string] '{0}\Scripts' -f $PSResourcePathWanted
            }
        )
        $Paths.ForEach{
            $null = Add-Member -InputObject $_ -MemberType 'NoteProperty' -Force -Name 'Resolved' -Value (
                [string](
                    [System.Environment]::ExpandEnvironmentVariables($_.'Unresolved')
                )
            )
        }


        # Fix paths
        foreach ($Path in $Paths) {
            # Create directory if it does not exist
            if (-not [System.IO.Directory]::Exists($Path.'Resolved')) {
                $null = [System.IO.Directory]::CreateDirectory($Path.'Resolved')
            }

            # Get current value without resolving the path / expanding the environmental variable
            $CurrentAsString = [string](
                (Get-Item -Path $RegistryPath).GetValue(
                    $Path.'EnvVariable',
                    '',
                    'DoNotExpandEnvironmentNames'
                )
            )

            # Make $Current string array for easier operations
            $CurrentAsArray = [string[]](
                $CurrentAsString.Split(
                    [System.IO.Path]::PathSeparator
                ).Where{
                    -not [string]::IsNullOrEmpty($_)
                }
            )

            # If Type=Modules
            if ($Path.'Type' -eq 'Modules') {
                # Remove "MyDocuments" if present, as it will resolve to OneDrive if Known Folder Move is enabled
                $NewAsArray = [string[]](
                    $CurrentAsArray.Where{
                        $_ -notlike ('{0}\*' -f [System.Environment]::GetFolderPath('MyDocuments'))
                    }
                )
            }
            else {
                $NewAsArray = [string[]]($CurrentAsArray.Where{-not [string]::IsNullOrEmpty($_)}.ForEach{$_})
            }

            # Add $Path.Unresolved if not already present
            if ($NewAsArray -notcontains $Path.'Unresolved' -and $NewAsArray -notcontains $Path.'Resolved') {
                if ($Path.'Type' -eq 'Modules') {
                    $NewAsArray = [string[]](
                        [string[]]($Path.'Unresolved') + [string[]]($NewAsArray) | Where-Object -FilterScript {
                            -not [string]::IsNullOrEmpty($_)
                        }
                    )
                }
                else {
                    $NewAsArray = [string[]](
                        [string[]]($NewAsArray) + [string[]]($Path.'Unresolved') | Where-Object -FilterScript {
                            -not [string]::IsNullOrEmpty($_)
                        }
                    )
                }
            }

            # Convert $PSModulePathNewAsArray to string for easier comparison to existing value
            $NewAsString = [string](($NewAsArray -join [System.IO.Path]::PathSeparator) + [System.IO.Path]::PathSeparator)

            # Set new value if it changed
            if ($NewAsString -ne $CurrentAsString -and $PSCmdlet.ShouldProcess($Path.'EnvVariable','set')) {
                $null = Set-ItemProperty -Path $RegistryPath -Name $Path.'EnvVariable' -Value $NewAsString -Force -Type ([Microsoft.Win32.RegistryValueKind]::ExpandString)
            }
        }
    }

    End {}
}
#endregion User context PSModulePath



#region    Get-Summary
function Get-Summary {
    <#
        .SYNOPSIS
            Outputs statistics after script has ran.
    #>

    # Input parameters and expected output
    [CmdletBinding()]
    [OutputType([System.String])]
    Param ()

    # Begin
    Begin {
    }

    # Process
    Process {
        # Help variables
        $FormatDescriptionLength = [byte]$($Script:StatisticsVariables.ForEach{$_.'Description'.'Length'} | Sort-Object -Descending | Select-Object -First 1)
        $FormatNewLineTab = [string] '{0}{1}' -f [System.Environment]::NewLine, $Indentation

        # Check if any got any objects
        if (
            $Script:StatisticsVariables.ForEach{
                Get-Variable -Name $_.'VariableName' -Scope 'Script' -ValueOnly -ErrorAction 'SilentlyContinue'
            }.Where{
                -not [string]::IsNullOrEmpty($_)
            }.'Count' -le 0
        ) {
            return 'No changes were made.'
        }

        # Output stats
        return (
            $(
                foreach ($Variable in $Script:StatisticsVariables) {
                    $CurrentObject = [string[]]$(Get-Variable -Name $Variable.'VariableName' -Scope 'Script' -ValueOnly -ErrorAction 'SilentlyContinue' | Sort-Object -Unique)
                    if ($CurrentObject.'Count' -le 0) {
                        Continue
                    }
                    $CurrentDescription = [string]$($Variable.'Description' + ':' + [string]$(' ' * [byte]$($FormatDescriptionLength - $Variable.'Description'.'Length')))
                    (
                        '{0} {1}{2}' -f
                        $CurrentDescription,
                        $CurrentObject.'Count',
                        [string]$(if($CurrentObject.'Count' -ge 1){$FormatNewLineTab + [string]$($CurrentObject -join $FormatNewLineTab)})
                    )
                }
            ) -join [System.Environment]::NewLine
        )
    }

    # End
    End {
    }
}
#endregion Get-Summary
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
                [System.IO.Path]::Combine([System.Environment]::GetFolderPath('LocalApplicationData'),'powershell')
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
                    '{0}PowerShell' -f $(if($PSVersionTable.'PSEdition' -eq 'Desktop') {'Windows'})
                )
            }
        )
    )
)
$null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'ModulesPathRoot' -Value (
    [string]($(if($SystemContext){$PSResourceHomeMachine}else{$PSResourceHomeUser}))
)

### Modules path
$null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'ModulesPath' -Value (
    [string] [System.IO.Path]::Combine($Script:ModulesPathRoot,'Modules')
)

### Scripts path
$null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'ScriptsPath' -Value (
    [string] [System.IO.Path]::Combine($Script:ModulesPathRoot,'Scripts')
)

### Temp path if -DevDrive
if ($DevDrive) {
    $null = Set-Variable -Scope 'Script' -Option 'ReadOnly' -Force -Name 'TemporaryPath' -Value (
        [string]($DevDrive + ':\Temp')
    )
}

### Create paths if they don't exist
if (-not $SystemContext) {
    $([string[]]($Script:ModulesPath,$Script:ScriptsPath,$Script:TemporaryPath)).ForEach{
        if (-not [string]::IsNullOrEmpty($_) -and -not [System.IO.Directory]::Exists($_)) {
            $null = [System.IO.Directory]::CreateDirectory($_)
        }
    }
}




# Introduce script run
Write-Information -MessageData ('# Script start at {0}' -f $TimeTotalStart.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Information -MessageData (
    'Scope: "{0}" ("{1}").' -f (
        $(if($SystemContext){'System'}else{'User'}),
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
            Test-Connection -TargetName 'www.powershellgallery.com' -TcpPort 443 -TimeoutSeconds 2 -IPv4 -Quiet -ErrorAction 'SilentlyContinue'
        }
        else {
            Test-NetConnection -ComputerName 'www.powershellgallery.com' -Port 443 -InformationLevel 'Quiet' -ErrorAction 'SilentlyContinue'
        }
    )
) {
    Throw ('ERROR - Could not TCP connect to www.powershellgallery.com:443 within reasonable time. Do you have internet connection, or is the web site down?')
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
                "https://www.powershellgallery.com/api/v2/Packages?`$filter=IsLatestVersion and not IsPrerelease and Id eq 'Microsoft.PowerShell.PSResourceGet'&semVerLevel=1.0.0"
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
            $(if($PSVersionTable.'Platform' -eq 'Unix'){'/tmp'}else{$env:TEMP}),
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
Write-Information -MessageData (Get-Summary)
Write-Information -MessageData ('{0}## Time' -f [System.Environment]::NewLine)
Write-Information -MessageData ('Start time:    {0} ({1}).' -f $Script:TimeTotalStart.ToString('HH\:mm\:ss'), $Script:TimeTotalStart.ToString('o'))
Write-Information -MessageData ('End time:      {0} ({1}).' -f (($Script:TimeTotalEnd = [datetime]::Now).ToString('HH\:mm\:ss'),$Script:TimeTotalEnd.ToString('o')))
Write-Information -MessageData ('Total runtime: {0}.' -f ([string]$([timespan]$($Script:TimeTotalEnd - $Script:TimeTotalStart)).ToString('hh\:mm\:ss')))
#endregion Main
