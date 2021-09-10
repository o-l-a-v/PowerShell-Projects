#Requires -PSEdition Desktop -Version 5.1
<#
    .NAME
        AndroidPlatformToolsUpdater.ps1


    .SYNOPSIS
        Installs and updates Android Platform Tools (ADB, Fastboot ++) and adds install path to Windows Environment Variables.


    .DESCRIPTION
        Installs and updates Android Platform Tools (ADB, Fastboot ++) and adds install path to Windows Environment Variables.

        User Context
            * Installs to "%localappdata%\Android Platform Tools"
            * Will make Android Platform Tools available only to the user logged in when running this script
        
        System Context
            * Installs to "%ProgramFiles(x86)%\Android Platform Tools"
            * Will make Android Platform Tools available to all users on the machine


    .PARAMETER SystemWide
        Optional, Boolean.
        If $true script will install system wide (%ProgramFiles(x86)%), else for current user (%LOCALAPPDATA%)).
        Default value: $true.


    .PARAMETER ForceInstallAndroidPlatformTools
        Optional, Boolean.
        If $true script will install platform-tools, ignoring what ever version might be installed already.


    .EXAMPLE
        # Run from PowerShell ISE, system context
        & $psISE.'CurrentFile'.'FullPath'


    .EXAMPLE
        # Run from PowerShell ISE, user context
        & $psISE.'CurrentFile'.'FullPath' -SystemWide $false


    .NOTES
        Author:   Olav Rønnestad Birkeland
        Created:  190310
        Modified: 210910
#>




# Input parameters
[OutputType($null)]
Param (
    [Parameter(Mandatory = $false, HelpMessage = 'Context - Current User only ($false) or System ($true).')]
    [bool] $SystemWide = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Will force reinstall no matter what installed version that might exist.')]
    [bool] $ForceInstallAndroidPlatformTools = $false
)




# PowerShell preferences
## Output Streams
$DebugPreference       = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$VerbosePreference     = 'SilentlyContinue'
$WarningPreference     = 'Continue'

## Behaviour
$ConfirmPreference     = 'None'
$ProgressPreference    = 'SilentlyContinue'




#region    Functions   
#region    Get-AndroidPlatformToolsInstalledVersion
function Get-AndroidPlatformToolsInstalledVersion {
    <#
        .SYNOPSIS
            Gets Android Platform Tools version already installed in given path on the system.
    #>
            
            
    # Input parameters            
    [CmdletBinding()]
    [OutputType([System.Version])]
    Param(
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string] $PathDirAndroidPlatformTools = '{0}\Android Platform Tools' -f $(
            if ($SystemWide) {
                ${env:ProgramFiles(x86)}
            }
            else {
                $env:LOCALAPPDATA
            }
        )
    )
            
    # Begin
    Begin {}


    # Process
    Process {
        # Assets
        $PathFileFastboot = [string]('{0}\fastboot.exe' -f ($PathDirAndroidPlatformTools))

        # Version of existing install, Version 0.0.0.0 if not found
        $VersionFileFastbootExisting = [System.Version]$(
            if ([System.IO.File]::Exists($PathFileFastboot)) {
                Try{[System.Version]$([string](cmd /c ('"{0}" --version' -f ($PathFileFastboot))).Split(' ')[2].Replace('-','.'))}Catch{'0.0.0.0'}
            }
            else {
                '0.0.0.0'
            }
        )
    }
            

    # End
    End {
        return $VersionFileFastbootExisting
    }
}
#endregion Get-AndroidPlatformToolsInstalledVersion



#region    Install-AndroidPlatformToolsLatest
function Install-AndroidPlatformToolsLatest {
    <#
        .SYNOPSIS
            Installs Android Platform Tools Latest Version to given Path.
                       
        .PARAMETER PathDirAndroidPlatformTools
            Path to where Android Platform Tools will be installed.
            Optional String [string].
            Default value: "%ProgramFiles(x86)%\Android Platform Tools"

        .PARAMETER VersionFileFastbootInstalled
            Currently installed version of Android Platform Tools.
            Optional Variable, Version [System.Version].
            Default value: Function "Get-AndroidPlatformToolsInstalledVersion".

        .PARAMETER CleanUpWhenDone
            Whether to clean up downloaded and extracted files when done.
            Optional variable, boolean.
            Default value: $True.
    #>
            
            
    # Input parameters
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    Param (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string] $PathDirAndroidPlatformTools = '{0}\Android Platform Tools' -f $(
            if ($SystemWide) {
                ${env:ProgramFiles(x86)}
            }
            else {
                $env:LOCALAPPDATA
            }
        ),

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Version] $VersionFileFastbootInstalled = [System.Version](Get-AndroidPlatformToolsInstalledVersion),

        [Parameter(Mandatory = $false)]
        [bool] $CleanUpWhenDone = $true
    )


    # Begin
    Begin {
        # Assets - Function help variables
        $CurrentUserIsAdmin = [bool](
            ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
                [Security.Principal.WindowsBuiltInRole]::Administrator
            )
        )
        $Success = [bool] $false
                
        # Assets - Downlaod link
        $UrlFileAndroidPlatformTools          = [string] 'https://dl.google.com/android/repository/platform-tools-latest-windows.zip'
                
        # Assets - Temp directory
        $PathDirTemp                          = [string] $env:TEMP
        $PathDirTempAndroidPlatformTools      = [string] '{0}\platform-tools' -f $PathDirTemp
        $PathFileTempFastboot                 = [string] '{0}\fastboot.exe' -f $PathDirTempAndroidPlatformTools
        $PathFileDownloadAndroidPlatformTools = [string] '{0}\{1}' -f $PathDirTemp,$UrlFileAndroidPlatformTools.Split('/')[-1]
    }
    
            
    # Process
    Process {
        # Make sure current user is admin
        if (-not($CurrentUserIsAdmin)) {
            Throw ('ERROR: This function must be run as administrator.')
        }
        
        
        # Remove existing files
        $([string[]]($PathDirTempAndroidPlatformTools,$PathFileDownloadAndroidPlatformTools)).ForEach{
            if (Test-Path -Path $_) {
                $null = Remove-Item -Path $_ -Recurse -Force -ErrorAction 'Stop'
            }
        }


        # Download                
        Write-Information -MessageData ('Downloading Android Platform Tools from "{0}".' -f $UrlFileAndroidPlatformTools)
        $Success = [bool]$(Try{[System.Net.WebClient]::new().DownloadFile($UrlFileAndroidPlatformTools,$PathFileDownloadAndroidPlatformTools);$?}Catch{$false})
        if (-not($Success -and [System.IO.File]::Exists($PathFileDownloadAndroidPlatformTools))) {
            Throw ('ERROR: Failed to download "{0}".' -f ($UrlFileAndroidPlatformTools))
        }
    
      
        # Extract
        ## Write information
        Write-Information -MessageData ('Extracting "{0}" to "{1}".' -f $PathFileDownloadAndroidPlatformTools, $PathDirTemp)
                
        ## See if 7-Zip is present
        $7Zip = [string]$(
            $([string[]]($env:ProgramW6432,${env:CommonProgramFiles(x86)})).Where{
                [System.IO.Directory]::Exists($_)
            }.ForEach{
                '{0}\7-Zip\7z.exe' -f $_
            }.Where{
                [System.IO.File]::Exists($_)
            } | Select-Object -First 1
        )

        ## Use 7-Zip if present, fall back to built in method                
        if (-not[string]::IsNullOrEmpty($7Zip)) {
            $7ZipArgs = [string] 'x "{0}" -o"{1}" -y' -f $PathFileDownloadAndroidPlatformTools, $PathDirTemp
            Write-Information -MessageData ('Using 7-Zip "{0}" "{1}".' -f $7Zip,$7ZipArgs)
            $null = Start-Process -WindowStyle 'Hidden' -Wait -FilePath $7Zip -ArgumentList $7ZipArgs                    
            $Success = [bool]($? -and [System.IO.Directory]::Exists($PathDirTempAndroidPlatformTools))
        }
        if ([string]::IsNullOrEmpty($7Zip) -or -not $Success) {
            Write-Information -MessageData 'Using .NET [System.IO.Compression.FileSystem]::ExtractToDirectory()'
            Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
            $Success = [bool]$(
                Try {
                    [System.IO.Compression.ZipFile]::ExtractToDirectory($PathFileDownloadAndroidPlatformTools,$PathDirTemp)
                    $?
                }
                Catch {
                    $false
                }
            )
        }

        ## Check success
        if (-not($Success -and [System.IO.File]::Exists($PathFileTempFastboot))) {
            Throw ('ERROR: Failed to extract "{0}".' -f ($PathFileDownloadAndroidPlatformTools))
        }
    

        # Version of download Android Platform Tools
        $VersionFileFastbootDownloaded = [System.Version]$(
            if (Test-Path -Path $PathFileTempFastboot) {
                Try{[System.Version]$([string](cmd /c ('"{0}" --version' -f ($PathFileTempFastboot))).Split(' ')[2].Replace('-','.'))}Catch{'0.0.0.0'}
            }
            else {
                '0.0.0.0'
            }
        )
        if ($VersionFileFastbootDownloaded -eq [System.Version]('0.0.0.0')) {
            Throw ('ERROR: Failed to get version info from "{0}".' -f ($PathFileTempFastboot))
        }


        # Install Downloaded version if newer that Installed Version
        if ($ForceInstallAndroidPlatformTools -or $VersionFileFastbootDownloaded -gt $VersionFileFastbootInstalled) {            
            # Write information
            if ($VersionFileFastbootInstalled -eq [System.Version]('0.0.0.0')) {
                Write-Information -MessageData 'Android Platform Tools are not already installed.'
            }
            else {                    
                if ($ForceInstallAndroidPlatformTools) {
                    Write-Information -MessageData '$ForceInstallAndroidPlatformTools is $true.'
                }
                else {
                    Write-Information -MessageData (
                        'Installed version (v{0}) is not already up to date (v{1}).' -f (
                            $VersionFileFastbootInstalled.ToString(),
                            $VersionFileFastbootDownloaded.ToString()
                        )
                    )
                }
            }

            # Kill ADB and Fastboot if running
            $([string[]]('adb','fastboot')).ForEach{
                $null = Get-Process -Name $_ -ErrorAction 'SilentlyContinue' | Stop-Process -ErrorAction 'Stop'
            }
            
            # Remove Existing Files if they Exist
            if ([System.IO.Directory]::Exists($PathDirAndroidPlatformTools)) {
                $null = Remove-Item -Path $PathDirAndroidPlatformTools -Recurse -Force -ErrorAction 'Stop'
                if ((-not($?)) -or [System.IO.Directory]::Exists($PathDirAndroidPlatformTools)) {
                    Throw ('ERROR: Failed to remove existing files in "{0}".' -f $PathAndroidPlatformTools)
                }
            }

            # Install Downloaded Files
            Move-Item -Path $PathDirTempAndroidPlatformTools -Destination $PathDirAndroidPlatformTools -Force -Include '*'
            
            # Capture Operation Success
            $Success = [bool] $?
        }
        else {
            # Write information
            Write-Information -MessageData ('Installed version (v{0}) is already up to date.' -f $VersionFileFastbootInstalled.ToString())
        }
    }


    # End
    End {
        # Clean up
        if ($CleanUpWhenDone) {
            $([string[]]($PathDirTempAndroidPlatformTools,$PathFileDownloadAndroidPlatformTools)).ForEach{
                if (Test-Path -Path $_) {
                    $null = Remove-Item -Path $_ -Recurse -Force -ErrorAction 'Stop'
                }
            }
        }


        # Return success status
        return $Success
    }
}
#endregion Install-AndroidPlatformToolsLatest



#region    Add-AndroidPlatformToolsToEnvironmentVariables
function Add-AndroidPlatformToolsToEnvironmentVariables {
    <#
        .SYNOPSIS
            Adds path to Android Platform Tools to Windows Environment Variables for Current User ONLY.

        .PARAMETER PathDirAndroidPlatformTools
            Path to where Android Platform Tools will be installed.
            Optional String [string].
            Default Value: "%ProgramFiles(x86)%\Android Platform Tools".

        .PARAMETER SystemWide
            Add to Environment Variables for Current User (HKCU) only, or System Wide (HKLM).
            Optional Boolean [bool].
            Default Value: $false.
    #>
            
            
    # Input parameters
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    Param(
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({[System.IO.Directory]::Exists($_)})]
        [string] $PathDirAndroidPlatformTools = '{0}\Android Platform Tools' -f $(
            if ($SystemWide) {
                ${env:ProgramFiles(x86)}
            }
            else {
                $env:LOCALAPPDATA
            }
        ),

        [Parameter(Mandatory = $false)]
        [ValidateSet($true,$false)]
        [bool] $SystemWide = $false
    )
    

    # Begin
    Begin {
        $Success = [bool] $true
        $Target  = [string] $(if($SystemWide){'Machine'}else{'User'})
    }

            
    # Process
    Process {
        # Get existing PATH Environment Variable
        $PathVariableExisting = [string[]]([System.Environment]::GetEnvironmentVariables($Target).'Path'.Split(';') | Sort-Object)
        $PathVariableNew      = [Management.Automation.PSSerializer]::DeSerialize([Management.Automation.PSSerializer]::Serialize($PathVariableExisting))

        # Add Android Platform Tools if not already present
        if ($PathVariableNew -notcontains $PathDirAndroidPlatformTools) {
            $PathVariableNew += $PathDirAndroidPlatformTools
        }

        # Clean up
        ## Remove ending '\' and return unique entries only
        $PathVariableNew = [string[]](
            $PathVariableNew.ForEach{
                if ($_[-1] -eq '\') {
                    $_.SubString(0,$_.'Length'-1)
                }
                else {
                    $_
                }
            } | Sort-Object -Unique
        )

        # Change PATH Environment Variable for Current User
        if (([string[]]$(Compare-Object -ReferenceObject $PathVariableNew -DifferenceObject $PathVariableExisting -PassThru)).'Count' -ge 1) {
            # Set new environmental variables
            [System.Environment]::SetEnvironmentVariable('Path',[string]$($PathVariableNew -join ';'),$Target)
            $Success = [bool] $?
        }
        else {
            $Success = [bool] $true
        }
    }


    # End
    End {
        Return $Success
    }
}
#endregion Add-AndroidPlatformToolsToEnvironmentVariables


#region    Get-AdbVersionFromWebpage
function Get-AndroidPlatformToolsFromWebpage {
    [OutputType([System.Version])]
    Param()
    Try {
        [System.Net.WebClient]::new().DownloadString('https://developer.android.com/studio/releases/platform-tools').Split(
            [System.Environment]::NewLine,
            [System.StringSplitOptions]::RemoveEmptyEntries
        ).ForEach{
            $_.Trim() -replace '\s{2,}', ' '
        }.Where{
            $_ -like '<h4 id=*'
        }[0].Split('>')[1].Split(' ')[0]
    }
    Catch {
        '0.0.0.0'
    }
}
#endregion Get-AdbVersionFromWebpage
#endregion Functions





#region    Main
    # Help variables
    ## Current user and context
    $CurrentUserName = [string] [System.Security.Principal.WindowsIdentity]::GetCurrent().'Name'
    $CurrentUserSID  = [string] [System.Security.Principal.WindowsIdentity]::GetCurrent().'User'.'Value'
    $IsAdmin  = [bool](([System.Security.Principal.WindowsPrincipal]([System.Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
    $IsSystem = [bool]($CurrentUserSID -eq 'S-1-5-18')
    
    ## Output path for Android Platform Tools
    $PathDirAndroidPlatformTools = [string]('{0}\Android Platform Tools' -f $(if($SystemWide){${env:ProgramFiles(x86)}}else{$env:LOCALAPPDATA}))
    

    # Write info     
    Write-Information -MessageData '# AndroodPlatformToolsUpdater.ps1'    
    Write-Information -MessageData '## Info'
    Write-Information -MessageData ('Current user name: "{0}".' -f $CurrentUserName)
    Write-Information -MessageData ('Current user SID:  "{0}".' -f $CurrentUserSID)
    Write-Information -MessageData ('Is administrator?  "{0}".' -f $IsAdmin.ToString())
    Write-Information -MessageData ('Is SYSTEM?         "{0}".' -f $IsSystem.ToString())
    Write-Information -MessageData ('Running in {0} context with{1} admin permissions.' -f $(if($SystemWide){'system'}else{'user'}),$(if(-not$IsAdmin){'out'}))
    Write-Information -MessageData ('Output path for ADB tools: "{0}".' -f $PathDirAndroidPlatformTools)
    

    # Failproof
    ## Check if running as Administrator if $SystemWide
    if ($SystemWide -and -not $IsAdmin) {
        Throw 'ERROR: Must run as administrator when $SystemWide is set to True.'
        Exit 1
    }
    
    ## Check if running as System if not $SystemWide
    if (-not $SystemWide -and $IsSystem) {
        Throw 'ERROR: Must not run as NT AUTHORITY\SYSTEM if $SystemWide is set to $false.'
        Exit 1
    }
        

    # Get version info
    ## Introduce step
    Write-Information -MessageData ('{0}# Get version info' -f [System.Environment]::NewLine)

    ## Installed version
    Write-Information -MessageData ('## Installed version')
    $VersionInstalled = [System.Version](Get-AndroidPlatformToolsInstalledVersion -PathDirAndroidPlatformTools $PathDirAndroidPlatformTools)    
    if ($VersionInstalled -ne [System.Version]('0.0.0.0')) {
        Write-Information -MessageData ('Installed version: "{0}".' -f $VersionInstalled.ToString())
    }
    else {
        Write-Information -MessageData 'Found no installed version.'
    }
    
    ## Available version
    Write-Information -MessageData ('{0}## Available version' -f [System.Environment]::NewLine)
    $VersionAvailable = [System.Version](Get-AndroidPlatformToolsFromWebpage)
    if ($VersionAvailable -ne [System.Version]('0.0.0.0')) {
        Write-Information -MessageData ('Available version: "{0}".' -f $VersionAvailable.ToString())
    }
    else {
        Write-Warning -Message 'Failed to parse version info from Google web page, will install newest version.'
    }


    # Install platform-tools
    Write-Information -MessageData ('{0}{0}# Install Android platform tools' -f [System.Environment]::NewLine)
    if (
        $ForceInstallAndroidPlatformTools -or
        $VersionInstalled -eq [System.Version]('0.0.0.0') -or
        $VersionAvailable -eq [System.Version]('0.0.0.0') -or 
        $VersionAvailable -gt $VersionInstalled        
    ) {
        Write-Information -MessageData 'Installing newest version.'
        $Success = [bool](
            Install-AndroidPlatformToolsLatest -PathDirAndroidPlatformTools $PathDirAndroidPlatformTools -VersionFileFastbootInstalled $VersionInstalled
        )
        Write-Information -MessageData ('{0}Success? {1}.' -f "`t", $Success.ToString())
    }
    else {
        Write-Information -MessageData 'Already on latest version.'
    }


    # Update environment variables for given context
    Write-Information -MessageData ('{0}{0}# Update environmental variables' -f [System.Environment]::NewLine)
    $Success = [bool](Add-AndroidPlatformToolsToEnvironmentVariables -PathDirAndroidPlatformTools $PathDirAndroidPlatformTools -SystemWide $SystemWide)
    Write-Information -MessageData (
        'Checking and eventually adding Android Platform Tools to {0} Environment Variables. Success? {1}.' -f (
            [string]$(if($SystemWide){'System Wide'}else{'Current User'}),
            $Success.ToString()
        )
    )
#endregion Main
