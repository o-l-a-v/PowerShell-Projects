#Requires  -Version 5.1 -RunAsAdministrator
<#
    .NAME
        AndroidPlatformToolsUpdater.ps1

    .SYNAPSIS
        Installs and updates Android Platform Tools (ADB, Fastboot ++) and adds install path to Windows Environment Variables.

    .NOTES
        Author:         Olav Rønnestad Birkeland
        Version:        1.0.0.0
        Creation Date:  190310
        Last Edit Date: 190310
#>




#region    Settings
    # Script Option - Path to where Android Platform Tools will be installed
    $PathDirAndroidPlatformTools = [string]$('{0}\Android Platform Tools' -f (${env:ProgramFiles(x86)}))

    # Script Option - Add Android Platform Tools to System Environment Variables System Wide (HKLM, $true) or Current User (HKCU, $false)
    $AddToEnvironmentVariablesSystemWide = [bool]$($false)

    # Script Option - Force Install: Used if you've accidently removed some of the files inside Android Platform Tools folder or similar.
    $ForceInstallAndroidPlatformTools = [bool]$($false)

    # Settings - PowerShell Output Streams
    $DebugPreference       = 'SilentlyContinue'
    $ErrorActionPreference = 'Stop'
    $VerbosePreference     = 'SilentlyContinue'

    # Settings - PowerShell Behaviour
    $ProgressPreference    = 'SilentlyContinue'
#endregion Settings




#region    Functions
    #region    Get-AndroidPlatformToolsLatestVersion
        function Get-AndroidPlatformToolsLatestVersion {
            <#
                .SYNAPSIS
                    Gets Android Platform Tools latest version number by reading from the website.
            #>
            [CmdletBinding()]
            Param()


            # Begin Function
            Begin {
                $Url = [string]$('https://developer.android.com/studio/releases/platform-tools')
                Write-Verbose -Message ('Will try to get latest version number from "{0}".' -f ($Url))
            }


            # Process
            Process {
                Try {
                    $WebPage = Invoke-WebRequest -Uri $Url -ErrorAction 'Stop' | Select-Object -ExpandProperty 'Content'
                    $Version = [System.Version]$(($WebPage.Split("`r`n") | Where-Object -FilterScript {$_ -like '<h4*'} | Select-Object -First 1).Split('>')[1].Split(' ')[0])
                }
                Catch {
                    Write-Verbose -Message ('ERROR: Logic for getting out latest version of Android Platform Tools did not work.')
                    $Version = [System.Version]('0.0')
                }
            }
            

            # End
            End {
                return $Version
            }
        }
    #endregion Get-AndroidPlatformToolsLatestVersion



    #region    Get-AndroidPlatformToolsInstalledVersion
        function Get-AndroidPlatformToolsInstalledVersion {
            <#
                .SYNAPSIS
                    Gets Android Platform Tools version already installed in given path on the system.
            #>
            [CmdletBinding()]
            Param(
                [Parameter(Mandatory = $false)]
                [ValidateNotNullOrEmpty()]
                [ValidateScript({[bool]$(Test-Path -Path $_ -ErrorAction 'SilentlyContinue')})]
                [string] $PathDirAndroidPlatformTools = [string]$('{0}\Android Platform Tools' -f (${env:ProgramFiles(x86)}))
            )
            
            # Begin
            Begin {}


            # Process
            Process {
                # Assets
                $PathFileFastboot = [string]('{0}\fastboot.exe' -f ($PathDirAndroidPlatformTools))

                # Version of existing install, Version 0.0.0.0 if not found
                $VersionFileFastbootExisting = [System.Version]$(
                    if (Test-Path -Path $PathFileFastboot) {
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



    #region     Install-AndroidPlatformToolsLatest
        function Install-AndroidPlatformToolsLatest {
            <#
                .SYNAPSIS
                    Installs Android Platform Tools Latest Version to given Path.
                       
                .PARAMETER PathDirAndroidPlatformTools
                    Path to where Android Platform Tools will be installed.
                    Optional String [string].
                    Default Value: "%ProgramFiles(x86)%\Android Platform Tools"

                .PARAMETER VersionFileFastbootInstalled
                    Currently installed version of Android Platform Tools.
                    Optional Variable, Version [System.Version].
                    Default Value: Function "Get-AndroidPlatformToolsInstalledVersion".
            #>
            [CmdletBinding()]
            Param (
                [Parameter(Mandatory = $false)]
                [ValidateNotNullOrEmpty()]
                [string] $PathDirAndroidPlatformTools = [string]$('{0}\Android Platform Tools' -f (${env:ProgramFiles(x86)})),

                [Parameter(Mandatory = $false)]
                [ValidateNotNullOrEmpty()]
                [System.Version] $VersionFileFastbootInstalled = [System.Version]$(Get-AndroidPlatformToolsInstalledVersion)
            )


            # Begin
            Begin {
                # Assets - Function Help Variables
                $CurrentUserIsAdmin = [bool]$([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                $Success = [bool]$($false)
                # Assets - Downlaod Link
                $UrlFileAndroidPlatformTools          = [string]$('https://dl.google.com/android/repository/platform-tools-latest-windows.zip')
                # Assets - Temp Directory
                $PathDirTemp                          = [string]$($env:TEMP)
                $PathDirTempAndroidPlatformTools      = [string]$('{0}\platform-tools' -f ($PathDirTemp))
                $PathFileTempFastboot                 = [string]$('{0}\fastboot.exe' -f ($PathDirTempAndroidPlatformTools))
                $PathFileDownloadAndroidPlatformTools = [string]$('{0}\{1}' -f ($PathDirTemp,$UrlFileAndroidPlatformTools.Split('/')[-1]))
            }
    
            
            # Process
            Process {
                # Make sure current user is admin
                if (-not($CurrentUserIsAdmin)) {
                    Throw ('ERROR: This function must be run as administrator.')
                }
        
        
                # Remove existing files
                [string[]]@($PathDirTempAndroidPlatformTools,$PathFileDownloadAndroidPlatformTools) | ForEach-Object {if (Test-Path -Path $_){$null = Remove-Item -Path $_ -Recurse -Force -ErrorAction 'Stop'}}


                # Download
                $Success = [bool]$(Try{[System.Net.WebClient]::new().DownloadFile($UrlFileAndroidPlatformTools,$PathFileDownloadAndroidPlatformTools);$?}Catch{$false})
                if (-not($Success -and [bool]$(Test-Path -Path $PathFileDownloadAndroidPlatformTools))) {
                    Throw ('ERROR: Failed to download "{0}".' -f ($UrlFileAndroidPlatformTools))
                }
    
      
                # Extract
                Add-Type -AssemblyName 'system.io.compression.filesystem'
                $Success = [bool]$(Try{[io.compression.zipfile]::ExtractToDirectory($PathFileDownloadAndroidPlatformTools,$PathDirTemp);$?}Catch{$false})
                if (-not($Success -and [bool]$(Test-Path -Path $PathFileTempFastboot))) {
                    Throw ('ERROR: Failed to extract "{0}".' -f ($PathFileDownloadAndroidPlatformTools))
                }
    

                # Version of download Android Platform Tools
                $VersionFileFastbootDownlaoded = [System.Version]$(
                    if (Test-Path -Path $PathFileTempFastboot) {
                        Try{[System.Version]$([string](cmd /c ('"{0}" --version' -f ($PathFileTempFastboot))).Split(' ')[2].Replace('-','.'))}Catch{'0.0.0.0'}
                    }
                    else {
                        '0.0.0.0'
                    }
                )
                if ($VersionFileFastbootDownlaoded -eq [System.Version]('0.0.0.0')) {
                    Throw ('ERROR: Failed to get version info from "{0}".' -f ($PathFileTempFastboot))
                }


                # Install Downloaded version if newer that Installed Version
                if ($VersionFileFastbootDownlaoded -gt $VersionFileFastbootInstalled) {
            
                    # Kill ADB if running
                    Get-Process -Name 'adb' -ErrorAction 'SilentlyContinue' | Stop-Process -ErrorAction 'Stop'
            
                    # Kill Fastboot if running
                    Get-Process -Name 'fastboot' -ErrorAction 'SilentlyContinue' | Stop-Process -ErrorAction 'Stop'
            
                    # Remove Existing Files if they Exist
                    if (Test-Path -Path $PathDirAndroidPlatformTools) {         
                        $null = Remove-Item -Path $PathDirAndroidPlatformTools -Recurse -Force -ErrorAction 'Stop'
                        if ((-not($?)) -or [bool]$(Test-Path -Path $PathDirAndroidPlatformTools)) {
                            Throw ('ERROR: Failed to remove existing files in "{0}".' -f ($PathAndroidPlatformTools))
                        }
                    }

                    # Install Downloaded Files
                    Move-Item -Path $PathDirTempAndroidPlatformTools -Destination $PathDirAndroidPlatformTools -Force -Include '*'
            
                    # Capture Operation Success
                    $Success = [bool]$($?)
                }
            }


            # End
            End {
                return $Success
            }
        }
    #endregion Install-AndroidPlatformToolsLatest



    #region    Add-AndroidPlatformToolsToEnvironmentVariables
        Function Add-AndroidPlatformToolsToEnvironmentVariables {
            <#
                .SYNAPSIS
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
            [CmdletBinding()]
            Param(
                [Parameter(Mandatory = $false)]
                [ValidateNotNullOrEmpty()]
                [ValidateScript({[bool]$(Test-Path -Path $_ -ErrorAction 'SilentlyContinue')})]
                [string] $PathDirAndroidPlatformTools = [string]$('{0}\Android Platform Tools' -f (${env:ProgramFiles(x86)})),

                [Parameter(Mandatory = $false)]
                [ValidateSet($true,$false)]
                [bool]   $SystemWide = [bool]$($false)
            )
    

            # Begin
            Begin {
                $Success = [bool]$($true)
                $Path    = [string]$(if($SystemWide){'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'}else{'Registry::HKEY_CURRENT_USER\Environment'})
            }

            
            # Process
            Process {
                # Get existing PATH Environment Variable
                $PathVariableExisting = [string[]]@((Get-ItemProperty -Path $Path -Name 'Path' | Select-Object -ExpandProperty 'Path').Split(';') | Sort-Object)
                $PathVariableNew      = [Management.Automation.PSSerializer]::DeSerialize([Management.Automation.PSSerializer]::Serialize($PathVariableExisting))

                # Add Android Platform Tools if not already present
                if (($PathVariableExisting | Where-Object -FilterScript {$_ -eq $PathDirAndroidPlatformTools}).Count -lt 1) {
                    $PathVariableNew = [string[]]@($PathVariableNew + $PathDirAndroidPlatformTools | Select-Object -Unique | Sort-Object)   
                }

                # Change PATH Environment Variable for Current User
                if (([string[]]$(Compare-Object -ReferenceObject $PathVariableNew -DifferenceObject $PathVariableExisting -PassThru)).Count -ge 1) {
                    $null = Set-ItemProperty -Path $Path -Name 'Path' -Value ($PathVariableNew -join ';')
                    $Success = [bool]$($?)
                }
                else {
                    $Success = [bool]$($true)
                }
            }


            # End
            End {
                Return $Success
            }
        }
    #endregion Add-AndroidPlatformToolsToEnvironmentVariables
#endregion Functions





#region    Main
        # Get Version Info
        $VersionInstalled = [System.Version]$(Get-AndroidPlatformToolsInstalledVersion)
        $VersionLatest    = [System.Version]$(Get-AndroidPlatformToolsLatestVersion)

        # Throw Error if both turned out to be 0.0.0.0
        if ($VersionInstalled -eq [System.Version]$('0.0.0.0') -and $VersionLatest -eq [System.Version]$('0.0.0.0')) {
            Throw ('ERROR: Both Installed Version and Latest Version of Android Platform Tools Turned Out To Be v0.0.0.0. Not possible.')
        }

        # Install New Version If Installed Version Is Outdated
        if ($VersionLatest -gt $VersionInstalled -or $ForceInstallAndroidPlatformTools) {
            Write-Output -InputObject ('Installed version is outdated, there is not installed version at all, or Force is specified. Installing newest version (v{0})...' -f ($VersionLatest))
            $Success = [bool]$(Install-AndroidPlatformToolsLatest -PathDirAndroidPlatformTools $PathDirAndroidPlatformTools -VersionFileFastbootInstalled $VersionInstalled)
            Write-Output -InputObject ('{0}Success? {1}.' -f ("`t",$Success.ToString()))
        }
        else {
            Write-Output -InputObject ('Installed version is up to date (v{0}).' -f ($VersionInstalled))
        }

        # Update Environment Variables for Current User
        $Success = [bool]$(Add-AndroidPlatformToolsToEnvironmentVariables -PathDirAndroidPlatformTools $PathDirAndroidPlatformTools -SystemWide $AddToEnvironmentVariablesSystemWide)
        Write-Output -InputObject ('Checking and eventually adding Android Platform Tools to {0} Environment Variables. Success? {1}.' -f ([string]$(if($SystemWide){'System Wide'}else{'Current User'}),$Success.ToString()))
#endregion Main