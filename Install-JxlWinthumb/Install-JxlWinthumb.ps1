<#
    .SYNOPSIS
        Install "jxl_winthumb" from GitHub.

    .NOTES
        Author:   Olav Rønnestad Birkeland | github.com/o-l-a-v
        Created:  2024-03-28
        Modified: 2026-02-21

    .EXAMPLE
        # Check only
        & $(Try{$psEditor.GetEditorContext().CurrentFile.Path}Catch{$psISE.CurrentFile.FullPath}) -WhatIf; $LASTEXITCODE

    .EXAMPLE
        # Write changes
        & $(Try{$psEditor.GetEditorContext().CurrentFile.Path}Catch{$psISE.CurrentFile.FullPath}); $LASTEXITCODE

    .EXAMPLE
        # Force reinstall
        & $(Try{$psEditor.GetEditorContext().CurrentFile.Path}Catch{$psISE.CurrentFile.FullPath}) -ForceReinstall; $LASTEXITCODE
#>



# Input and expected output
[CmdletBinding(SupportsShouldProcess)]
[OutputType([System.Void])]
Param(
    [Parameter(HelpMessage = 'Whether to force reinstall, even if the same DLL already is installed.')]
    [switch] $ForceReinstall
)



# PowerShell preferences
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'



# Failproof
## Running on Windows
if ($PSVersionTable.'Platform' -eq 'Unix') {
    Throw 'Script is made for Windows only.'
}

## Running as 64 bit process
if (-not [System.Environment]::Is64BitProcess) {
    Throw 'Script must be run as 64 bit process.'
}

## Running as administrator if not -WhatIf
if (
    -not $PSBoundParameters.ContainsKey('WhatIf') -and
    -not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
) {
    Write-Information -MessageData 'Running without -WhatIf requires script to run as administrator.'
    # Check if gsudo is present
    $GsudoIsPresent = [bool](Get-Command -Name 'gsudo' -ErrorAction 'Ignore')
    # Ask to escalate if gsudo was found
    if ($GsudoIsPresent) {
        # Ask user whether to escalate
        $Choice = [string]::Empty
        do {
            $Choice = Read-Host -Prompt 'Escalate to run as admin using gsudo? Y / N'
            if ($Choice -notin 'n', 'y') {
                Write-Warning -Message ('You entered "{0}", but choice must be either "n" or "y". CTRL+C to abort.' -f $Choice)
            }
        }
        until ($Choice -in 'n', 'y')
        # Escalate if user chose yes
        if ($Choice -eq 'y') {
            $InputParameters = [string[]](
                $PSBoundParameters.GetEnumerator().ForEach{
                    if ($_.'Value' -is [switch]) {
                        if ($_.'Value'.'IsPresent') { "-$($_.Key)" }
                    }
                    else {
                        "-$($_.Key)", "`"$($_.Value)`""
                    }
                }
            )
            gsudo (Get-Process -Id $PID).'Path' -File $MyInvocation.'MyCommand'.'Path' $InputParameters
            Exit $LASTEXITCODE
        }
        else {
            Write-Information -MessageData 'You chose not to escalate, script exited.'
            Exit 0
        }
    }
    # Else exit the script
    else {
        Write-Information -MessageData 'Script will exit as it did not find gsudo nor sudo. Run script as admin manually instead.'

    }
}



# Assets
$GitHubRepo = [string] 'https://github.com/saschanaz/jxl-winthumb'
$FileName = [string](
    'jxl_winthumb_{0}.dll' -f
    $(
        switch -Exact ([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture) {
            'Arm64' {'aarch64'}
            'X64' {'x86_64'}
            'X86' {'i686'}
            default {Throw ('"{0}" is an unhandled OS architecture.' -f $_)}
        }
    )
)
$DownloadPath = [string] [System.IO.Path]::Combine($env:TEMP, $FileName)
$DestinationPath = [string] [System.IO.Path]::Combine($env:windir, 'System32', $FileName)



# Script help variable
$Success = [bool] $true
$ExitCode = [int32] 0



# Wrap in Try/Catch
Try {
    # Get latest version info from GitHub
    ## Introduce step
    Write-Information -MessageData '# Get latest version info from GitHub'
    $Latest = [PSCustomObject](
        $(
            [PSCustomObject[]](
                Invoke-RestMethod -Method 'Get' -Uri ('https://api.github.com/repos/{0}/releases' -f $([uri]$GitHubRepo).'AbsolutePath'.TrimStart('/'))
            )
        ) | Sort-Object -Property @{'Expression' = {[datetime]$_.'published_at'}} -Descending | Select-Object -First 1
    )

    ## Get latest version from GitHub
    $LatestVersion = [System.Version]($Latest.'tag_name'.TrimStart('v'))
    $LatestFile = [PSCustomObject] $Latest.'assets'.Where({$_.'name' -eq $FileName}, 'First')[0]
    $LatestFileUri = [string] $LatestFile.'browser_download_url'
    $LatestFileSize = [uint32] $LatestFile.'size'
    $LatestSha256 = [string] $LatestFile.'digest'.Split(':')[-1]

    ## Output info
    Write-Information -MessageData ('LatestVersion: {0}' -f $LatestVersion.ToString())
    Write-Information -MessageData ('LatestFileUri: {0}' -f $LatestFileUri)
    Write-Information -MessageData ('LatestSha256:  {0}' -f $LatestSha256)

    ## Failproof
    if ([string]::IsNullOrWhitespace($LatestFileUri)) {
        Throw 'Failed to get URL of latest version DLL.'
    }
    if ([string]::IsNullOrWhiteSpace($LatestSha256)) {
        Throw 'Failed to get SHA256 of latest version DLL.'
    }


    # Detect and compare existing file
    ## Introduce step
    Write-Information -MessageData ('{0}# Detect and compare aldready installed DLL by checksum and version' -f [System.Environment]::NewLine)

    ## See if existing file already exist
    if ([System.IO.File]::Exists($DestinationPath)) {
        Write-Information -MessageData ('Destination "{0}" already exists.' -f $DestinationPath)
        # Collect information
        $DestinationChecksum = Get-FileHash -Path $DestinationPath -Algorithm 'SHA256'
        $DestinationVersion = [System.Version]((Get-Item -Path $DestinationPath).'VersionInfo'.'FileVersion')
        # Compare checmsums
        $ChecksumMatch = [bool]($LatestSha256 -eq $DestinationChecksum.'Hash')
        # Compare versions
        $VersionMatch = [bool]($DestinationVersion -ge $LatestVersion)
        # Determine if already installed
        $AlreadyInstalled = [bool]($ChecksumMatch -and $VersionMatch)
        # Output findings
        Write-Information -MessageData (
            'Latest version "{0}", installed version "{1}".' -f
            $LatestVersion.ToString(),
            $DestinationVersion.ToString()
        )
        Write-Information -MessageData ('Does SHA256 checksums match? {0}.' -f $AlreadyInstalled)
    }
    else {
        Write-Information -MessageData ('Destination "{0}" does not already exist.' -f $DestinationPath)
    }


    # Install
    ## Introduce step
    Write-Information -MessageData ('{0}# Install' -f [System.Environment]::NewLine)

    ## Check if script should proceed with install
    if (
        [bool]$(
            if ($AlreadyInstalled) {
                $ForceReinstall.'IsPresent'
            }
            else {
                $true
            }
        )
    ) {
        if ($PSCmdlet.ShouldProcess('jxl-winthumb', 'install')) {
            # Give info
            if ($AlreadyInstalled) {
                Write-Information -MessageData '-WhatIf was not specified, and -ForceReinstall was specified, will install.'
            }
            else {
                Write-Information -MessageData '-WhatIf was not specified, will install.'
            }

            # Download
            ## Introduce step
            Write-Information -MessageData ('{0}# Download' -f [System.Environment]::NewLine)

            ## Output download target location
            Write-Information -MessageData ('Download to "{0}"' -f $DownloadPath)

            ## Delete item if it already exists
            if ([System.IO.File]::Exists($DownloadPath)) {
                $null = [System.IO.File]::Delete($DownloadPath)
            }

            ## Download
            $null = [System.Net.WebClient]::new().DownloadFile(
                $LatestFileUri,
                $DownloadPath
            )

            ## Failproof
            ### Downloaded file does not exist
            if (-not [System.IO.File]::Exists($DownloadPath)) {
                Throw 'Failed to download - Did not find the expected file.'
            }
            ### Downloaded file size is not equal to what GitHub says it should be
            if ((Get-Item -Path $DownloadPath).'Length' -ne $LatestFileSize) {
                Throw 'Failed to download - Downloaded file does not match expected file size.'
            }
            ### Downloaded file SHA256 does not match GitHub asset digest
            $DownloadChecksum = [string] (Get-FileHash -Path $DownloadPath -Algorithm 'SHA256').'Hash'
            if ($DownloadChecksum -ne $LatestSha256) {
                Throw (
                    'SHA256 from downloaded file ("{0}..") does not match GitHub asset digest ("{1}..").' -f
                    $LatestSha256.Substring(0, 5),
                    $DownloadChecksum.Substring(0, 5)
                )
            }

            ## Output success
            Write-Information -MessageData 'Success.'

            # Get explorer.exe path
            $ExplorerPath = [string]((Get-Process -Name 'explorer')[0].'path')
            if (-not $? -or [string]::IsNullOrWhitespace($ExplorerPath)) {
                Throw 'Failed to get path of explorer.exe.'
            }

            # Kill explorer.exe
            Write-Information -MessageData 'Kill explorer.exe and wait a second.'
            $null = Stop-Process -Name 'explorer' -Force
            Start-Sleep -Seconds 1

            # Copy new file, overwrite if it already exists
            Write-Information -MessageData 'Copy downloaded file to destination, overwrite if it already exists.'
            [System.IO.File]::Copy($DownloadPath, $DestinationPath, $true)

            # Register new file
            Write-Information -MessageData 'Register dll with regsvr32.'
            $Process = Start-Process -Path 'regsvr32' -ArgumentList '/s', ('"{0}"' -f $DestinationPath) -NoNewWindow -PassThru -Wait
            if (-not $? -or $Process.'ExitCode' -ne 0) {
                if ($Process.'ExitCode' -ne 0) {
                    $ExitCode = $Process.'ExitCode'
                }
                Throw ('regsvr32 failed with exit code "{0}".' -f $Process.'ExitCode')
            }

            # Start explorer
            Write-Information -MessageData 'Start explorer.exe again.'
            $null = Start-Process -FilePath $ExplorerPath
        }
    }
    else {
        Write-Information -MessageData 'Already installed and -ForceReinstall was not specified, will not install.'
    }
}
Catch {
    # Make sure explorer.exe runs
    if (-not (Get-Process -Name 'explorer' -ErrorAction 'Ignore')) {
        $null = Start-Process -FilePath $ExplorerPath
    }

    # Set exit code
    if (
        $ExitCode -eq 0 -and
        $null -ne $LASTEXITCODE -and
        $LASTEXITCODE -ne 0
    ) {
        $ExitCode = $LASTEXITCODE
    }

    # Set success to false
    $Success = [bool] $false

    # Write the catched error
    Write-Error -ErrorAction 'Continue' -ErrorRecord $_
}



# Clean up files
if (-not [string]::IsNullOrEmpty($DownloadPath) -and [System.IO.Path]::Exists($DownloadPath)) {
    # Introduce step
    Write-Information -MessageData ('{0}# Clean up files' -f [System.Environment]::NewLine)

    # Clean up files
    Write-Information -MessageData (
        'Deleting "{0}". Success? "{1}".' -f $DownloadPath, $(
            Try {
                $null = [System.IO.File]::Delete($DownloadPath)
                $?
            }
            Catch {
                $false
            }
        ).ToString()
    )
}



# Exit
## Introduce step
Write-Information -MessageData ('{0}# Exit' -f [System.Environment]::NewLine)

## Exit
if ($Success) {
    Write-Output -InputObject 'Done.'
    Exit 0
}
else {
    Write-Error -ErrorAction 'Continue' -Message ('Failed. $ExitCode = "{0}", $LASTEXITCODE = "{1}".' -f $ExitCode, $LASTEXITCODE)
    Exit $(if($ExitCode -and $ExitCode -ne 0){$ExitCode}else{1})
}
