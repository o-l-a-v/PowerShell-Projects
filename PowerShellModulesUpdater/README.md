# PowerShellModulesUpdater

> [!CAUTION]
> Don't run random scripts from the internet.
>
> * This script is provided as-is, you should run it with caution.
> * If it breaks anything on your computer it's your fault.

## About

A script that installs, updates (and cleans up old versions), and removes PowerShell modules and scripts.

The script is self contained and maintains its' own version of the `Microsoft.PowerShell.PSResourceGet` module:

* Self contained in `%LOCALAPPDATA%\Microsoft\PowerShell\PowerShellModulesUpdater`, outside of where other modules are installed.
* Installed manually by downloading latest stable version from PowerShellGallery.
* Ensures the ability to:
  * Update any module, also `Microsoft.PowerShell.PSResourceGet`.
  * Always use latest stable `Microsoft.PowerShell.PSResourceGet`, whether the user wants it in its' `$env:PSModulePath` or not.
* Since PSScriptAnalyzer loads modules and DLLs in the background for doing script analysis, the script have to check if an older `PSResourceGet` assembly / DLL is already loaded in the session. If it is the script must use whatever version of `PSResourceGet` is already loaded.

## History of major changes

The script originally ran on Windows using Windows PowerShell only.

Since v1.17.0 the script can run on both (1) PowerShell and (2) Windows PowerShell.

1. PowerShell \>= 7.3.X: Directly, or in Visual Studio Code with the PowerShell extension.
1. Windows PowerShell 5.1: Directly, in PowerShell ISE, or in Visual Studio Code with the PowerShell extension.

Since v1.18.0 the script runs on both Windows and Linux. Probably MacOS too, but I don't have hardware to test that.

## Findings

### PSResourceGet, PowerShellGet and PackageManagement

Module install directory path is hard coded in both scopes (AllUsers, CurrentUser), both for:

* Installing modules
  * Workarounds:
    * PackageManagement: `Save-Package`
    * PowerShellGet v3: `Save-PSResource`
    * PowerShellGet v2: `Save-Module`
  * Related issues:
    * 2022-04-07: <https://github.com/PowerShell/PSResourceGet/issues/627>
* Searching for installed modules
  * Workarounds:
    * Custom logic: Modules versions folders where the hidden file `PSGetModuleInfo.xml` is present.
    * `Microsoft.PowerShell.Core\Get-Module`.
    * `Microsoft.PowerShell.PSResourceGet\Get-InstalledPSResource` with `-Path`.
  * Related issues:
    * 2023-01-12: <https://github.com/PowerShell/PSResourceGet/issues/889>

### Dotnet

Does not support semantic versioning v2.0.0

* Dotnet issue 2016-11-09: <https://github.com/dotnet/runtime/issues/19317>
* PowerShell issue 2017-01-10: <https://github.com/PowerShell/PowerShell/issues/2983>

### PowerShell Gallery API

* Stuck on NuGet API v2
  * Poorly documented.
  * A lot of useful filtering and searching functionality does not work.
  * Meta issue and feature request for PowerShell modules to be hosted on NuGet Gallery: <https://github.com/NuGet/NuGetGallery/issues/10071>.
* `sortBy=Version desc` does not work.

## Known issues

### Does not handle non-stable versioning

#### Why

* Some version numbers can't be parsed/read as `[System.Version]` and PowerShell does not have proper semantic versioning support, nor NuGet.Versioning.

#### What problems does it cause

* Can't handle version numbers that can't be parsed as `[System.Version]`.
  * Example: `PowerShellGet` beta version `3.0.12-beta`.
* No logic for handling scenarios like:
  * Pre-release version and stable version installed of the same module.

#### Fix short term

* Don't validate such version numbers at all.

#### Fix long term

##### Option 1 - Parse version info from `PackageManagement` manually

###### How

If raw version can't be parsed, manually parse it.

* Regex to remove info that can't be parsed as `[System.Version]`, store it somewhere.
####### Ideas
* Create own class for handling version numbers that `[System.Version]` does not handle?
* Add a attribute to the object "nonstable", if raw version can't be parsed as `[System.Version]`?

##### Option 2 - PowerShell Gallery API to get release info

###### How

* Get version info for all installed versions, and newest available version, using PowerShellGallery API.

###### Cons

* Slow
