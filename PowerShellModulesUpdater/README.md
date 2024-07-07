# PowerShellModulesUpdater

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
2. Windows PowerShell 5.1: Directly, in PowerShell ISE, or in Visual Studio Code with the PowerShell extension.

Since v1.18.0 the script runs on both Windows and Linux. Probably MacOS too, but I don't have hardware to test on.

## Goals for future version

### Short term

#### Can be done now

* [x] Use `Microsoft.PowerShell.PSResourceGet` now that it's GA.
* [x] Self contained with `Microsoft.PowerShell.PSResourceGet` as only module dependency.
  * Installed to a separate location than where modules and scripts are installed.
* [ ] Add a oneliner to the `README.md` to download and run the script.
* [ ] Use Dev Drive on Windows if present for both `-TemporaryPath` and `-Path` with `Save-PSResource` for greater speed and hopefully not be blocked by antimalware realtime scanning.
  * Related issue: <https://github.com/PowerShell/PSResourceGet/issues/1662>
  * Related info about Dev Drive and Performance Mode: <https://learn.microsoft.com/en-us/defender-endpoint/microsoft-defender-endpoint-antivirus-performance-mode>

#### Depends on others

* [ ] Stop using `Microsoft.PowerShell.PSResourceGet\Save-PSResource` (and `PackageManagement\Save-Package` before it) to install modules in user context as soon as `Microsoft.PowerShell.PSResourceGet` supports user specified modules location.
  * 181018: <https://github.com/PowerShell/PowerShell/issues/8069>
  * 210609: <https://github.com/PowerShell/PowerShell/issues/15552>
  * 220407: <https://github.com/PowerShell/PSResourceGet/issues/627>

### Long term

#### Can be done now

* [x] Linux support.
* [x] Speed up by not doing everything sequentially.
  * Modules not sharing name or author should be able to install and update simultaneously.
    * Find out all modules to install, then install with `-SkipDependencyCheck` in parallel?
      * Easy with PowerShell >= 7.3.X `ForEach-Object -Parallel`, not so easy with Windows PowerShell.
        * Runspace pool? Module `PoshRSJob`?
        * Speed ups in `Microsoft.PowerShell.PSResourceGet`?
          * 180525 - "Find-Module -Command need perf improvement": <https://github.com/PowerShell/PSResourceGet/issues/31>
          * 190726 - "[Feature Request] Parallel Dependency installations": <https://github.com/PowerShell/PSResourceGet/issues/69>
          * 230402 - "Find-PSResource with a string array of module names is slow: Bulk/batch the API calls?": <https://github.com/PowerShell/PSResourceGet/issues/1045>
* [ ] Have user preferences of what modules to install outside the actual script.
* [ ] Publish to PSGallery as script or module.

#### Depends on others

* [ ] Use `Microsoft.PowerShell.PSResourceGet` for uninstalling modules too.
  * Can't currently uninstall modules saved to a path outside the hardcoded paths for `AllUsers` and `CurrentUser` scopes.

## Findings

### PackageManagement and PowerShellGet

* Module install directory path is hard coded in both scopes (AllUsers, CurrentUser), both for:
  * Installing modules
    * Workarounds:
      * PackageManagement: `Save-Package`
      * PowerShellGet v3: `Save-PSResource`
      * PowerShellGet v2: `Save-Module`
    * Related issues:
      * <https://github.com/PowerShell/PowerShellGet/issues/627>
  * Searching for installed modules
    * Workarounds:
      * Custom logic: Modules versions folders where the hidden file `PSGetModuleInfo.xml` is present.
      * `Microsoft.PowerShell.Core\Get-Module`.
    * Related issues:
      * <https://github.com/PowerShell/PowerShellGet/issues/889>
* Can't skip dependencies when installing a module, these old versions of commonly used modules often get installed at the same time.
  * Coming with PowerShellGet v3. Related issue: <https://github.com/PowerShell/PowerShellGet/issues/620>

### Dotnet

* Does not support semantic versioning v2.0.0
  * Dotnet issue: <https://github.com/dotnet/runtime/issues/19317>
  * PowerShell issue: <https://github.com/PowerShell/PowerShell/issues/2983>

### PowerShell Gallery API

* Stuck on NuGet API v2
  * Poorly documented.
  * A lot of useful filtering and searching functionality does not work.
* `sortBy=Version desc` does not work.

## Known issues

### Does not handle non-stable versioning

#### Why

* Some version numbers can't be parsed/ read as `[System.Version]`.

#### What problem does it cause

* Can't handle version numbers that can't be parsed as `[System.Version]`.
  * Example: `PowerShellGet` beta version `3.0.12-beta`.
* No logic for handling scenarios like:
  * Pre-release version and stable version installed of the same module.

#### Dependencies I'd like to wait for

* `PowerShellGet` and `PackageManagement` currently does not have an easy way of telling whether a version should be considered pre-release or stable.

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

##### Option 2 - PowerShellGallery API to get release info

###### How

* Get version info for all installed versions, and newest available version, using PowerShellGallery API.

###### Cons

* Slow
