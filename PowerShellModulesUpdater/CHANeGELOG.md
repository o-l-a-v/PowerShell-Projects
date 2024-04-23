# Changelog

## About

* Changelog format follows [Keep a Changelog](https://keepachangelog.com/en).
* Versioning adheres to [Semantic Versioning](https://semver.org/).

## [1.19.0-beta2] - 2024-04-15

### Added

* Learned that PowerShell searches for scripts in the PATH environment variable.
  * `Set-PSModulePathUserContext` will now also add wanted scripts path to user context PATH environment variable.

## [1.19.0-beta1] - 2024-03-27

### Added

* Function `Find-PSGalleryPackageLatestVersionUsingApiInBatch` to find latest version of resources from PowerShell Gallery in batches of up to 30 at a time.
  * `Microsoft.PowerShell.PSResourceGet` cmdlet `Find-PSResource` can only do one at a time, even if provided a string array of package IDs.
* Function `Save-PSResourceInParallel` to use PowerShell runspace factory to parallelize `Save-PSResource` to install multiple PowerShell resources at a time.
  * `Microsoft.PowerShell.PSResourceGet` cmdlet `Save-PSResource` can only do one at a time, even if provided a string array of package IDs.

### Changed

* Huge speedups by parallization in both querying the PowerShell Gallery, and installing packages.
* Greatly reduced vervosity and amount of script output by only showing info about packages that will be installed instead.
  * Previously the script iterated all packages, and outputted whether to install/update or not do anything.

## [1.18.2] - 2024-03-08

### Fixed

* Indentation in `Write-Statistics` was not following the new way introduced in v1.18.0.

## [1.18.1] - 2024-03-02

### Changed

* Versioning is now using [Semantic Versioning](https://semver.org/) again.

### Fixed

* In PowerShell 5.1, doing string `.Contains()` on an empty attribute throws an error.
  * Fix: Do `-not [string]::IsNullOrEmpty($_.'<attribute>')` before `$_.'<attribute>'.Contains('<something>')`.

## [1.18.0] - 2024-02-18

### Added

* Linux support: Very much beta, not properly tested.

### Changed

* Make script output indentation / tabulation uniform across terminals by hardcoding indention to two whitespaces with varible `$Indentation`.
  * Because it's up to the host terminal application to handle tab expansion.
    * <https://github.com/MicrosoftDocs/PowerShell-Docs/issues/10533>
  * Tab expansion varies between PowerShell ISE (4 whitespaces) and the regular PowerShell terminal (8 whitespaces).

### Fixed

* PSScriptAnalyzer complaints:
  * Use approved verbs for functions.
    * `Output-Statistics` -> `Write-Statistics`
    * `Refresh-ModulesInstalled` -> `Get-ModulesInstalled`.
  * Use `-ErrorAction SilentlyContinue` because PSScriptAnalyzer falsefully detects it as `PSPossibleIncorrectUsageOfRedirectionOperator`.
* Exiting functions that has `[OutputType([System.Void])]` using `Break` exited the whole script, now using `return` instead.

## [1.17.0] - 2024-02-08

### Added
* User context: Avoid installing to OneDrive (KFM) by overriding default install path `%USERPROFILE%\Documents\WindowsPowerShell\Modules` by:
  * Set user context environmental variable `PSModulePath` to `%LOCALAPPDATA%\Microsoft\PowerShell\Modules`.
  * Use `Microsoft.PowerShell.PSResourceGet\Save-PSResource` to override install path.
  * Use own logic for detecting installed modules given scope.
    * More info on how: <https://github.com/PowerShell/PSResourceGet/issues/627#issuecomment-1380881825>
* Input parameter `-DoScripts` to control whether to touch PowerShell scripts.
* Detect and use conflicting assembly for `Microsoft.PowerShell.PSResourceGet` if already present in the session. Related issues:
  * 190228: <https://github.com/PowerShell/PSScriptAnalyzer/issues/1154>
  * 240208: <https://github.com/PowerShell/PowerShell/issues/21201>
  * 240208: <https://github.com/PowerShell/PowerShell/issues/21199>

### Fixed
* Function `Refresh-ModulesInstalled`: Wrongfully listed prerelease versions.
* Use `[OutputType([System.Void])]` instead of `[OutputType($null)]` for functions that returns no output.
* Detect, install and update scripts in non-standard directory.

### Changed
* Now using the new `Microsoft.PowerShell.PSResourceGet` module instead of `PackageManagement` and `PowerShellGet`.
  * Self contained in `%LOCALAPPDATA%\Microsoft\PowerShell\PowerShellModulesUpdater`.
* Made the script compatible with PowerShell Core (> 5.1) by:
  * Removed `Requires -Edition Desktop`.
  * Use `Test-Connection` if PowerShell Core, else `Test-NetConnection`.
* Installing to user context now actually works, for both modules and scripts. More info on how under "Added".
* Changed changelog syntax to ["Keep a Changelog"](https://keepachangelog.com/en), except for version numbers.

## [1.16.0] - 2022-10-20

### Changed
* Created a test to see if ```-AcceptLicense``` parameter is available for command ```PackageManagement\Install-Package```.
* Converted some of the settings to input parameters (```param()```).

## [1.15.0] - 220320

### Added
* Make sure not to use PowerShellGet v3 if installed.
  * It's not ready yet IMO, and it's still in beta.

### Fixed
* Script would not load the newest versions of required modules ```PowerShellGet``` and ```PackageManagement```.
* Implemented short term fix for version numbers that can't be read as ```[System.Version]```, like prereleases.
  * For instance, ```PowerShellGet v3.0.12-beta```.
  * Short term fix: Ignore such version numbers.
  * See README.md for more info.


## [1.14.0] - 2021-12-20

### Added
* Created a ```README.md```.

### Fixed
* Quick fix to handle beta/ pre-release versions
  * Like [```PowerShellGet``` pre-release ```3.0.12-beta```](https://www.powershellgallery.com/packages/PowerShellGet/3.0.12-beta).
  * In other words, version numbers that can't be parsed as ```[System.Version]```.
  * Workaround: Don't validate such version numbers at all.
  * Future: See ```README.md```.

## [1.13.0] - 2021-09-14

### Added
* Added installing scripts from PowerShellGallery defined in variable $ScriptsWanted.
* Added updating already installed scripts that origin from PowerShellGallery.

### Changed
* Some touchups on output to reflect that the script now does both modules and scripts.

## [1.12.0] - 2021-08-19

### Added
* Started on the mission of enabling the script to install modules either in user (CurrentUser) or system (AllUsers) context.
  * System context works, user context is not properly tested yet.

### Fixed
* Check for prerequired components and modules:
  * Better logic on using built in ```PowerShellGet``` and ```PackageManagement``` binaries until their modules are installed from PowerShellGallery.

### Changed

## [1.11.0] - 2021-06-14

### Changed
* Added quicker check for prerequired components and modules:
  * If ```PackageManagement``` exists as module type ```Script```, not just ```Binary```, we should be good.

## [1.10.0] - 2021-05-03

### Added
* ~~Script version is now just ```YYMMdd``` for simplicity.~~
  * Changed back to semantic versioning later.
* On a fresh client where the module has never run before, script can now install prerequirements and continue with the script in one run by loading newest versions of the prerequirements automatically.
* Can now update and uninstall old version of ```PowerShellGet``` in one run, by simply:
  * Update to newer version if available.
  * Unload (```Remove-Module -Force```) all versions from current session.
  * Import (```Import-Module```) newest installed version to current session.
  * Uninstall old version.

### Fixed
* Fixed some buggy text output to terminal.

### Changed
* ```Write-Information``` instead of ```Write-Output``` for outputting text to terminal.
  * ```Write-Output``` is for returning objects.
  * ```Write-Information``` has no downsides with PowerShell 6 and newer, but for PowerShell 5.1 this output stream is:
    * Not enabled by default (requires ```$InformationPreference = 'Continue'```).
    * Will not be visable if logging using ```Start-Transcript``` function.
* Will always start by updating prerequirement PowerShell modules ```PackageManagement``` and ```PowerShellGet```.
* Using markdown syntax for headers for script output. One ```#``` for H1 etc.
* Some syntax cleanup.

## [1.9.1] - 2020-10-03

### Added
* Added check to see if powershellgallery.com is up and responding.
* Added ability to harcode module versions to keep/ now remove.

## [1.9.0] - 2020-05-23

### Added
* Function "Uninstall-ModuleManually" to speed up uninstall of outdated modules.
* More useful modules to install.

### Changed
* Corrected info about some of the modules in the list of modules to install.

## [1.8.0] - 2020-04-07

### Changed
* Now using "PackageManagemeng\Install-Package" instead of "Install-Module", because "Install-Module" does not set error variable $? correctly if it fails.
* Some speed ups
  * Use ".'Property'" instead of "Select-Object -ExpandProperty 'Property'

## [1.7.0] - 2020-03-23

### Fixed
* Install-MissingSubModules
  * Lists a installed module thats not available from PowerShell Gallery anymore as missing sub module.
    * "Compare-Object" returns objects not in reference object AND difference object.
    * Use .Where instead

### Changed
* Syntax fix, quotation marks on all dot properties
* Some speed ups using .Where.

## [1.6.0] - 2019-11-28

### Added
* Added setting for modules you don't want to get updated.
  * $ModulesDontUpdate.

## [1.5.2] - 2019-11-27

### Added
* Added stats.

### Fixed
* Minor code refactoring.

## [1.5.1] - 2019-11-14

### Added
* Added "human readable" time start and time end output at the end, who reads time in the ToString('o') format anyway.

### Changed
* Uses Begin, Process and End in every function.
* Use `[OutputType]` in each function.
* Better markdown syntax for changelog.

## [1.5.0] - 2019-10-22

### Added
* Greatly improved speed by writing my own function for getting all installed versions of a module.

## [1.4.1] - 2019-08-08

### Added
* Added option to automatically accept licenses when installing modules. Suddenly saw the first module requiring this: "Az.ApplicationMonitor".

### Fixed
* Fixed code style places I saw it lagged behind. Esthetics.

## [1.4.0] - 2019-06-19

### Changed
* Prerequirements are now handled automatically in the script, no need to flip a boolean for that anymore.

## [1.3.4] - 2019-06-16

### Added
* Better handling of uninstalling PackageManagement if module is updated during current session.
  * If PackageManagement was updated during the same session, PowerShell must be closed and reopened before outdated version can be removed.
* Will make sure that the user does not try to uninstall "PackageManagement" or "PowerShellGet".

## [1.3.3] - 2019-05-24

### Fixed
* Uninstall unwanted modules did not find any modules because the script scoped variable for installed modules is not just keeping the name of the modules anymore. Fixed with a Select-Object.
* Other bugfixes since last release, lots of small stuff I don't remember. Nothing major.

### Changed
* Better logic in the update installed modules function when a parent module is updated and submodules might have been updated aswell.

## [1.3.2] - 2019-05-03

### Added
* Added check for Execution Policy, will attempt to fix it if necessary by setting it to "Unrestricted" just for current process.

## [1.3.1] - 2019-04-04

### Added
* Added modules
  * ImportExcel, can be used to import/modify/export to/from Excel files

### Fixed
* Would fail on a clean install of Windows 10, fixed now

## [1.3.0] - 2019-04-01

### Fixed
* Install-MissingSubModules
  * Used punctuation as check weather a module was a sub module. This does not work for "Microsoft.Graph.Intune" for instance. Now it removes the last punctuation and name, and checks weather this name exist in the list of installed modules.

### Changed
* UpdateInstalledModules
  * If a submodule, punctuation in the name, will check once again what version is installed, in case parent module updated it.

## [1.2.0] - 2019-03-31

### Fixed
* Used "$InstallMissingModules" instead of "$InstallMissingSubModules" for controlling weather the script should add missing sumodules. Type'o.

## [1.2.0] - 2019-03-24

### Added
* Added three new modules
  * IntuneBackupAndRestore. John Seerden. Uses "MSGraphFunctions" module to backup and restore Intune config.
  * MSGraphFunctions. John Seerden. Wrapper for Microsoft Graph Rest API.
  * PSScriptAnalyzer. Microsoft. Used to analyze PowerShell scripts to look for common mistakes + give advice.

### Fixed
* Missed a "#" on one of the output headings
* Install missing modules: Success status after install would not display.

## [1.1.0] - 2019-03-16

### Added
* Added changelog while I still remember the version history and changes.

### Fixed
* Added ability to check for and install missing submodules.
  * For instance: Az currently has 81 submodules, I only had about 60 installed. Now the script will add all submodules available from PowerShellGallery not already installed on the computer.
  * Controllable by boolean $InstallMissingSubModules

## [1.0.1] - 2019-03-13

### Added
* Added ability to install prerequirements
  * NuGet (Package Provider)
  * PowerShellGet (PowerShell Module)
  * Controllable by boolean $InstallPrerequirements

## [1.0.0] - 2019-03-10
* Initial Release
