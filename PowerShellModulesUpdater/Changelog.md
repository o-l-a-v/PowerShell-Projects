# Changelog
## v1.3.2.0 190503
* Additions
	* Added check for Execution Policy, will attempt to fix it if necessary by setting it to "Unrestricted" just for current process.

## v1.3.1.0 190404
* Bugfixes
	* Would fail on a clean install of Windows 10, fixed now
* Additions
	* Added modules
		* ImportExcel, can be used to import/modify/export to/from Excel files

## v1.3.0.0 190401
* Bugfixes
	* Install-MissingSubModules
		* Used punctuation as check weather a module was a sub module. This does not work for "Microsoft.Graph.Intune" for instance. Now it removes the last punctuation and name, and checks weather this name exist in the list of installed modules.
* Additions
	* UpdateInstalledModules
		* If a submodule, punctuation in the name, will check once again what version is installed, in case parent module updated it.

## v1.2.0.1 190331
* Bugfixes
	* Used "$InstallMissingModules" instead of "$InstallMissingSubModules" for controlling weather the script should add missing sumodules. Type'o.

## v1.2.0.0 190324
* Bugfixes
	* Missed a "#" on one of the output headings
	* Install missing modules: Success status after install would not display.
* Additions
	* Added three new modules
		* IntuneBackupAndRestore. John Seerden. Uses "MSGraphFunctions" module to backup and restore Intune config.
		* MSGraphFunctions. John Seerden. Wrapper for Microsoft Graph Rest API.
		* PSScriptAnalyzer. Microsoft. Used to analyze PowerShell scripts to look for common mistakes + give advice.

## v1.1.0.0 190316
* Bugfixes
* Added ability to check for and install missing submodules.
	* For instance: Az currently has 81 submodules, I only had about 60 installed. Now the script will add all submodules available from PowerShellGallery not already installed on the computer.
	* Controllable by boolean $InstallMissingSubModules
* Added changelog while I still remember the version history and changes.


## v1.0.1.0 190313
* Bugfixes
* Added ability to install prerequirements
	* NuGet (Package Provider)
	* PowerShellGet (PowerShell Module)
	* Controllable by boolean $InstallPrerequirements


## v1.0.0.0 190310
* Initial Release