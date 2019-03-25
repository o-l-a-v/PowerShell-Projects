# Changelog
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