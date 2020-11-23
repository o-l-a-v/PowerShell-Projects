# Changelog
## v1.9.1.0 201003
### Additions
* Added check to see if powershellgallery.com is up and responding.
* Added ability to harcode module versions to keep/ now remove.

### Fixes

### Improvements



## v1.9.0.0 200523
### Additions
* Function "Uninstall-ModuleManually" to speed up uninstall of outdated modules.
* More useful modules to install.

### Fixes

### Improvements
* Corrected info about some of the modules in the list of modules to install.



## v1.8.0.0 200407
### Additions

### Fixes

### Improvements
* Now using "PackageManagemeng\Install-Package" instead of "Install-Module", because "Install-Module" does not set error variable $? correctly if it fails.
* Some speed ups
	* Use ".'Property'" instead of "Select-Object -ExpandProperty 'Property'



## v1.7.0.0 200323
### Additions

### Fixes
* Install-MissingSubModules
	* Lists a installed module thats not available from PowerShell Gallery anymore as missing sub module.
		* "Compare-Object" returns objects not in reference object AND difference object.
		* Use .Where instead

### Improvements
* Syntax fix, quotation marks on all dot properties
* Some speed ups using .Where.



## v1.6.0.0 191128
### Additions
* Added setting for modules you don't want to get updated. 
	* $ModulesDontUpdate.

### Fixes

### Improvements



## v1.5.2.0 191127
### Additions
* Added stats.

### Fixes
* Minor code refactoring.

### Improvements



## v1.5.1.0 191114
### Additions
* Added "human readable" time start and time end output at the end, who reads time in the ToString('o') format anyway.

### Fixes

### Improvements
* Uses Begin, Process and End in every function.
* Use [OutputType] in each function.
* Better markdown syntax for changelog.



## v1.5.0.0 191022
### Additions
* Greatly improved speed by writing my own function for getting all installed versions of a module.

### Fixes

### Improvements



## v1.4.1.0 190808
### Additions
* Added option to automatically accept licenses when installing modules. Suddenly saw the first module requiring this: "Az.ApplicationMonitor".

### Fixes

### Improvements
* Fixed code style places I saw it lagged behind. Esthetics.



## v1.4.0.0 190619
### Additions

### Fixes

### Improvements
* Prerequirements are now handled automatically in the script, no need to flip a boolean for that anymore.



## v1.3.4.0 190616
### Additions
* Better handling of uninstalling PackageManagement if module is updated during current session.
	* If PackageManagement was updated during the same session, PowerShell must be closed and reopened before outdated version can be removed.
* Will make sure that the user does not try to uninstall "PackageManagement" or "PowerShellGet".

### Fixes

### Improvements



## v1.3.3.0 190524
### Additions

### Fixes
* Uninstall unwanted modules did not find any modules because the script scoped variable for installed modules is not just keeping the name of the modules anymore. Fixed with a Select-Object.
* Other bugfixes since last release, lots of small stuff I don't remember. Nothing major.

### Improvements
* Better logic in the update installed modules function when a parent module is updated and submodules might have been updated aswell.



## v1.3.2.0 190503
### Additions
* Added check for Execution Policy, will attempt to fix it if necessary by setting it to "Unrestricted" just for current process.

### Fixes

### Improvements



## v1.3.1.0 190404
### Additions
* Added modules
	* ImportExcel, can be used to import/modify/export to/from Excel files

### Fixes
* Would fail on a clean install of Windows 10, fixed now

### Improvements



## v1.3.0.0 190401
### Additions

### Fixes
* Install-MissingSubModules
	* Used punctuation as check weather a module was a sub module. This does not work for "Microsoft.Graph.Intune" for instance. Now it removes the last punctuation and name, and checks weather this name exist in the list of installed modules.

### Improvements
* UpdateInstalledModules
	* If a submodule, punctuation in the name, will check once again what version is installed, in case parent module updated it.



## v1.2.0.1 190331
### Additions

### Fixes
* Used "$InstallMissingModules" instead of "$InstallMissingSubModules" for controlling weather the script should add missing sumodules. Type'o.

### Improvements



## v1.2.0.0 190324
### Additions
* Added three new modules
	* IntuneBackupAndRestore. John Seerden. Uses "MSGraphFunctions" module to backup and restore Intune config.
	* MSGraphFunctions. John Seerden. Wrapper for Microsoft Graph Rest API.
	* PSScriptAnalyzer. Microsoft. Used to analyze PowerShell scripts to look for common mistakes + give advice.

### Fixes
* Missed a "#" on one of the output headings
* Install missing modules: Success status after install would not display.

### Improvements



## v1.1.0.0 190316
### Additions
* Added changelog while I still remember the version history and changes.

### Fixes
* Added ability to check for and install missing submodules.
	* For instance: Az currently has 81 submodules, I only had about 60 installed. Now the script will add all submodules available from PowerShellGallery not already installed on the computer.
	* Controllable by boolean $InstallMissingSubModules

### Improvements



## v1.0.1.0 190313
### Additions
* Added ability to install prerequirements
	* NuGet (Package Provider)
	* PowerShellGet (PowerShell Module)
	* Controllable by boolean $InstallPrerequirements

### Fixes

### Improvements



## v1.0.0.0 190310
* Initial Release