v1.1.0.0 190316
	* Bugfixes
	* Added ability to check for and install missing submodules.
		* For instance: Az currently has 81 submodules, I only had about 60 installed. Now the script will add all submodules available from PowerShellGallery not already installed on the computer.
		* Controllable by boolean $InstallMissingSubModules


v1.0.1.0 190313
	* Bugfixes
	* Added ability to install prerequirements, controllable by boolean $InstallPrerequirements
		* NuGet (Package Provider)
		* PowerShellGet (PowerShell Module)


v1.0.0.0 190310
	* Initial Release