# PowerShellModulesUpdater
## About
A script that installs, updates (and cleans up), and removes PowerShell modules and scripts.

The script is made for Windows PowerShell 5.1, running inside PowerShell ISE.

Why not VSCode and PowerShell Core? I've not yet made the jump because:
* I'm doing a lot of client management/ Intune scripting.
* Some Microsoft cloud modules I rely on only works with Windows PowerShell 5.1.
* I like the simplicity of PS ISE.


## Known issues
### Does not handle non-stable versioning
#### Why
* Some version numbers can't be parsed/ read as ```[System.Version]```.

#### What problem does it cause
* Can't handle version numbers that can't be parsed as ```[System.Version]```.
  * Example: ```PowerShellGet``` beta version ```3.0.12-beta```.
* No logic for handling scenarios like:
  * Pre-release version and stable version installed of the same module.

#### Dependencies I'd like to wait for
* ```PowerShellGet``` and ```PackageManagement``` currently does not have an easy way of telling whether a version should be considered pre-release or stable.

#### Fix short term
* Don't validate such version numbers at all.

#### Fix long term
##### Option 1 - Parse version info from ```PackageManagement``` manually
###### How
If raw version can't be parsed, manually parse it.
* Regex to remove info that can't be parsed as ```[System.Version]```, store it somewhere.
####### Ideas
* Create own class for handling version numbers that ```[System.Version]``` does not handle?
* Add a attribute to the object "nonstable", if raw version can't be parsed as ```[System.Version]```?
##### Option 2 - PowerShellGallery API to get release info
###### How
* Get version info for all installed versions, and newest available version, using PowerShellGallery API.
###### Cons
* Slow


## Inspiration and similar projects
* https://github.com/JustinGrote/ModuleFast
