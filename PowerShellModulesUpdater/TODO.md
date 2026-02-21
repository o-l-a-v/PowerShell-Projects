# Todo

## Small effort

### Can be done now

* [x] Use `Microsoft.PowerShell.PSResourceGet` now that it's GA.
* [x] Self contained with `Microsoft.PowerShell.PSResourceGet` as only module dependency.
  * Installed to a separate location than where modules and scripts are installed.
* [ ] Add a oneliner to the `README.md` to download and run the script.
* [ ] Use Dev Drive on Windows if present for both `-TemporaryPath` and `-Path` with `Save-PSResource` for greater speed and hopefully not be blocked by antimalware realtime scanning.
  * Related issue: <https://github.com/PowerShell/PSResourceGet/issues/1662>
  * Related info about Dev Drive and Performance Mode: <https://learn.microsoft.com/en-us/defender-endpoint/microsoft-defender-endpoint-antivirus-performance-mode>
* [ ] Speed up discovery of installed modules.
* [ ] Speed up uninstall of outdated module versions.

### Depends on others

* [ ] Stop using `Microsoft.PowerShell.PSResourceGet\Save-PSResource` (and `PackageManagement\Save-Package` before it) to install modules in user context as soon as `Microsoft.PowerShell.PSResourceGet` supports user specified modules location.
  * Issue 2018-10-18: <https://github.com/PowerShell/PowerShell/issues/8069>
  * Issue 2021-06-09: <https://github.com/PowerShell/PowerShell/issues/15552>
  * Issue 2022-04-07: <https://github.com/PowerShell/PSResourceGet/issues/627>
  * PR draft 2024-07-13: <https://github.com/PowerShell/PSResourceGet/pull/1673>
* [ ] Re-add ability to choose whether to `-AcceptLicense` by default, once PSResourceGet `Save-PSResource` get's it in a stable release.
  * Issue 2023-10-17: <https://github.com/PowerShell/PSResourceGet/issues/1450>
  * PR 2024-10-06: <https://github.com/PowerShell/PSResourceGet/pull/1718>

## Large effort

### Can be done now

* [x] Linux support.
* [x] Speed up by not doing everything sequentially.
  * Modules not sharing name or author should be able to install and update simultaneously.
    * Find out all modules to install, then install with `-SkipDependencyCheck` in parallel?
      * Easy with PowerShell >= 7.3.X `ForEach-Object -Parallel`, not so easy with Windows PowerShell.
        * Runspace pool? Module `PoshRSJob`?
        * Speed ups in `Microsoft.PowerShell.PSResourceGet`?
          * 2018-05-25 - "Find-Module -Command need perf improvement": <https://github.com/PowerShell/PSResourceGet/issues/31>
          * 2019-07-26 - "[Feature Request] Parallel Dependency installations": <https://github.com/PowerShell/PSResourceGet/issues/69>
          * 2023-04-02 - "Find-PSResource with a string array of module names is slow: Bulk/batch the API calls?": <https://github.com/PowerShell/PSResourceGet/issues/1045>
* [ ] Have user preferences of what modules to install outside the actual script.
* [ ] Publish to PSGallery as script or module.
* [ ] Write actual progress bar, either `Write-Progress` or PwshSpectreConsole?

### Depends on others

* [ ] Use `Microsoft.PowerShell.PSResourceGet` for uninstalling modules too.
  * Can't currently uninstall modules saved to a path outside the hardcoded paths for `AllUsers` and `CurrentUser` scopes.
