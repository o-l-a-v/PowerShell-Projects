# Changelog

## About

* Changelog format follows [Keep a Changelog](https://keepachangelog.com/en).
* Versioning adheres to [Semantic Versioning](https://semver.org/).

## v1.2.2 - 2024-12-20

### Fixed

* The logic for adding and cleaning up env variables.

## v1.2.1 - 2024-12-19

### Fixed

* Exclude PSSCriptAnalyzer rule `Add-AndroidPlatformToolsToEnvironmentVariables` for function `Add-AndroidPlatformToolsToEnvironmentVariables`.

## v1.2.0 - 2021-09-10

### Added

* Get available version from Google webpage by downloading it and parsing the HTML
  * Logic might break, then the script will just depend on the install function to check installed version vs. the downloaded vesion.

### Changed

* Better output to terminal.

## v1.1.0 - 2020-12-14

### Added

* Use 7-Zip if installed and found.
* ```[OutputType()]``` to all functions.
* More output to terminal.
* Cleanup downloaded and extracted file when done.

### Changed

* All functions are now context aware.
* Syntax cleanup.

### Fixed

* Earlier variable ```$SystemContext``` changed name to ```$SystemWide``` but forgot to update it several places.

## v1.0.2 - 2019-10-05

### Fixed

* Fixed broken variables due renaming + not testing it from a clean PowerShell session after done coding for previous version.

## v1.0.1 - 2019-04-28

### Fixed

* Some bugfixes to Environmental Variables

## v1.0.0 - 2019-03-10

* Initial version
