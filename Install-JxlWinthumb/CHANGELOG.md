# Changelog

## About

* Changelog format follows [Keep a Changelog](https://keepachangelog.com/en).
* Versioning adheres to [Semantic Versioning](https://semver.org/).

## v1.3.0 - 2026-02-21

### Added

* Optionally escalate using gsudo if present and not `-WhatIf`.

### Changed

* Get SHA256 from GitHub asset digest, don't have to download anything to get the checksum.
* Reduce unneccessary verbosity in some if checks.

### Fixed

* Did not clean input parameter `-WriteChanges` after `-WhatIf` what implemented.

## v1.2.0 - 2024-12-19

### Added

* Failproof: Make sure script runs on Windows.

### Changed

* Use PowerShell native `-WhatIf` instead of custom boolean `-WriteChanges`.

## v1.1.0 - 2024-11-02

### Added

* Check installed version vs. available version on GitHub.
  * Required following feature request to be fixed: <https://github.com/saschanaz/jxl-winthumb/issues/40>.

## v1.0.0 - 2024-03-28

* Initial version.
