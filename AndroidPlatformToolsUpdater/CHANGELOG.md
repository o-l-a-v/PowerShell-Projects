# Changelog
## v210910
### Additions
* Get available version from Google webpage by downloading it and parsing the HTML
  * Logic might break, then the script will just depend on the install function to check installed version vs. the downloaded vesion.

### Fixes

### Improvements
* Better output to terminal.



## v201214
### Additions
* Use 7-Zip if installed and found.

### Fixes
* Earlier variable ```$SystemContext``` changed name to ```$SystemWide``` but forgot to update it several places.

### Improvements
* Added ```[OutputType()]``` to all functions.
* All functions are now context aware.
* Cleanup downloaded and extracted file when done.
* More output to terminal.
* Syntax cleanup.



## v191005
### Additions

### Fixes
* Fixed broken variables due renaming + not testing it from a clean PowerShell session after done coding for previous version.

### Improvements



## v190428
### Additions

### Fixes
* Some bugfixes to Environmental Variables

### Improvements



## v190310
* Initial version