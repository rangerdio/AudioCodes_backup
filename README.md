# AudioCodes_backup
Simple app to extract AudioCodes configuration file on safe place.

Created while learning Python, using Python 3.9.13.
DIST exe file is created to be used directly on systems that do not have python installed. Builded with pyinstaller.

Used packages:
pandas 2.2.2
paramiko 3.3.1
colorama 0.4.6

Usage
- put the .exe file along with devices.xlsx in desired windows folder.
- supply "devices.xlsx" with the required ip address/username/password.
- start the app, or create scheduled task

- AudioCodes configuration files will be populated in the same directory, along with log file with details to track errors.
- Name contains the IP address and timestamp of creation, so version of the backup is uniqly identified by the timestamp.

Tested with following AudioCodes devoces:987
Tested on:
  - Windows 10 Pro
  - Windows 11 Pro

Important note:
App will not directly work on "Windows 7 SP1" and "Windows Server 2008 R2" due to compitability issue of Python 3.9+ with these old systems.
When Ran, an error for missing  / python39.dll / api-ms-win-core-path-l1-1-0.dll is missing from your computer /
I was able to workaround this by using "https://github.com/adang1345/PythonWin7" on target Windwos Server 2008 R2 (tested with 3.10.13 on Windows Server 2008 R2).
