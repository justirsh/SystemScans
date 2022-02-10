# SystemScans
These scripts can be used to scan a system and gather maintenance requirements.

-The osScan.ps1 will reference a list file to remotely scan Windows machines using winRM.  The purpose is to automate the gathering of running Operating Systems and their build versions for remote desktops and servers.
Output is read out to the console.

-The appScan.ps1 will reference a list file to remotely scan Windows machines using winRM.  The purpose is to automate the gathering of installed applications and their versions on remote desktops and servers.  appScan has osScan worked into it.  Still testing.
Output is exported to an Excel spreadsheet.
