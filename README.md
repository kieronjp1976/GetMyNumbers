# Instructions
Calculates a ringers totals in the style of the old IPMR report, uses data export from demon

This script does not filter on dates, this will need to be done when exporting the data

Kieron Palmer

## Instructions

This has been tested on Windows 10 with the latest version of MS Excel.

This was written for my own use but you are free to use without support from me.

1/ Download the script from Github (The .ps1 file)

2/ Export the data you need from Demon

3/ Edit the script ( Notepad is OK to use) to point to the demon export and to amend the ringers initials - this is labelled at the top of the script

4/ Open PowerShell as an Admin. (Type powershell into the start menu and right click it, select run as admin)

5/ Run "set-executionpolicy remotesigned" and choose Y when prompted/ This allows local scripts to run.

6/ Change to the directory that you saved the script in  (eg Type "cd c:\temp")

7/ Type in .\GetMyNumbers.ps1

8/ The script will create an excel file and a pdf in the path that you specified in $path and $pdf
