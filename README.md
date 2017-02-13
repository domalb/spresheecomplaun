# spresheecomplaun
Launcher for spreadsheetcompare.exe

Allows invoking spreadsheetcompare.exe for excel files comparison using basic command line, e.g. :
`sscl.exe <file1> <file2>`
This was originally done to enable xls files diff in version control software. 

Spreadsheetcompare.exe is searched using various methods:
- MSI component location
- Excel.exe location from registry key
- Default office installation folders
It can also be explicitly given in commandline using -d argument, e.g. :
`sscl.exe <file1> <file2> -d="C:\My Program Files (x86)\Microsoft Office\Office15\DCF"`

 
