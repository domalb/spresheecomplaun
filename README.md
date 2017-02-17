# Spresheecomplaun (SSCL)

Utility to ease Excel files comparison with Spreadsheetcompare.
More specifically, enables Excel files diff in version control softwares

## Spreadsheetcompare

Spreadsheetcompare is a tool available in Office which is not command-line friendly.
The problem was reported in Microsoft TechCenter: http://social.technet.microsoft.com/Forums/office/en-US/65a0b5f1-e58f-4916-b090-551bbf9c719d/problem-with-spreadsheetcompare?forum=officeitpro

SSCL is a launcher for Spreadsheetcompare.
It consists in a small executable with reworked arguments, that spawns Spreadsheetcompare.

Using SSCL enables diff of xls files in version control software (like Git, Perforce, SVN...).

## SSCL Syntax

### Basic

Basic command line is :

`sscl.exe C:\book1.xlsx C:\book2.xlsx`

- Filenames are expected to be absolute pathes.
- Filenames may have double quotes

### Locate Spreadsheetcompare.exe

Spreadsheetcompare.exe is searched using various methods:

1. MSI component location
2. Excel.exe location from registry key
3. Default office installation folders

It can also be explicitly given in commandline using `-d` argument, e.g. :

`sscl.exe C:\book1.xlsx C:\book2.xlsx -d="C:\My Program Files (x86)\Microsoft Office\Office15\DCF"`

### Pause

To enable easy debugger attachment, sscl.exe can be paused at start using '-p' argument, e.g. :

`sscl.exe C:\book1.xlsx C:\book2.xlsx -p`

### Verbose

Log more information using '-v' argument, e.g. :

`sscl.exe C:\book1.xlsx C:\book2.xlsx -v`

## Related project

[Excomp](https://githubusercontent.com/kniklas/excomp) Excel compare script - batch file that allows execution of Excel 2013 SPREADSHEETCOMPARE tool from command line with two files as arguments. SSCL was inpired from Excomp, but with some improuvements (Office folder detection, filenames quotes, debuggable)

