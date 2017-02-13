#include <windows.h>
#include <msi.h>
#include <iostream>
#include <fstream>

#define ARG_DIRECTORY L"-d="
static const size_t argDirectoryLength = sizeof(ARG_DIRECTORY) / sizeof(wchar_t) - 1;

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool folderExists(const wchar_t* folderPath)
{
	DWORD dwAttributes = ::GetFileAttributesW(folderPath);
	return ((dwAttributes != INVALID_FILE_ATTRIBUTES) && (dwAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0);
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool fileExists(const wchar_t* filePath)
{
	DWORD dwAttribs = GetFileAttributes(filePath);
	return ((dwAttribs != INVALID_FILE_ATTRIBUTES) && ((dwAttribs & FILE_ATTRIBUTE_DIRECTORY) == 0));
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool testSpreadsheetCompareFileInfolder(const wchar_t* folder)
{
	wchar_t testPath[MAX_PATH];
	wcscpy_s(testPath, folder);
	wcscat_s(testPath, L"\\SPREADSHEETCOMPARE.EXE");
	return fileExists(testPath);
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
int getSpreadsheetCompareFolder(wchar_t* buffer)
{
	// Locate component
	static const wchar_t* const office = L"{00000409-78E1-11D2-B60F-006097C998E7}";
	// 	const wchar_t* office2013 = L"{91150000-0011-0000-0000-0000000FF1CE}";
	// 	const wchar_t* excel = L"{CC29E96F-7BC2-11D1-A921-00A0C91E2AA2}";
	DWORD officeFolderLength = MAX_PATH;
	INSTALLSTATE install = MsiLocateComponentW(office, buffer, &officeFolderLength);
	const wchar_t* installError;
	switch(install)
	{
		// Ok
	case INSTALLSTATE_LOCAL:
	case INSTALLSTATE_SOURCE: return 0;
		// Error
	case INSTALLSTATE_NOTUSED: installError = L"INSTALLSTATE_NOTUSED"; break;
	case INSTALLSTATE_ABSENT: installError = L"INSTALLSTATE_ABSENT"; break;
	case INSTALLSTATE_INVALIDARG: installError = L"INSTALLSTATE_INVALIDARG"; break;
	case INSTALLSTATE_MOREDATA: installError = L"INSTALLSTATE_MORE_DATA"; break;
	case INSTALLSTATE_SOURCEABSENT: installError = L"INSTALLSTATE_SOURCEABSENT"; break;
	case INSTALLSTATE_UNKNOWN: installError = L"INSTALLSTATE_UNKNOWN"; break;
	default: installError = L"unexpected"; break;
	}
	if(installError != NULL)
	{
		std::wcout << installError << L" error detecting office component installation folder";
	}

	// Registry key
	static const wchar_t* const excelKey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe";
	HKEY hKey = NULL;
	LONG openError = ::RegOpenKeyExW(HKEY_LOCAL_MACHINE, excelKey, 0, KEY_READ, &hKey);
	if(openError != ERROR_SUCCESS)
	{
		std::wcout << L"error opening excel installation folder registry key";
	}
	else
	{
		DWORD bufferSize = MAX_PATH;
		ULONG queryError = ::RegQueryValueExW(hKey, NULL, NULL, NULL, (LPBYTE)buffer, &bufferSize);
		if(queryError != ERROR_SUCCESS)
		{
			std::wcout << L"error reading excel installation folder registry key";
		}
		else
		{
			size_t bufferLength = bufferSize / sizeof(buffer[0]) - 1;
#define EXCEL_KEY_VALUE_END L"EXCEL.EXE"
			const size_t excelKeyValueEndLength = sizeof(EXCEL_KEY_VALUE_END) / sizeof(EXCEL_KEY_VALUE_END[0]) - 1;
			if(bufferLength <= excelKeyValueEndLength)
			{
				std::wcout << L"excel installation folder registry key is too short";
			}
			else
			{
				size_t folderLength = bufferLength - excelKeyValueEndLength;
				if(_wcsicmp(EXCEL_KEY_VALUE_END, buffer + folderLength) != 0)
				{
					std::wcout << L"excel installation folder registry key should end with '" << EXCEL_KEY_VALUE_END << "' : " << buffer;
				}
				else
				{
					buffer[folderLength] = 0;
					wcscat_s(buffer, MAX_PATH, L"DCF");
					if(testSpreadsheetCompareFileInfolder(buffer) == false)
					{
						std::wcout << L"'spreadsheetcompare.exe' not found in excel installation folder '" << buffer << "'";
					}
					else
					{
						// Success at getting folder from registry key
						return 0;
					}
				}
			}
		}
	}

	// Default installation folders
	static const wchar_t* const defaultFolders[] =
	{
		L"C:\\Program Files\\Microsoft Office\\Office15\\DCF",
		L"C:\\Program Files\\Microsoft Office\\Office14\\DCF",
		L"C:\\Program Files\\Microsoft Office\\Office13\\DCF",
		L"C:\\Program Files\\Microsoft Office\\Office12\\DCF",
		L"C:\\Program Files\\Microsoft Office\\Office11\\DCF",
		L"C:\\Program Files\\Microsoft Office\\Office10\\DCF",

		L"C:\\Program Files (x86)\\Microsoft Office\\Office15\\DCF",
		L"C:\\Program Files (x86)\\Microsoft Office\\Office14\\DCF",
		L"C:\\Program Files (x86)\\Microsoft Office\\Office13\\DCF",
		L"C:\\Program Files (x86)\\Microsoft Office\\Office12\\DCF",
		L"C:\\Program Files (x86)\\Microsoft Office\\Office11\\DCF",
		L"C:\\Program Files (x86)\\Microsoft Office\\Office10\\DCF",
	};
	for(int i = 0; i < sizeof(defaultFolders) / sizeof(defaultFolders[0]); ++i)
	{
		const wchar_t* defaultFolder = defaultFolders[i];
		if(folderExists(defaultFolder) && testSpreadsheetCompareFileInfolder(buffer))
		{
			// Default installation directory
			wcscpy_s(buffer, MAX_PATH, defaultFolder);
			return 0;
		}
	}

	std::wcout << L"no default installation folder detected";
	return -1;
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
int writeTmpFile(const wchar_t* file1, const wchar_t* file2, wchar_t* tmpFile)
{
	wchar_t tmpPath [MAX_PATH];
	DWORD dwPathLength = GetTempPathW(MAX_PATH, tmpPath);
	if((dwPathLength == 0) || (dwPathLength >= MAX_PATH))
	{
		std::wcout << L"error getting temp path";
		return -1;
	}

	UINT fileUid = GetTempFileNameW(tmpPath, L"sscl", 0, tmpFile);
	if(fileUid == 0)
	{
		std::wcout << L"error getting temp file name";
		return -1;
	}

	std::wofstream oStream;
	oStream.open(tmpFile);
	oStream << file1;
	oStream << "\n";
	oStream << file2;
	oStream.close();
	return 0;
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
int launchCompare(const wchar_t* exeFolder, const wchar_t* tmpFile)
{
	wchar_t exeFile[MAX_PATH];
	wcscpy_s(exeFile, exeFolder);
	wcscat_s(exeFile, L"\\spreadsheetcompare.exe");
	HINSTANCE hShellExec = ::ShellExecuteW(NULL, L"open", exeFile, tmpFile, exeFolder, SW_SHOWNORMAL);
	bool success = (size_t(hShellExec) > 32);
	return (success ? 0 : -1);
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
int getUnquoted(const wchar_t* arg, wchar_t* buffer)
{
	const wchar_t* val = arg;

	bool quotes = (val[0] == L'"');
	if(quotes)
	{
		++val;
	}

	wcscpy_s(buffer, MAX_PATH, val);

	if(quotes)
	{
		size_t argLength = wcslen(val);
		if(argLength < 2)
		{
			std::wcout << L"invalid argument length";
			return -1;
		}
		else if(val[argLength - 1] != L'"')
		{
			std::wcout << L"quote detection error for argument " << arg;
			return -1;
		}
		else
		{
			buffer[argLength - 1] = 0;
		}
	}

	return 0;
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
int wmain(int argc, wchar_t *argv[], wchar_t *envp[])
{
	// Parse arguments
	wchar_t exeFolder[MAX_PATH] = { 0 };
	wchar_t file1[MAX_PATH] = { 0 };
	wchar_t file2[MAX_PATH] = { 0 };

	for(int i = 1; i < argc; ++i)
	{
		const wchar_t * arg = argv[i];
		if(arg != NULL)
		{
			if(_wcsnicmp(arg, L"-d=", argDirectoryLength) == 0)
			{
				if(exeFolder[0] != 0)
				{
					std::wcout << L"multiple exe folder arguments";
					return -1;
				}
				arg += argDirectoryLength;
				int unquoteCode = getUnquoted(arg, exeFolder);
				if(unquoteCode != 0)
				{
					std::wcout << L"error getting unquoted exe folder";
					return unquoteCode;
				}
			}
			else
			{
				wchar_t* fileBuffer = NULL;
				if(file1[0] == 0)
				{
					fileBuffer = file1;
				}
				else if(file2[0] == 0)
				{
					fileBuffer = file2;
				}
				else
				{
					std::wcout << L"unexpected argument " << arg;
					return -1;
				}

				int unquoteCode = getUnquoted(arg, fileBuffer);
				if(unquoteCode != 0)
				{
					std::wcout << L"error getting unquoted argument " << arg;
					return unquoteCode;
				}
			}
		}
	}
	if(exeFolder[0] == 0)
	{
		wchar_t officeFolder[MAX_PATH] = { 0 };
		int officeCode = getSpreadsheetCompareFolder(officeFolder);
		if(officeCode != 0)
		{
			std::wcout << L"error getting office folder";
			return -1;
		}
	}
	if(file1[0] == 0)
	{
		std::wcout << L"no compare file detected";
		return -1;
	}
	else if(file2[0] == 0)
	{
		std::wcout << L"only one compare file detected";
		return -1;
	}

	// Write tmp file
	wchar_t tmpFile[MAX_PATH] = { 0 };
	int writeCode = writeTmpFile(argv[1], argv[2], tmpFile);
	if(writeCode != 0)
	{
		std::wcout << L"error writing tmp file";
		return writeCode;
	}

	// Launch comparison
	 int launchCode = launchCompare(exeFolder, tmpFile);
	 if(launchCode != 0)
	 {
		 std::wcout << L"error writing tmp file";
		 return launchCode;
	 }

	 std::wcout << L"Launch success";
	 return 0;
}

