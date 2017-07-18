#include <windows.h>
#include <msi.h>
#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <shlwapi.h>
#include <assert.h>

#define SSCL_ARG_VERBOSE L"-v"
#define SSCL_ARG_PAUSE L"-p"
#define SSCL_ARG_FILE L"-f="
static const size_t SSCL_ARG_FILE_LENGTH = sizeof(SSCL_ARG_FILE) / sizeof(SSCL_ARG_FILE[0]) - 1;
#define SSCL_ARG_DIRECTORY L"-d="
static const size_t SSCL_ARG_DIRECTORY_LENGTH = sizeof(SSCL_ARG_DIRECTORY) / sizeof(SSCL_ARG_DIRECTORY[0]) - 1;

bool verbose = false;
#define SSCL_LOG(x_msg) if(verbose) { std::wcout << L"SSCL: " << x_msg << std::endl; }

#ifdef _WIN32
static const wchar_t fileSeparator = L'\\';
static const wchar_t fileWrongSeparator = L'/';
#else // _WIN
static const wchar_t fileSeparator = L'/';
static const wchar_t fileWrongSeparator = L'\\';
#endif // _WIN

typedef const wchar_t* sz;

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool ssclGetCurrentDirectory(wchar_t* directoryPath)
{
	DWORD get = GetCurrentDirectoryW(MAX_PATH, directoryPath);
	if(get != 0)
	{
		return true;
	}

	return false;
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool ssclGetExecutableDirectory(wchar_t* directoryPath)
{
	BOOL get = GetModuleFileNameW(NULL, directoryPath, MAX_PATH);
	if(get != FALSE)
	{
		wchar_t* found = wcsrchr(directoryPath, fileSeparator);
		if(found != NULL)
		{
			*found = 0;
		}
		return true;
	}

	return false;
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool ssclDirectoryExists(sz directoryPath)
{
	DWORD attribs = GetFileAttributesW(directoryPath);
	return ((attribs != INVALID_FILE_ATTRIBUTES) && (attribs & FILE_ATTRIBUTE_DIRECTORY) != 0);
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool ssclFileExists(sz filePath)
{
	DWORD attribs = GetFileAttributes(filePath);
	return ((attribs != INVALID_FILE_ATTRIBUTES) && ((attribs & FILE_ATTRIBUTE_DIRECTORY) == 0));
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool ssclFilePathIsAbsolute(sz fileName)
{
	return (wcschr(fileName, L':') != NULL);
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
void ssclCanonicalizeFilePath(sz in, wchar_t* out)
{
	wchar_t tmpBuff1[MAX_PATH];
	wcscpy_s(tmpBuff1, in);
	wchar_t* val = tmpBuff1;
	while(*val != 0)
	{
		if(*val == fileWrongSeparator)
		{
			*val = fileSeparator;
		}
		++val;
	}

	// PathCchCanonicalize
	// PathCchCanonicalizeEx
#ifndef NDEBUG
	BOOL canon =
#endif // NDEBUG
		PathCanonicalize(out, tmpBuff1);
	assert(canon == TRUE);
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool ssclTryFindAbsolute(sz localFilePath, const sz* absoluteDirs, size_t absoluteDirsCount, wchar_t* out)
{
	static const wchar_t separator [] = { fileSeparator, 0 };
	for(size_t i = 0; i < absoluteDirsCount; ++i)
	{
		sz absDir = absoluteDirs[i];
		if(absDir != NULL)
		{
			wchar_t absFilePath[MAX_PATH] = { 0 };
			wcscpy_s(absFilePath, absDir);
			wcscat_s(absFilePath, separator);
			wcscat_s(absFilePath, localFilePath);
			wchar_t canonAbsFilePath[MAX_PATH] = { 0 };
			ssclCanonicalizeFilePath(absFilePath, canonAbsFilePath);
			if(ssclFileExists(canonAbsFilePath))
			{
				wcscpy_s(out, MAX_PATH, canonAbsFilePath);
				return true;
			}
		}
	}
	return false;
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
bool ssclTestSpreadsheetCompareFileInDirectory(const wchar_t* folder)
{
	wchar_t testPath[MAX_PATH];
	wcscpy_s(testPath, folder);
	wcscat_s(testPath, L"\\SPREADSHEETCOMPARE.EXE");
	return ssclFileExists(testPath);
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
int ssclGetSpreadsheetCompareDirectory(wchar_t* buffer, const wchar_t* const officeUid)
{
	DWORD officeDirLength = MAX_PATH;
	INSTALLSTATE install = ::MsiLocateComponentW(officeUid, buffer, &officeDirLength);
	const wchar_t* installError;
	switch(install)
	{
		// Ok
	case INSTALLSTATE_LOCAL:
	case INSTALLSTATE_SOURCE: installError = NULL; break;
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
		SSCL_LOG(L" could not detect office component installation directory, error " << installError);
	}
	else
	{
		wcscat_s(buffer, MAX_PATH, L"\\DCF");

		if(ssclTestSpreadsheetCompareFileInDirectory(buffer) == false)
		{
			SSCL_LOG(L"'spreadsheetcompare.exe' not found in excel installation directory '" << buffer << "'");
		}
		else
		{
			// Success at getting installation directory
			return 0;
		}
	}

	// Registry key
	static const wchar_t* const excelKey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe";
	HKEY hKey = NULL;
	LONG openError = RegOpenKeyExW(HKEY_LOCAL_MACHINE, excelKey, 0, KEY_READ, &hKey);
	if(openError != ERROR_SUCCESS)
	{
		SSCL_LOG(L"error opening excel installation directory registry key");
	}
	else
	{
		DWORD bufferSize = MAX_PATH;
		ULONG queryError = RegQueryValueExW(hKey, NULL, NULL, NULL, (LPBYTE)buffer, &bufferSize);
		if(queryError != ERROR_SUCCESS)
		{
			SSCL_LOG(L"error reading excel installation directory registry key");
		}
		else
		{
			size_t bufferLength = bufferSize / sizeof(buffer[0]) - 1;
#define EXCEL_KEY_VALUE_END L"EXCEL.EXE"
			const size_t excelKeyValueEndLength = sizeof(EXCEL_KEY_VALUE_END) / sizeof(EXCEL_KEY_VALUE_END[0]) - 1;
			if(bufferLength <= excelKeyValueEndLength)
			{
				SSCL_LOG(L"excel installation directory registry key is too short");
			}
			else
			{
				size_t dirLength = bufferLength - excelKeyValueEndLength;
				if(_wcsicmp(EXCEL_KEY_VALUE_END, buffer + dirLength) != 0)
				{
					SSCL_LOG(L"excel installation folder registry key should end with '" << EXCEL_KEY_VALUE_END << "' : " << buffer);
				}
				else
				{
					buffer[dirLength] = 0;
					wcscat_s(buffer, MAX_PATH, L"DCF");
					if(ssclTestSpreadsheetCompareFileInDirectory(buffer) == false)
					{
						SSCL_LOG(L"'spreadsheetcompare.exe' not found in excel registry key installation folder '" << buffer << "'");
					}
					else
					{
						// Success at getting directory from registry key
						return 0;
					}
				}
			}
		}
	}

	// Default installation directories
	static const wchar_t* const defaultDirs[] =
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
	for(int i = 0; i < sizeof(defaultDirs) / sizeof(defaultDirs[0]); ++i)
	{
		const wchar_t* defaultDir = defaultDirs[i];
		if(ssclDirectoryExists(defaultDir) && ssclTestSpreadsheetCompareFileInDirectory(buffer))
		{
			// Default installation directory
			wcscpy_s(buffer, MAX_PATH, defaultDir);
			return 0;
		}
	}

	SSCL_LOG(L"no default installation directory detected");
	return -1;
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
int ssclWriteTmpFile(const wchar_t* file1, const wchar_t* file2, wchar_t* tmpFile)
{
	wchar_t tmpPath [MAX_PATH];
	DWORD dwPathLength = GetTempPathW(MAX_PATH, tmpPath);
	if((dwPathLength == 0) || (dwPathLength >= MAX_PATH))
	{
		SSCL_LOG(L"error getting temp path");
		return -1;
	}

	UINT fileUid = GetTempFileNameW(tmpPath, L"sscl", 0, tmpFile);
	if(fileUid == 0)
	{
		SSCL_LOG(L"error getting temp file name");
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
int ssclLaunchCompare(const wchar_t* exeDirectory, const wchar_t* tmpFile)
{
	wchar_t exeFile[MAX_PATH];
	wcscpy_s(exeFile, exeDirectory);
	wcscat_s(exeFile, L"\\spreadsheetcompare.exe");
	HINSTANCE hShellExec = ShellExecuteW(NULL, L"open", exeFile, tmpFile, exeDirectory, SW_SHOWNORMAL);
	bool success = (size_t(hShellExec) > 32);
	return (success ? 0 : -1);
}

//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
int ssclGetUnquoted(const wchar_t* arg, wchar_t* buffer)
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
			SSCL_LOG(L"invalid argument length");
			return -1;
		}
		else if(val[argLength - 1] != L'"')
		{
			SSCL_LOG(L"quote detection error for argument " << arg);
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
int wmain(int argc, wchar_t *argv[] /*, wchar_t *envp[]*/)
{
	wchar_t exeDir[MAX_PATH] = { 0 };
	if(ssclGetExecutableDirectory(exeDir) == false)
	{
		SSCL_LOG(L"error getting executable directory");
		return -1;
	}
	wchar_t curDir[MAX_PATH] = { 0 };
	if(ssclGetCurrentDirectory(curDir) == false)
	{
		SSCL_LOG(L"error getting current directory");
		return -1;
	}

	// Test for pause argument
	for(int i = 1; i < argc; ++i)
	{
		const wchar_t* arg = argv[i];
		if((arg != NULL) && (_wcsicmp(arg, SSCL_ARG_PAUSE) == 0))
		{
			system("pause");
			break;
		}
	}

	// Test for verbose argument
	for(int i = 1; i < argc; ++i)
	{
		const wchar_t* arg = argv[i];
		if((arg != NULL) && (_wcsicmp(arg, SSCL_ARG_VERBOSE) == 0))
		{
			verbose = true;
			break;
		}
	}

	// Parse arguments
	std::vector<std::wstring> argDirs;
	wchar_t sscDir[MAX_PATH] = { 0 };
	wchar_t file1[MAX_PATH] = { 0 };
	wchar_t file2[MAX_PATH] = { 0 };
	for(int i = 1; i < argc; ++i)
	{
		const wchar_t* arg = argv[i];
		if(arg != NULL)
		{
			if(_wcsicmp(arg, SSCL_ARG_PAUSE) == 0)
			{
				continue;
			}
			if(_wcsicmp(arg, SSCL_ARG_VERBOSE) == 0)
			{
				continue;
			}
			else if(_wcsnicmp(arg, SSCL_ARG_DIRECTORY, SSCL_ARG_DIRECTORY_LENGTH) == 0)
			{
				if(sscDir[0] != 0)
				{
					SSCL_LOG(L"multiple exe directory arguments");
					return -1;
				}
				arg += SSCL_ARG_DIRECTORY_LENGTH;
				int unquoteCode = ssclGetUnquoted(arg, sscDir);
				if(unquoteCode != 0)
				{
					SSCL_LOG(L"error getting unquoted exe directory from argument " << arg);
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
					SSCL_LOG(L"unexpected argument " << arg);
					return -1;
				}

				int unquoteCode = ssclGetUnquoted(arg, fileBuffer);
				if(unquoteCode != 0)
				{
					SSCL_LOG(L"error getting unquoted argument " << arg);
					return unquoteCode;
				}
			}
		}
	}

	// Check arguments
	if(sscDir[0] == 0)
	{
		// Locate component
		static const wchar_t* office = L"{00000409-78E1-11D2-B60F-006097C998E7}";
		static const wchar_t* office2013 = L"{91150000-0011-0000-0000-0000000FF1CE}";
		static const wchar_t* excel = L"{CC29E96F-7BC2-11D1-A921-00A0C91E2AA2}";
		if((ssclGetSpreadsheetCompareDirectory(sscDir, office) != 0) &&
		   (ssclGetSpreadsheetCompareDirectory(sscDir, office2013) != 0) &&
		   (ssclGetSpreadsheetCompareDirectory(sscDir, excel) != 0))
		{
			SSCL_LOG(L"error getting Spreadsheetcompare directory");
			return -1;
		}
	}
	if(sscDir[0] == 0)
	{
		SSCL_LOG(L"no Spreadsheetcompare directory detected");
		return -1;
	}
	else
	{
		SSCL_LOG(L"Spreadsheetcompare directory : " << sscDir);
	}

	sz absDirs[] = { sz(exeDir), sz(curDir), sz(sscDir) };
	size_t absDirsCount = sizeof(absDirs) / sizeof(absDirs[0]);

	if(file1[0] == 0)
	{
		SSCL_LOG(L"no compare file detected");
		return -1;
	}
	else
	{
		if(ssclFilePathIsAbsolute(file1) == false)
		{
			if(ssclTryFindAbsolute(file1, absDirs, absDirsCount, file1) == false)
			{
				SSCL_LOG(L"Could not find absolute file1 : " << file1);
				return -1;
			}
		}
		SSCL_LOG(L"File1 : " << file1);
	}
	if(file2[0] == 0)
	{
		SSCL_LOG(L"only one compare file detected");
		return -1;
	}
	else
	{
		if(ssclFilePathIsAbsolute(file2) == false)
		{
			if(ssclTryFindAbsolute(file2, absDirs, absDirsCount, file2) == false)
			{
				SSCL_LOG(L"Could not find absolute file2 : " << file2);
				return -1;
			}
		}
		SSCL_LOG(L"File2 : " << file2);
	}

	// Find file1

	// Write tmp file
	wchar_t tmpFile[MAX_PATH] = { 0 };
	int writeCode = ssclWriteTmpFile(file1, file2, tmpFile);
	if(writeCode != 0)
	{
		SSCL_LOG(L"error writing tmp file");
		return writeCode;
	}
	else
	{
		SSCL_LOG(L"Tmp file : " << tmpFile);
	}

	// Launch comparison
	int launchCode = ssclLaunchCompare(sscDir, tmpFile);
	if(launchCode != 0)
	{
		SSCL_LOG(L"error writing tmp file");
		return launchCode;
	}
	else
	{
		SSCL_LOG(L"Launch success");
	}

	return 0;
}
