#include <windows.h>
#include <iostream>
#include <fstream>

int write_tmp_file(const wchar_t* file1, const wchar_t* file2)
{
	wchar_t tmpPath [MAX_PATH];
	DWORD dwPathLength = GetTempPathW(MAX_PATH, tmpPath);
	if((dwPathLength == 0) || (dwPathLength >= MAX_PATH))
	{
		std::cout << "error getting temp path";
		return -1;
	}

	wchar_t tmpFileName [MAX_PATH];
	UINT fileUid = GetTempFileNameW(tmpPath, L"sprecomplaun", 0, tmpFileName);
	if(fileUid == 0)
	{
		std::cout << "error getting temp file name";
		return -1;
	}

	std::ofstream oStream;
	oStream.open(tmpFileName);
	oStream << file1;
	oStream << "\n";
	oStream << file2;
	oStream.close();
	return 0;
}

int wmain(int argc, wchar_t *argv[], wchar_t *envp[])
{
	if(argc < 3)
	{
		std::cout << "expecting three arguments";
		return -1;
	}

	int writeCode = write_tmp_file(argv[1], argv[2]);
	if(writeCode != 0)
	{
		std::cout << "error writing tmp file";
		return writeCode;
	}

	return 0;
}

