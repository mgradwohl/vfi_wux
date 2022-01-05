#include "pch.h"

#include <shlwapi.h>
#pragma comment (lib, "shlwapi.lib")
#include <shlobj_core.h>

#include "helpers.h"

bool LoadStringResource(UINT id, std::wstring& strDest)
{
	LPWSTR pszBuffer = nullptr;
	strDest.clear();
	const size_t length = ::LoadString(GetModuleHandle(nullptr), id, (LPWSTR)&pszBuffer, 0);
	if (length == 0)
	{
		return false;
	}

	strDest.assign(pszBuffer, length);
	return true;
}

// Get number as a string
bool int2str(std::wstring& strDest, QWORD i)
{
	strDest = std::format(L"{:L}", i);

	//// get the separator
	//std::wstring strDec;
	//GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, &strDec[0], 2);

	//wchar_t pBuffer[200];
	//std::wstring strBuffer = strDest;
	//// TODO where did these numbers come from
	//strBuffer.resize(68);
	//_ui64tow_s(i, &strBuffer[0], 67, 10);

	//int length = GetNumberFormatW(LOCALE_USER_DEFAULT, 0, strBuffer.c_str(), nullptr, pBuffer, 65);
	//if (length == 0)
	//	return false;

	//strDest.assign(pBuffer, length);
	//const size_t decimal = strDest.find_last_of(strDec[0]);
	//strDest.resize(decimal);

	return true;
}

bool pipe2null(std::wstring& strSource)
{
	if (strSource.length() < 1)
	{
		return false;
	}

	replace(strSource.begin(), strSource.end(), 'L', '\0');
	return true;
}

bool MyGetUserName(std::wstring& strUserName)
{
	DWORD dwLen = MAX_USERNAME;

	wchar_t pBuffer[MAX_USERNAME];

	if (!GetUserName(pBuffer, &dwLen))
	{
		return false;
	}

	// todo use shrink_to_fit
	strUserName.assign(pBuffer, dwLen);
//	strUserName.resize(wcslen(strUserName.c_str()));
	return true;
}

//TODO all wstring
uint64_t GetFileSize64(const std::wstring& pszFileSpec)
{
	HANDLE hFile = CreateFile(pszFileSpec.c_str(),
		GENERIC_READ,
		FILE_SHARE_READ,
		NULL,
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL, NULL);

	if (INVALID_HANDLE_VALUE == hFile)
	{
		CloseHandle(hFile);
		return 0;
	}

	LARGE_INTEGER qwSize;
	if (!GetFileSizeEx(hFile, &qwSize))
	{
		CloseHandle(hFile);
		return 0;
	}
	CloseHandle(hFile);
	return qwSize.QuadPart;
}

bool GetModuleFolder(HINSTANCE hInst, std::wstring& pszFolder)
{
	wchar_t pBuffer[_MAX_PATH];
	::GetModuleFileName(hInst, pBuffer, _MAX_PATH);
	::PathRemoveFileSpec(pBuffer);

	pszFolder.assign(pBuffer);

	return true;
}

bool DoesFileExist(const std::wstring& pszFileName)
{
	DWORD const dw = ::GetFileAttributes(pszFileName.c_str());
	if (dw == (DWORD)-1 || dw & FILE_ATTRIBUTE_DIRECTORY || dw & FILE_ATTRIBUTE_OFFLINE)
	{
		return false;
	}

	return true;
}

bool DoesFolderExist(const std::wstring& pszFolder)
{
	DWORD const dw = ::GetFileAttributes(pszFolder.c_str());
	if (dw == (DWORD)-1 || dw & FILE_ATTRIBUTE_OFFLINE)
	{
		return false;
	}

	if (dw & FILE_ATTRIBUTE_DIRECTORY)
	{
		return true;
	}

	return false;
}

// TODO all wstrings
bool OpenBox(const HWND hWnd, const std::wstring& pszTitle, const std::wstring& pszFilter, std::wstring& strFile, int cchFile, const std::wstring& pszFolder, const DWORD dwFlags)
{
	OPENFILENAME of;
	::ZeroMemory(&of, sizeof(OPENFILENAME));

	LPWSTR wszPath;
	if (pszFolder.empty())
	{
		SHGetKnownFolderPath(FOLDERID_Desktop, KF_FLAG_DEFAULT, NULL, &wszPath);
		std::wstring strPath(wszPath);
		of.lpstrInitialDir = strPath.c_str();
		CoTaskMemFree(wszPath);
	}
	else
	{
		of.lpstrInitialDir = pszFolder.c_str();
	}

	strFile.clear();
	strFile.resize(maxExtendedPathLength);
	of.lStructSize = sizeof(OPENFILENAME);
	of.hwndOwner = hWnd;
	of.hInstance = GetModuleHandle(NULL);
	of.lpstrFilter = pszFilter.c_str();
	of.lpstrFile = strFile.data();
	of.nMaxFile = cchFile;
	of.lpstrTitle = pszTitle.c_str();
	of.Flags = dwFlags | OFN_DONTADDTORECENT | OFN_LONGNAMES | OFN_ENABLESIZING | OFN_EXPLORER | OFN_NONETWORKBUTTON;

	if (!GetOpenFileNameW(&of))
	{
		return false;
	}

	return true;
}
