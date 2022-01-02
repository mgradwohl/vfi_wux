// Visual File Information
// Copyright (c) Microsoft Corporation
// All rights reserved. 
// 
// MIT License
// 
// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the ""Software""), 
// to deal in the Software without restriction, including without limitation 
// the rights to use, copy, modify, merge, publish, distribute, sublicense, 
// and/or sell copies of the Software, and to permit persons to whom 
// the Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included 
// in all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS 
// OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR 
// IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

#include "pch.h"
#include <winnt.h>
#include <wil/cppwinrt.h>
#include <wil/common.h>
#include <Imagehlp.h>
#include <format>
#include "resource.h"
#include "helpers.h"
#include "WiseFile.h"
#pragma comment (lib, "imagehlp.lib")

#include <winver.h>
#pragma comment (lib, "version.lib")

CWiseFile::CWiseFile()
{
	m_strFullPath.clear();
	m_strPath.clear();
	m_strName.clear();
	m_strExt.clear();

	m_stLocalCreation = {};
	m_stLocalLastAccess = {};
	m_stLocalLastWrite = {};

	m_stUTCCreation = {};
	m_stUTCLastAccess = {};
	m_stUTCLastWrite = {};

	m_dwAttribs = 0;
	m_qwSize = 0;
	m_qwFileVersion = 0;
	m_qwProductVersion = 0;
	m_wLanguage = 0;
	m_CodePage = 0;
	m_dwOS = 0;
	m_dwType = 0;
	m_dwFlags = 0;
	m_dwCRC = 0;
	m_State = FileState::Valid;
	m_fDebugStripped = false;

}

CWiseFile::CWiseFile(std::wstring strFileSpec) : CWiseFile()
{
	Attach(strFileSpec);
	WI_SetFlag(m_State, FileState::CRC_Pending);
}

CWiseFile::~CWiseFile()
{
	m_State = FileState::Invalid;
}

bool CWiseFile::Attach(std::wstring strFileSpec)
{
	if (strFileSpec.starts_with(L"\\\\?\\"))
	{
		return false;
	}

	WIN32_FIND_DATA fd = {};
	HANDLE hff = FindFirstFile(strFileSpec.c_str(), &fd);
	if (INVALID_HANDLE_VALUE == hff)
	{
		//FindClose(hff);
		return false;
	}

	// this method is only for adding a single file, not a folder
	m_dwAttribs = fd.dwFileAttributes;
	if (m_dwAttribs & FILE_ATTRIBUTE_DIRECTORY)
	{
		FindClose(hff);
		return false;
	}

	m_qwSize = GetFileSize64(strFileSpec.c_str());
	WI_SetFlag(m_State, FileState::Size);

	FILETIME ft;
	::FileTimeToSystemTime(&(fd.ftCreationTime), &m_stUTCCreation);
	::FileTimeToLocalFileTime(&(fd.ftCreationTime), &ft);
	::FileTimeToSystemTime(&ft, &m_stLocalCreation);

	::FileTimeToSystemTime(&(fd.ftLastAccessTime), &m_stUTCLastAccess);
	::FileTimeToLocalFileTime(&(fd.ftLastAccessTime), &ft);
	::FileTimeToSystemTime(&ft, &m_stLocalLastAccess);

	::FileTimeToSystemTime(&(fd.ftLastWriteTime), &m_stUTCLastWrite);
	::FileTimeToLocalFileTime(&(fd.ftLastWriteTime), &ft);
	::FileTimeToSystemTime(&ft, &m_stLocalLastWrite);
	WI_SetFlag(m_State, FileState::Date);

	wchar_t szDrive[_MAX_DRIVE];
	wchar_t szDir[_MAX_DIR];
	wchar_t szName[_MAX_PATH];
	wchar_t szExt[_MAX_PATH];

	LPWSTR pszDot;
	_wsplitpath_s(strFileSpec.c_str(), szDrive, _MAX_DRIVE, szDir, _MAX_DIR, szName, _MAX_PATH, szExt, _MAX_PATH);

	m_strFullPath = strFileSpec;
	m_strPath = szDrive;
	m_strPath = szDir;
	m_strName = szName;

	pszDot = CharNext(szExt);
	m_strExt.assign(pszDot);

	SetDateTimeStrings();
	SetAttribs();

	WI_SetFlag(m_State, FileState::DateString);
	WI_SetFlag(m_State, FileState::Attached);

	return true;
}

bool CWiseFile::Detach()
{
	m_State = FileState::Invalid;
	return true;
}

bool CWiseFile::SetSize(bool bHex)
{
	if (!WI_IsFlagSet(m_State, FileState::Size))
	{
		return false;
	}

	if (bHex)
	{
		m_strCRC = std::format(L"{:#08x}", m_qwSize);
		WI_SetFlag(m_State, FileState::SizeString);
		return true;
	}
	else
	{
		// Not using StrFormatByteSize because it wordifies everything
		if (int2str(m_strSize, m_qwSize))
		{
			WI_SetFlag(m_State, FileState::SizeString);
			return true;
		}
		else
		{
			return false;
		}
	}
}

bool CWiseFile::SetCRC(bool bHex)
{
	if (WI_IsFlagSet(m_State, FileState::CRC_Complete))
	{
		if (bHex)
		{
			m_strCRC = std::format(L"{:#08x}", m_dwCRC);
		}
		else
		{
			int2str(m_strCRC, m_dwCRC);
		}
		WI_SetFlag(m_State, FileState::CRC_String);
		return true;
	}

	if (WI_IsFlagSet(m_State, FileState::CRC_Pending))
	{
		LoadStringResource(STR_CRC_PENDING, m_strCRC);
		return true;
	}

	if (WI_IsFlagSet(m_State, FileState::CRC_Working))
	{
		LoadStringResource(STR_CRC_WORKING, m_strCRC);
		return true;
	}

	if (WI_IsFlagSet(m_State, FileState::CRC_Error))
	{
		LoadStringResource(STR_CRC_ERROR, m_strCRC);
		return true;
	}
	return false;
}

bool CWiseFile::SetVersionStrings()
{
	if (!WI_IsFlagSet(m_State, FileState::VersionValid))
	{
		return false;
	}

	wchar_t szDec[10];
	GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, szDec, 2);

	// %d%s%02d%s%04d%s%04d
	m_strFileVersion = std::format(L"{}{}{:02d}{}{:04d}{}{:04d}",
		(int)HIWORD(HIDWORD(m_qwFileVersion)),
		szDec,
		(int)LOWORD(HIDWORD(m_qwFileVersion)),
		szDec,
		(int)HIWORD(LODWORD(m_qwFileVersion)),
		szDec,
		(int)LOWORD(LODWORD(m_qwFileVersion)));
	
	m_strProductVersion = std::format(L"{}{}{:02d}{}{:04d}{}{:04d}",
		(int)HIWORD(HIDWORD(m_qwProductVersion)),
		szDec,
		(int)LOWORD(HIDWORD(m_qwProductVersion)),
		szDec,
		(int)HIWORD(LODWORD(m_qwProductVersion)),
		szDec,
		(int)LOWORD(LODWORD(m_qwProductVersion)));

		switch (m_dwType)
		{
			case VFT_UNKNOWN: m_strType = std::format(L"VFT_UNKNOWN: {:#010x}", m_dwType);
				break;
			case VFT_APP: m_strType = L"VFT_APP";
				break;
			case VFT_DLL: m_strType = L"VFT_DLL";
				break;
			case VFT_DRV: m_strType = L"VFT_DRV";
				break;
			case VFT_FONT: m_strType = L"VFT_FONT";
				break;
			case VFT_VXD: m_strType = L"VFT_VXD";
				break;
			case VFT_STATIC_LIB: m_strType = L"VFT_STATIC_LIB";
				break;
			default: m_strType = std::format(L"Reserved: {:#010x}", m_dwType);
		}

		switch (m_dwOS)
		{
			case VOS_UNKNOWN: m_strOS = std::format(L"VOS_UNKNOWN: {:#010x}", m_dwOS);
				break;
			case VOS_DOS: m_strOS = L"VOS_DOS";
				break;
			case VOS_OS216: m_strOS = L"VOS_OS216";
				break;
			case VOS_OS232: m_strOS = L"VOS_OS232";
				break;
			case VOS_NT: m_strOS = L"VOS_NT";
				break;
			case VOS__WINDOWS16: m_strOS = L"VOS__WINDOWS16";
				break;
			case VOS__PM16: m_strOS = L"VOS__PM16";
				break;
			case VOS__PM32: m_strOS = L"VOS__PM32";
				break;
			case VOS__WINDOWS32: m_strOS = L"VOS__WINDOWS32";
				break;
			case VOS_OS216_PM16: m_strOS = L"VOS_OS216_PM16";
				break;
			case VOS_OS232_PM32: m_strOS = L"VOS_OS232_PM32";
				break;
			case VOS_DOS_WINDOWS16: m_strOS = L"VOS_DOS_WINDOWS16";
				break;
			case VOS_DOS_WINDOWS32: m_strOS = L"VOS_DOS_WINDOWS32";
				break;
			case VOS_NT_WINDOWS32: m_strOS = L"VOS_NT_WINDOWS32";
				break;
			default:		m_strOS = std::format(L"Reserved: {:#010x}", m_dwOS);
		}

		// list seperator
		wchar_t szSep[10];
		GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLIST, szSep, 2);

		std::wstring strFlags;
		if (m_dwFlags & VS_FF_DEBUG)
		{
			if (m_fDebugStripped)
			{

				LoadStringResource(STR_FLAG_DEBUG_STRIPPED, strFlags);
			}
			else
			{
				LoadStringResource(STR_FLAG_DEBUG, strFlags);
			}
			m_strFlags += strFlags;
			m_strFlags += szSep;
		}
		if (m_dwFlags & VS_FF_PRERELEASE)
		{
			LoadStringResource(STR_FLAG_PRERELEASE, strFlags);
			m_strFlags += strFlags;
			m_strFlags += szSep;
		}
		if (m_dwFlags & VS_FF_PATCHED)
		{
			LoadStringResource(STR_FLAG_PATCHED, strFlags);
			m_strFlags += strFlags;
			m_strFlags += szSep;
		}
		if (m_dwFlags & VS_FF_PRIVATEBUILD)
		{
			LoadStringResource(STR_FLAG_PRIVATEBUILD, strFlags);
			m_strFlags += strFlags;
			m_strFlags += szSep;
		}
		if (m_dwFlags & VS_FF_INFOINFERRED)
		{
			LoadStringResource(STR_FLAG_INFOINFERRED, strFlags);
			m_strFlags += strFlags;
			m_strFlags += szSep;
		}
		if (m_dwFlags & VS_FF_SPECIALBUILD)
		{
			LoadStringResource(STR_FLAG_SPECIALBUILD, strFlags);
			m_strFlags += strFlags;
			m_strFlags += szSep;
		}

		if (!m_strFlags.empty())
		{
			const size_t sep = m_strFlags.find_last_of(szSep[0]);
			m_strFlags.resize(sep);
		}
		// if we've looked at it, and we still don't have a string for it
		// put a default one in (unless the flags are 00000000)
		if ((m_strFlags.length() < 3) && (m_dwFlags != 0))
		{
			m_strFlags = std::format(L"{:010x}", m_dwFlags);
		}

	WI_SetFlag(m_State, FileState::VersionString);
	return true;
}

bool CWiseFile::GetFileVersion(LPDWORD pdwMS, LPDWORD pdwLS)
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (pdwMS)
		*pdwMS = HIDWORD(m_qwFileVersion);

	if (pdwLS)
		*pdwLS = LODWORD(m_qwFileVersion);

	return true;
}

bool CWiseFile::GetFileVersion(LPWORD pwHighMS, LPWORD pwLowMS, LPWORD pwHighLS, LPWORD pwLowLS)
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (pwHighMS)
		*pwHighMS = HIWORD(HIDWORD(m_qwFileVersion));

	if (pwLowMS)
		*pwLowMS = LOWORD(HIDWORD(m_qwFileVersion));

	if (pwHighLS)
		*pwHighLS = HIWORD(LODWORD(m_qwFileVersion));

	if (pwLowLS)
		*pwLowLS = LOWORD(LODWORD(m_qwFileVersion));

	return true;
}

bool CWiseFile::GetProductVersion(LPDWORD pdwMS, LPDWORD pdwLS)
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (pdwMS)
		*pdwMS = HIDWORD(m_qwProductVersion);

	if (pdwLS)
		*pdwLS = LODWORD(m_qwProductVersion);

	return true;
}

bool CWiseFile::GetProductVersion(LPWORD pwHighMS, LPWORD pwLowMS, LPWORD pwHighLS, LPWORD pwLowLS)
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (pwHighMS)
		*pwHighMS = HIWORD(HIDWORD(m_qwProductVersion));

	if (pwLowMS)
		*pwLowMS = LOWORD(HIDWORD(m_qwProductVersion));

	if (pwHighLS)
		*pwHighLS = HIWORD(LODWORD(m_qwProductVersion));

	if (pwLowLS)
		*pwLowLS = LOWORD(LODWORD(m_qwProductVersion));

	return true;
}

bool CWiseFile::ReadVersionInfo()
{
	if (WI_IsFlagSet(m_State, FileState::Version))
	{
		if (WI_IsFlagSet(m_State, FileState::VersionValid))
		{
			std::cout << "ReadVersionInfo: Version information already read." << std::endl;
		}
		else
		{
			std::cout << "ReadVersionInfo: Version information already read and FAILED" << std::endl;
		}
		return true;
	}

	char szPath[_MAX_PATH];
	WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, m_strFullPath.c_str(), _MAX_PATH, szPath, _MAX_PATH, nullptr, nullptr);
	PLOADED_IMAGE pli;
	pli = ::ImageLoad(szPath, NULL);
	if (NULL != pli)
	{
		if (pli->Characteristics & IMAGE_FILE_DEBUG_STRIPPED)
		{
			m_fDebugStripped = true;
		}
		else
		{
			m_fDebugStripped = false;
		}
		::ImageUnload(pli);
	}

	DWORD	dwVerSize = 0;
	LPVOID	lpVerBuffer = NULL;
	LPVOID	lpVerData = NULL;
	UINT	cbVerData = 0;

	// this takes a long time to call, max size I've seen is 5476
	DWORD	dwVerHandle = 0;
	dwVerSize = ::GetFileVersionInfoSize(m_strFullPath.c_str(), &dwVerHandle);

	if (dwVerSize < 1)
		return false;

	lpVerBuffer = (LPVOID) new BYTE[dwVerSize];

	if (NULL == lpVerBuffer)
	{
		delete[] lpVerBuffer;
		return false;
	}

	ZeroMemory(lpVerBuffer, dwVerSize);
#pragma warning(suppress: 6388)
// the warning is that dwVerHandle can not be zero. We never call this with dwVerHandle == 0
// because of the dwVerSize < 1 check above, that returns an error if there is no version information
	if (!::GetFileVersionInfo(m_strFullPath.c_str(), dwVerHandle, dwVerSize, lpVerBuffer))
	{
		delete[] lpVerBuffer;
		return false;
	}

	if (!::VerQueryValue(lpVerBuffer, L"\\", &lpVerData, &cbVerData))
	{
		delete[] lpVerBuffer;
		return false;
	}

#define pVerFixedInfo ((VS_FIXEDFILEINFO FAR*)lpVerData)
	m_qwFileVersion = MAKEQWORD(pVerFixedInfo->dwFileVersionLS, pVerFixedInfo->dwFileVersionMS);
	m_qwProductVersion = MAKEQWORD(pVerFixedInfo->dwProductVersionLS, pVerFixedInfo->dwProductVersionMS);
	m_dwType = pVerFixedInfo->dwFileType;
	m_dwOS = pVerFixedInfo->dwFileOS;
	m_dwFlags = pVerFixedInfo->dwFileFlags;
#undef pVerFixedInfo

	if (!VerQueryValue(lpVerBuffer, L"\\VarFileInfo\\Translation", &lpVerData, &cbVerData))
	{
		delete[] lpVerBuffer;
		return false;
	}

	m_wLanguage = LOWORD(*(LPDWORD)lpVerData);
	m_CodePage = HIWORD(*(LPDWORD)lpVerData);
	delete[] lpVerBuffer;

	WI_SetFlag(m_State, FileState::Version);
	WI_SetFlag(m_State, FileState::VersionValid);

	return true;
}

bool CWiseFile::SetDateTimeStrings(bool fLocal)
{
	// only called from Attach()
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	if (WI_IsFlagSet(m_State, FileState::DateString))
	{
		// strings are set, bail
		return true;
	}
	const int cDate = 200;
	wchar_t szDate[cDate];
	std::wstring strFail = L"? fail";
	SYSTEMTIME* pst = nullptr;

	// Created DATE
	if (fLocal)
	{
		pst = &m_stLocalCreation;
	}
	else
	{
		pst = &m_stUTCCreation;
	}

	if (0 == ::GetDateFormat(::GetThreadLocale(), DATE_SHORTDATE, pst, NULL, szDate, cDate))
	{
		m_strDateCreated = strFail;
	}
	else
	{
		m_strDateCreated = szDate;
	}
	if (0 == ::GetTimeFormat(::GetThreadLocale(), TIME_NOSECONDS, pst, NULL, szDate, cDate))
	{
		m_strTimeCreated = strFail;
	}
	else
	{
		m_strTimeCreated = szDate;
	}

	// Last accessed DATE
	if (fLocal)
	{
		pst = &m_stLocalLastAccess;
	}
	else
	{
		pst = &m_stUTCLastAccess;
	}

	if (0 == ::GetDateFormat(::GetThreadLocale(), DATE_SHORTDATE, pst, NULL, szDate, cDate))
	{
		m_strDateLastAccess = strFail;
	}
	else
	{
		m_strDateLastAccess = szDate;
	}
	if (0 == ::GetTimeFormat(::GetThreadLocale(), TIME_NOSECONDS, pst, NULL, szDate, cDate))
	{
		m_strTimeLastAccess = strFail;
	}
	else
	{
		m_strTimeLastAccess = szDate;
	}

	// Last write DATE
	if (fLocal)
	{
		pst = &m_stLocalLastWrite;
	}
	else
	{
		pst = &m_stUTCLastWrite;
	}

	if (0 == ::GetDateFormat(::GetThreadLocale(), DATE_SHORTDATE, pst, NULL, szDate, cDate))
	{
		m_strDateLastWrite = strFail;
	}
	else
	{
		m_strDateLastWrite = szDate;
	}
	if (0 == ::GetTimeFormat(::GetThreadLocale(), TIME_NOSECONDS, pst, NULL, szDate, cDate))
	{
		m_strTimeLastWrite = strFail;

	}
	else
	{
		m_strTimeLastWrite = szDate;
	}
	return true;
}

bool CWiseFile::TouchFileTime(FILETIME* lpTime)
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	if (!lpTime)
		return false;

	LPCWSTR pszPath = m_strFullPath.c_str();
	const DWORD dwOldAttribs = ::GetFileAttributes(pszPath);
	if (false == ::SetFileAttributes(pszPath, FILE_ATTRIBUTE_NORMAL))
	{
		::SetFileAttributes(pszPath, dwOldAttribs);
		return false;
	}

	HANDLE hFile = ::CreateFile(pszPath,
		GENERIC_WRITE,
		FILE_SHARE_READ | FILE_SHARE_WRITE,
		NULL,
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL | FILE_FLAG_NO_BUFFERING,
		NULL);

	if (INVALID_HANDLE_VALUE == hFile)
	{
		::SetFileAttributes(pszPath, dwOldAttribs);
		::CloseHandle(hFile);
		return false;
	}

	if (false == ::SetFileTime(hFile, lpTime, lpTime, lpTime))
	{
		::SetFileAttributes(pszPath, dwOldAttribs);
		::CloseHandle(hFile);
		return false;
	}

	::SetFileAttributes(pszPath, dwOldAttribs);
	::CloseHandle(hFile);

	FILETIME ft;
	::FileTimeToSystemTime(lpTime, &m_stUTCCreation);
	::FileTimeToLocalFileTime(lpTime, &ft);
	::FileTimeToSystemTime(&ft, &m_stLocalCreation);

	::FileTimeToSystemTime(lpTime, &m_stUTCLastAccess);
	::FileTimeToLocalFileTime(lpTime, &ft);
	::FileTimeToSystemTime(&ft, &m_stLocalLastAccess);

	::FileTimeToSystemTime(lpTime, &m_stUTCLastWrite);
	::FileTimeToLocalFileTime(lpTime, &ft);
	::FileTimeToSystemTime(&ft, &m_stLocalLastWrite);

	return true;
}

bool CWiseFile::SetAttribs()
{
	// called from Attach only. If the file is attached, the strings are valid
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		m_strAttribs = L"        ";
		return false;
	}

	if (m_dwAttribs & FILE_ATTRIBUTE_NORMAL)
	{
		m_strAttribs = L"        ";
		return true;
	}

	m_strAttribs = std::format(L"{}{}{}{}{}{}{}{}",
		((m_dwAttribs & FILE_ATTRIBUTE_ARCHIVE) ? 'A' : ' '),
		((m_dwAttribs & FILE_ATTRIBUTE_HIDDEN) ? 'H' : ' '),
		((m_dwAttribs & FILE_ATTRIBUTE_READONLY) ? 'R' : ' '),
		((m_dwAttribs & FILE_ATTRIBUTE_SYSTEM) ? 'S' : ' '),
		((m_dwAttribs & FILE_ATTRIBUTE_ENCRYPTED) ? 'E' : ' '),
		((m_dwAttribs & FILE_ATTRIBUTE_COMPRESSED) ? 'C' : ' '),
		((m_dwAttribs & FILE_ATTRIBUTE_TEMPORARY) ? 'T' : ' '),
		((m_dwAttribs & FILE_ATTRIBUTE_OFFLINE) ? 'O' : ' '));

	return true;
}

bool CWiseFile::SetLanguage(UINT Language, bool bNumeric)
{
	if (WI_IsFlagSet(m_State, FileState::LanguageString))
	{
		// string is already set, bail
		return true;
	}

	// Set the Language first
	wchar_t szTemp[255];
	if (!Language)
	{
		VerLanguageName(Language, szTemp, 255);
	}
	else
	if (!GetLocaleInfoW(Language, LOCALE_SLANGUAGE, szTemp, 255))
	{
		VerLanguageName(Language, szTemp, 255);
	}
	m_strLanguage = std::format(L"%lu %s", Language, szTemp);

	// Set the CodePage
	if (bNumeric)
	{
		m_strCodePage = std::format(L"{:#06x}", (WORD)m_CodePage);
	}
	else
	{
		m_CodePage.GetCodePageName(LANGIDFROMLCID(GetThreadLocale()), m_CodePage, m_strCodePage);
	}

	WI_SetFlag(m_State, FileState::Language);
	WI_SetFlag(m_State, FileState::LanguageString);
	return true;
}

const std::wstring& CWiseFile::GetFieldString(int iField, bool fOptions)
{
	fOptions;
	switch (iField)
	{
	case  0: return m_strFullPath;
	case  1: return m_strName;
	case  2: return m_strExt;
	case  3: return GetSize();
	case  4: return GetDateCreated();
	case  5: return GetTimeCreated();
	case  6: return GetDateLastWrite();
	case  7: return GetTimeLastWrite();
	case  8: return GetDateLastAccess();
	case  9: return GetTimeLastAccess();
	case 10: return GetAttribs();
	case 11: return GetFileVersion();
	case 12: return GetProductVersion();
	case 13: return GetLanguage();
	case 14: return GetCodePage();
	case 15: return GetOS();
	case 16: return GetType();
	case 17: return GetFlags();
	case 18: return GetCRC();
	}
	return L"Unknown";
}

bool CWiseFile::ReadVersionInfoEx()
{
	HINSTANCE h = ::LoadLibraryEx(m_strFullPath.c_str(), NULL, LOAD_LIBRARY_AS_DATAFILE);
	if (NULL == h)
	{
		return false;
	}

	HRSRC hrsrc = ::FindResource(h, MAKEINTRESOURCE(1), RT_VERSION);
	if (NULL == hrsrc)
	{
		FreeLibrary(h);
		return false;
	}

	UINT cb = ::SizeofResource(h, hrsrc);
	if (0 == cb)
	{
		FreeLibrary(h);
		return false;
	}

	const UINT xcb = ::GetFileVersionInfoSize(m_strFullPath.c_str(), 0);

	HGLOBAL hglobal = ::LoadResource(h, hrsrc);
	if (hglobal == NULL)
	{
		FreeLibrary(h);
		return false;
	}

	LPBYTE pb = static_cast<BYTE*>(::LockResource(hglobal));
	if (NULL == pb)
	{
		FreeLibrary(h);
		return false;
	}

	VS_FIXEDFILEINFO* pvi = new VS_FIXEDFILEINFO;
	LPBYTE pVerBuffer = pb;

	// Usually before you get here, you call GetFileVersionInfoSize to figure out cb
	// to get that size manually, you have to walk thru umpteen fields
	// then you call get file version info, which supposedly gives you VS_VERSION_INFO
	// but pb points to the VS_VERSION_INFO, so we'll VerQueryValue directly

	// next line causes access violation
	if (!::VerQueryValue(pb, L"\\", (void**)&pvi, &cb))
	{
		return false;
	}

	if (pvi->dwSignature != 0xFEEF04BD)
	{
		FreeLibrary(h);
		return false;
	}

	m_qwFileVersion = MAKEQWORD(pvi->dwFileVersionLS, pvi->dwFileVersionMS);
	m_qwProductVersion = MAKEQWORD(pvi->dwProductVersionLS, pvi->dwProductVersionMS);
	m_dwType = pvi->dwFileType;
	m_dwOS = pvi->dwFileOS;
	m_dwFlags = pvi->dwFileFlags;

	LPVOID pVerData = NULL;
	UINT cbVerData = 0;
	if (!::VerQueryValue(pVerBuffer, L"\\VarFileInfo\\Translation", &pVerData, &cbVerData))
	{
		return false;
	}

	m_wLanguage = LOWORD(*(LPDWORD)pVerBuffer);
	m_CodePage = HIWORD(*(LPDWORD)pVerBuffer);

	FreeLibrary(h);
	return true;
}