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
	m_fHasVersion = false;
	m_fszFlags = false;
	m_fszAttribs = false;
	m_fszOS = false;
	m_fszType = false;

}

CWiseFile::CWiseFile(std::wstring strFileSpec) : CWiseFile()
{
	Attach(strFileSpec);
	WI_SetFlag(m_State, FileState::CRC_Working);
}

CWiseFile::~CWiseFile()
{
	m_State = FileState::Invalid;
}

bool CWiseFile::Attach(std::wstring strFileSpec)
{
	// TODO check to see if it's a weird path
	//if (0 == StrCmpN(pszFileSpec, L"\\\\?\\", 4))
	//{
	//	return FWF_ERR_SPECIALPATH;
	//}

	WIN32_FIND_DATA fd = {};

	HANDLE hff = FindFirstFile(strFileSpec.c_str(), &fd);
	if (INVALID_HANDLE_VALUE == hff)
	{
		//FindClose(hff);
		return false;
	}

	m_dwAttribs = fd.dwFileAttributes;
	if (m_dwAttribs & FILE_ATTRIBUTE_DIRECTORY)
	{
		FindClose(hff);
		return false;
	}

	m_qwSize = GetFileSize64(strFileSpec.c_str());

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

	wchar_t szDrive[_MAX_DRIVE];
	wchar_t szDir[_MAX_DIR];
	wchar_t szName[_MAX_PATH];
	wchar_t szExt[_MAX_PATH];

	LPWSTR pszDot;
	_wsplitpath_s(strFileSpec.c_str(), szDrive, _MAX_DRIVE, szDir, _MAX_DIR, szName, _MAX_PATH, szExt, _MAX_PATH);

	m_strFullPath = strFileSpec;
	m_strPath.assign(szDrive);
	m_strPath.append(szDir);
	m_strName.assign(szName);

	pszDot = CharNext(szExt);
	m_strExt.assign(pszDot);

	SetDateCreation();
	SetTimeCreation();
	SetDateLastWrite();
	SetTimeLastWrite();
	SetDateLastAccess();
	SetTimeLastAccess();

	//if (0 != FindNextFile(hff, &fd))
	//{
	//	FindClose(hff);
	//	return FWF_ERR_WILDCARD;
	//}
	//FindClose(hff);

	WI_SetFlag(m_State, FileState::Attached);

	return true;
}

bool CWiseFile::Detach()
{
	// does nothing
	return true;
}

bool CWiseFile::SetSize(bool bHex)
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	if (bHex)
	{
		return true;
	}
	else
	{
		// Not using StrFormatByteSize because it wordifies everything
		if (int2str(m_szSize, m_qwSize))
		{
			WI_SetFlag(m_State, FileState::Size);
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
			m_szCRC = std::format(L"{:#08x}", m_dwCRC);
		}
		else
		{
			int2str(m_szCRC, m_dwCRC);
		}
		return true;
	}

	if (WI_IsFlagSet(m_State, FileState::CRC_Pending))
	{
		LoadStringResource(STR_CRC_PENDING, m_szCRC);
		return true;
	}

	if (WI_IsFlagSet(m_State, FileState::CRC_Working))
	{
		LoadStringResource(STR_CRC_WORKING, m_szCRC);
		return true;
	}

	if (WI_IsFlagSet(m_State, FileState::CRC_Error))
	{
		LoadStringResource(STR_CRC_ERROR, m_szCRC);
		return true;
	}
	return false;
}

bool CWiseFile::SetFileVersion()
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	if (!m_fHasVersion)
	{
		return true;
	}

	wchar_t szDec[2];
	GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, szDec, 2);

	// %d%s%02d%s%04d%s%04d
	m_szFileVersion = std::format(L"{}{}{:02d}{}{:04d}{}{:04d}",
		(int)HIWORD(HIDWORD(m_qwFileVersion)),
		szDec,
		(int)LOWORD(HIDWORD(m_qwFileVersion)),
		szDec,
		(int)HIWORD(LODWORD(m_qwFileVersion)),
		szDec,
		(int)LOWORD(LODWORD(m_qwFileVersion)));
	
	return true;
}

bool CWiseFile::GetFileVersion(LPDWORD pdwMS, LPDWORD pdwLS)
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (!m_fHasVersion)
	{
		return true;
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

	if (!m_fHasVersion)
	{
		return true;
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

bool CWiseFile::SetProductVersion()
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (!m_fHasVersion)
	{
		return true;
	}

	WCHAR szDec[2];
	GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, szDec, 2);

	m_szProductVersion = std::format(L"{}{}{:02d}{}{:04d}{}{:04d}",
		(int)HIWORD(HIDWORD(m_qwProductVersion)),
		szDec,
		(int)LOWORD(HIDWORD(m_qwProductVersion)),
		szDec,
		(int)HIWORD(LODWORD(m_qwProductVersion)),
		szDec,
		(int)LOWORD(LODWORD(m_qwProductVersion)));

	return true;
}

bool CWiseFile::GetProductVersion(LPDWORD pdwMS, LPDWORD pdwLS)
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (!m_fHasVersion)
	{
		return true;
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

	if (!m_fHasVersion)
	{
		return true;
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

bool CWiseFile::SetType()
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (!m_fHasVersion)
		return true;

	if (m_fszType)
	{
		return true;
	}

	m_szType.clear();
	switch (m_dwType)
	{
		case VFT_UNKNOWN: m_szType = std::format(L"VFT_UNKNOWN: {:#010x}", m_dwType);
			break;
		case VFT_APP: m_szType = L"VFT_APP";
			break;
		case VFT_DLL: m_szType = L"VFT_DLL";
			break;
		case VFT_DRV: m_szType = L"VFT_DRV";
			break;
		case VFT_FONT: m_szType = L"VFT_FONT";
			break;
		case VFT_VXD: m_szType = L"VFT_VXD";
			break;
		case VFT_STATIC_LIB: m_szType = L"VFT_STATIC_LIB";
			break;
		default: m_szType = std::format(L"Reserved: {:#010x}", m_dwType);
	}
	m_fszType = true;
	return true;
}

bool CWiseFile::SetOS()
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (!m_fHasVersion)
		return true;

	if (m_fszOS)
	{
		return true;
	}

	switch (m_dwOS)
	{
	case VOS_UNKNOWN: m_szOS = std::format(L"VOS_UNKNOWN: {:#010x}", m_dwOS);
			break;
		case VOS_DOS: m_szOS = L"VOS_DOS";
			break;
		case VOS_OS216: m_szOS = L"VOS_OS216";
			break;
		case VOS_OS232: m_szOS = L"VOS_OS232";
			break;
		case VOS_NT: m_szOS = L"VOS_NT";
			break;
		case VOS__WINDOWS16: m_szOS = L"VOS__WINDOWS16";
			break;
		case VOS__PM16: m_szOS = L"VOS__PM16";
			break;
		case VOS__PM32: m_szOS = L"VOS__PM32";
			break;
		case VOS__WINDOWS32: m_szOS = L"VOS__WINDOWS32";
			break;
		case VOS_OS216_PM16: m_szOS = L"VOS_OS216_PM16";
			break;
		case VOS_OS232_PM32: m_szOS = L"VOS_OS232_PM32";
			break;
		case VOS_DOS_WINDOWS16: m_szOS =  L"VOS_DOS_WINDOWS16";
			break;
		case VOS_DOS_WINDOWS32: m_szOS = L"VOS_DOS_WINDOWS32";
			break;
		case VOS_NT_WINDOWS32: m_szOS = L"VOS_NT_WINDOWS32";
			break;
		default:		m_szOS = std::format(L"Reserved: {:#010x}", m_dwOS);
	}
	m_fszOS = true;
	return true;
}

//int CWiseFile::GetOSString(LPWSTR pszText)
//{
//	if (!CheckState(FWFS_VERSION))
//	{
//		lstrinit(pszText);
//		return FWF_ERR_INVALID;
//	}
//
//	if (!m_fHasVersion)
//	{
//		*pszText = 0;
//		return FWF_SUCCESS;
//	}
//
//	if (!m_fszOS)
//	{
//		GetOSString();
//	}
//
//	if (m_fszOS)
//	{
//		wcscpy_s(pszText, 64, m_szOS);
//	}
//
//	return FWF_SUCCESS;
//}

bool CWiseFile::SetCodePage(bool bNumeric)
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (!m_fHasVersion)
	{
		return true;
	}

	if (bNumeric)
	{
		m_szCodePage = std::format(L"{:#06x}", (WORD)m_CodePage);
		//wsprintf(m_szCodePage, L"%04x", (WORD)m_CodePage);
	}
	else
	{
		m_CodePage.GetCodePageName(LANGIDFROMLCID(GetThreadLocale()), m_CodePage, m_szCodePage);
	}
	return true;
}

bool CWiseFile::ReadVersionInfo()
{
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

	m_fHasVersion = true;
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

	return true;
}

bool CWiseFile::SetDateCreation(bool fLocal)
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	SYSTEMTIME* pst = nullptr;
	if (fLocal)
	{
		pst = &m_stLocalCreation;
	}
	else
	{
		pst = &m_stUTCCreation;
	}

	wchar_t szDate[200];
	if (0 == ::GetDateFormat(::GetThreadLocale(), DATE_SHORTDATE, pst, NULL, szDate, 64))
	{
		return false;
	}
	else
	{
		m_szDateCreated = szDate;
		return true;
	}
}

bool CWiseFile::SetDateLastAccess(bool fLocal)
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	SYSTEMTIME* pst = nullptr;
	if (fLocal)
	{
		pst = &m_stLocalLastAccess;
	}
	else
	{
		pst = &m_stUTCLastAccess;
	}

	wchar_t szDate[200];
	if (0 == ::GetDateFormat(::GetThreadLocale(), DATE_SHORTDATE, pst, NULL, szDate, 64))
	{
		return false;
	}
	else
	{
		m_szDateLastAccess = szDate;
		return true;
	}
}

bool CWiseFile::SetDateLastWrite(bool fLocal)
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	SYSTEMTIME* pst = nullptr;
	if (fLocal)
	{
		pst = &m_stLocalLastWrite;
	}
	else
	{
		pst = &m_stUTCLastWrite;
	}

	wchar_t szDate[200];
	if (0 == ::GetDateFormat(::GetThreadLocale(), DATE_SHORTDATE, pst, NULL, szDate, 64))
	{
		return false;
	}
	else
	{
		m_szDateLastWrite = szDate;
		return true;
	}
}

bool CWiseFile::SetTimeCreation(bool fLocal)
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	SYSTEMTIME* pst = nullptr;
	if (fLocal)
	{
		pst = &m_stLocalCreation;
	}
	else
	{
		pst = &m_stUTCCreation;
	}

	wchar_t szTime[200];
	if (0 == ::GetTimeFormat(::GetThreadLocale(), TIME_NOSECONDS, pst, NULL, szTime, 64))
	{
		return false;
	}
	else
	{
		m_szTimeCreated = szTime;
		return true;
	}
}

bool CWiseFile::SetTimeLastAccess(bool fLocal)
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	SYSTEMTIME* pst = nullptr;
	if (fLocal)
	{
		pst = &m_stLocalLastAccess;
	}
	else
	{
		pst = &m_stUTCLastAccess;
	}

	wchar_t szTime[200];
	if (0 == ::GetTimeFormat(::GetThreadLocale(), TIME_NOSECONDS, pst, NULL, szTime, 64))
	{
		return false;
	}
	else
	{
		m_szTimeLastAccess = szTime;
		return true;
	}
}

bool CWiseFile::SetTimeLastWrite(bool fLocal)
{
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		return false;
	}

	SYSTEMTIME* pst = nullptr;
	if (fLocal)
	{
		pst = &m_stLocalLastWrite;
	}
	else
	{
		pst = &m_stUTCLastWrite;
	}

	wchar_t szTime[200];
	if (0 == ::GetTimeFormat(::GetThreadLocale(), TIME_NOSECONDS, pst, NULL, szTime, 64))
	{
		return false;

	}
	else
	{
		m_szTimeLastWrite = szTime;
		return true;
	}
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
	if (!WI_IsFlagSet(m_State, FileState::Attached))
	{
		m_szAttribs = L"        ";
		return false;
	}

	if (m_fszAttribs)
	{
		// the string is already built, bail out
		return true;
	}
	m_fszAttribs = true;

	if (m_dwAttribs & FILE_ATTRIBUTE_NORMAL)
	{
		m_szAttribs = L"        ";
		return true;
	}

	m_szAttribs = std::format(L"{}{}{}{}{}{}{}{}",
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

bool CWiseFile::SetFlags()
{
	if (!WI_IsFlagSet(m_State, FileState::Version))
	{
		return false;
	}

	if (!m_fHasVersion)
	{
		return true;
	}

	if (m_fszFlags)
	{
		return true;
	}
	m_fszFlags = true;

	// list seperator
	wchar_t szSep[10];
	GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLIST, szSep, 2);
	wcscpy_s(szSep, L" \0");

	std::wstring szStr;
	if (m_dwFlags & VS_FF_DEBUG)
	{
		if (m_fDebugStripped)
		{

			LoadStringResource(STR_FLAG_DEBUG_STRIPPED, szStr);
		}
		else
		{
			LoadStringResource(STR_FLAG_DEBUG, szStr);
		}
		m_szFlags += szStr;
		m_szFlags += szSep;
	}
	if (m_dwFlags & VS_FF_PRERELEASE)
	{
		LoadStringResource(STR_FLAG_PRERELEASE, szStr);
		m_szFlags += szStr;
		m_szFlags += szSep;
	}
	if (m_dwFlags & VS_FF_PATCHED)
	{
		LoadStringResource(STR_FLAG_PATCHED, szStr);
		m_szFlags += szStr;
		m_szFlags += szSep;
	}
	if (m_dwFlags & VS_FF_PRIVATEBUILD)
	{
		LoadStringResource(STR_FLAG_PRIVATEBUILD, szStr);
		m_szFlags += szStr;
		m_szFlags += szSep;
	}
	if (m_dwFlags & VS_FF_INFOINFERRED)
	{
		LoadStringResource(STR_FLAG_INFOINFERRED, szStr);
		m_szFlags += szStr;
		m_szFlags += szSep;
	}
	if (m_dwFlags & VS_FF_SPECIALBUILD)
	{
		LoadStringResource(STR_FLAG_SPECIALBUILD, szStr);
		m_szFlags += szStr;
		m_szFlags += szSep;
	}

	if (!m_szFlags.empty())
	{
		const size_t sep = m_szFlags.find_last_of(szSep[0]);
		m_szFlags.resize(sep);
	}
	// if we've looked at it, and we still don't have a string for it
	// put a default one in (unless the flags are 00000000)
	if ( (m_szFlags.length() < 3) && (m_dwFlags != 0))
	{
		m_szFlags = std::format(L"{:010x}", m_dwFlags);
	}

	return true;
}

bool CWiseFile::SetLanguage(UINT Language)
{
	wchar_t szTemp[255];

	if (!Language)
	{
		VerLanguageName(Language, szTemp, 255);
		m_szLanguage = std::format(L"%lu %s", Language, szTemp);
		WI_SetFlag(m_State, FileState::Language);
		return true;
	}

	if (!GetLocaleInfo(Language, LOCALE_SLANGUAGE, szTemp, 255))
	{
		VerLanguageName(Language, szTemp, 255);
	}
	m_szLanguage = std::format(L"%lu %s", Language, szTemp);
	WI_SetFlag(m_State, FileState::Language);
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