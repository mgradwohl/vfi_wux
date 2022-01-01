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

#pragma once
#include <wil/common.h>
#include "tcodepage.h"

class CWiseFile
{
public:
	// file states
	// get rid of these and just use bools e.g. hasCRC, hasVersion etc.
	// this is just too many confusing states
	enum class FileState
	{
		Invalid			= 0x0000,		// invalid state
		Valid			= 0x0001,		// constructed only
		Attached		= 0x0002,		// attached to a file successfully
		Size			= 0x0004,		// size information is valid
		Version			= 0x0008,		// version information is read
		Language		= 0x0010,		// language information is read
		CRC_Working		= 0x0020,		// working on CRC
		CRC_Complete	= 0x0040,		// CRC complete
		CRC_Error		= 0x0080,		// error generating CRC
		PendingDeletion = 0x0100,		// pending deletion from list
	}; 
	DEFINE_ENUM_FLAG_OPERATORS(FileState);

private:
	// Member Variables
	// Set when a file is added
	std::wstring		m_strFullPath;
	std::wstring		m_strPath;
	std::wstring		m_strName;
	std::wstring		m_strExt;

	// Set when information requested, if not set
	std::wstring		m_szAttribs;
	std::wstring		m_szFlags;
	std::wstring		m_szOS;
	std::wstring		m_szType;
	std::wstring		m_szSize;
	std::wstring		m_szDateCreated;
	std::wstring		m_szTimeCreated;
	std::wstring		m_szDateLastAccess;
	std::wstring		m_szTimeLastAccess;
	std::wstring		m_szDateLastWrite;
	std::wstring		m_szTimeLastWrite;
	std::wstring		m_szLanguage;
	std::wstring		m_szCodePage;
	std::wstring		m_szCRC;
	std::wstring		m_szFileVersion;
	std::wstring		m_szProductVersion;

	// from WIN32_FIND_DATA
	DWORD		m_dwAttribs;
	SYSTEMTIME	m_stUTCCreation;
	SYSTEMTIME	m_stUTCLastAccess;
	SYSTEMTIME	m_stUTCLastWrite;

	// set when requested, if not set
	SYSTEMTIME	m_stLocalCreation;
	SYSTEMTIME	m_stLocalLastAccess;
	SYSTEMTIME	m_stLocalLastWrite;

	QWORD		m_qwSize;
	QWORD		m_qwFileVersion;
	QWORD		m_qwProductVersion;
	WORD		m_wLanguage;
	TCodePage	m_CodePage;
	DWORD		m_dwOS;
	DWORD		m_dwType;
	DWORD		m_dwFlags;

	DWORD		m_dwCRC;
	FileState	m_State;

	// Flags
	bool m_fDebugStripped;
	bool m_fHasVersion;
	bool m_fszOS;
	bool m_fszType;
	bool m_fszFlags;
	bool m_fszAttribs;

public:
	// construction, destruction
	CWiseFile();
	CWiseFile(std::wstring strFileSpec);
	~CWiseFile();

	// initialization, release
	int Attach(std::wstring strFileSpec);
	int Detach();

	// handlers for listview
	const std::wstring& GetFieldString(int iField, bool fOptions);

	int ReadVersionInfoEx();

	const std::wstring& GetFullPath()
	{
		if (WI_IsFlagSet(m_State, CWiseFile::FileState::Attached))
		{
			return m_strFullPath;
		}
		return L"\0";
	}

	const std::wstring& GetPath()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_strPath;
		}
		return L"\0";
	}

	const std::wstring& GetName()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_strName;
		}
		return L"\0";
	}

	const std::wstring GetExt()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{	
			return m_strExt;
		}
		return L"\0";
	}

	const std::wstring GetSize()
	{
		if (m_State == CWiseFile::FileState::FWFS_SIZE)
		{
			return m_szSize;
		}
		SetSize(false);
		return m_szSize;
	}

	uint64_t Size()
	{
		return m_qwSize;
	}

	const std::wstring& GetCRC()
	{
		if (m_State == CWiseFile::FileState::FWFS_CRC_COMPLETE)
		{
			SetCRC(true);
		}
		return m_szCRC;
	}

	const std::wstring& GetOS()
	{
		if (m_State == CWiseFile::FileState::FWFS_VERSION)
		{
			SetOS();
		}
		return m_szOS;
	}

	const std::wstring& GetType()
	{
		if (m_State == CWiseFile::FileState::FWFS_VERSION)
		{
			SetType();
		}
		return m_szType;
	}

	const std::wstring& GetDateLastAccess()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_szDateLastAccess;
		}
		return L"\0";
	}

	const std::wstring& GetTimeLastAccess()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_szTimeLastAccess;
		}
		return L"\0";
	}

	const std::wstring& GetDateCreated()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_szDateCreated;
		}
		return L"\0";
	}

	const std::wstring& GetTimeCreated()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_szTimeCreated;
		}
		return L"\0";
	}

	const std::wstring& GetDateLastWrite()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_szDateLastWrite;
		}
		return L"\0";
	}

	const std::wstring& GetTimeLastWrite()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_szTimeLastWrite;
		}
		return L"\0";
	}

	const std::wstring& GetAttribs()
	{
		if (m_State == CWiseFile::FileState::FWFS_ATTACHED)
		{
			return m_szAttribs;
		}
		SetAttribs();
		return m_szAttribs;
	}

	const std::wstring& GetFileVersion()
	{
		if (m_State == CWiseFile::FileState::FWFS_VERSION)
		{
			return m_szFileVersion;
		}

		if (FWF_SUCCESS == SetProductVersion() && FWF_SUCCESS == SetFileVersion())
		{
			m_State = CWiseFile::FileState::FWFS_VERSION/* || CWiseFile::FileState::FWFS_ATTACHED*/;
		}
		return m_szFileVersion;
	}

	const std::wstring& GetProductVersion()
	{
		if (m_State == CWiseFile::FileState::FWFS_VERSION)
		{
			return m_szProductVersion;
		}

		if (FWF_SUCCESS == SetProductVersion() && FWF_SUCCESS == SetFileVersion())
		{
			m_State = CWiseFile::FileState::FWFS_VERSION/* || CWiseFile::FileState::FWFS_ATTACHED*/;
		}
		return m_szProductVersion;
	}

	const std::wstring& GetLanguage()
	{
		if (m_State == CWiseFile::FileState::FWFS_LANGUAGE)
		{
			return m_szLanguage;
		}
		SetLanguage(m_wLanguage);
		return m_szLanguage;
	}

	const std::wstring& GetCodePage()
	{
		if (m_State == CWiseFile::FileState::FWFS_VERSION)
		{
			return m_szCodePage;
		}
		SetCodePage(m_CodePage);
		return m_szCodePage;
	}

	const std::wstring& GetFlags()
	{
		if (m_State == CWiseFile::FileState::FWFS_VERSION)
		{
			return m_szFlags;
		}
		SetFlags();
		return m_szFlags;
	}

	int SetSize(bool bHex = false);
	int SetAttribs();

	int SetDateCreation(bool fLocal = true);
	int SetDateLastAccess(bool fLocal = true);
	int SetDateLastWrite(bool fLocal = true);

	int SetTimeCreation(bool fLocal = true);
	int SetTimeLastAccess(bool fLocal = true);
	int SetTimeLastWrite(bool fLocal = true);

	int SetFileVersion();
	int GetFileVersion(LPDWORD pdwHigh, LPDWORD pdwLow);
	int GetFileVersion(LPWORD pwHighMS, LPWORD pwLowMS, LPWORD pwHighLS, LPWORD pwLowLS);

	int SetProductVersion();
	int GetProductVersion(LPDWORD pdwMS, LPDWORD pdwLS);
	int GetProductVersion(LPWORD pwHighMS, LPWORD pwLowMS, LPWORD pwHighLS, LPWORD pwLowLS);

	bool SetLanguage(UINT Language);
	int SetCodePage(bool bNumeric = false);

	void SetOS();
	void SetType();
	int SetFlags();

	int SetCRC(bool bHex = true);

	void SetState(FileState State)
	{
		m_State = State;
	}

	FileState GetState()
	{
		return m_State;
	}

	int TouchFileTime(FILETIME* lpTime);

	void _SetCRC(DWORD dwCRC)
	{
		m_dwCRC = dwCRC;
	}
	// execute helpers
	int ReadVersionInfo();
};