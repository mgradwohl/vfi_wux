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
#include <winnt.h>
#include <wil/cppwinrt.h>
#include <wil/common.h>
#include "tcodepage.h"

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
		CRC_Pending		= 0x0020,
		CRC_Working		= 0x0040,		// working on CRC
		CRC_Complete	= 0x0080,		// CRC complete
		CRC_Error		= 0x0100,		// error generating CRC
		PendingDeletion = 0x0200,		// pending deletion from list
	}; 
	DEFINE_ENUM_FLAG_OPERATORS(FileState);

class CWiseFile
{
public:

private:
	// Member Variables
	// Set when a file is added
	std::wstring		m_strFullPath;
	std::wstring		m_strPath;
	std::wstring		m_strName;
	std::wstring		m_strExt;

	// Set when information requested, if not set
	std::wstring		m_strAttribs;
	std::wstring		m_strFlags;
	std::wstring		m_strOS;
	std::wstring		m_strType;
	std::wstring		m_strSize;
	std::wstring		m_strDateCreated;
	std::wstring		m_strTimeCreated;
	std::wstring		m_strDateLastAccess;
	std::wstring		m_strTimeLastAccess;
	std::wstring		m_strDateLastWrite;
	std::wstring		m_strTimeLastWrite;
	std::wstring		m_strLanguage;
	std::wstring		m_strCodePage;
	std::wstring		m_strCRC;
	std::wstring		m_strFileVersion;
	std::wstring		m_strProductVersion;

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
	bool Attach(std::wstring strFileSpec);
	bool Detach();

	// handlers for listview
	const std::wstring& GetFieldString(int iField, bool fOptions);

	bool ReadVersionInfoEx();

	const std::wstring& GetFullPath()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strFullPath;
		}
		return L"\0";
	}

	const std::wstring& GetPath()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strPath;
		}
		return L"\0";
	}

	const std::wstring& GetName()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strName;
		}
		return L"\0";
	}

	const std::wstring& GetExt()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strExt;
		}
		return L"\0";
	}

	const std::wstring& GetSize()
	{
		if (WI_IsFlagSet(m_State, FileState::Size))
		{
			return m_strSize;
		}
		SetSize(false);
		return m_strSize;
	}

	uint64_t Size()
	{
		return m_qwSize;
	}

	const std::wstring& GetCRC()
	{
		if (WI_IsFlagSet(m_State, FileState::CRC_Complete))
		{
			return m_strCRC;
		}
		SetCRC(true);
		return m_strCRC;
	}

	const std::wstring& GetOS()
	{
		if (WI_IsFlagSet(m_State, FileState::Version) && m_fszOS)
		{
			return m_strOS;
		}
		SetOS();
		return m_strOS;
	}

	const std::wstring& GetType()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached) && m_fszType)
		{
			return m_strType;
		}
		SetType();
		return m_strType;
	}

	const std::wstring& GetDateLastAccess()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strDateLastAccess;
		}
		return L"\0";
	}

	const std::wstring& GetTimeLastAccess()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strTimeLastAccess;
		}
		return L"\0";
	}

	const std::wstring& GetDateCreated()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strDateCreated;
		}
		return L"\0";
	}

	const std::wstring& GetTimeCreated()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strTimeCreated;
		}
		return L"\0";
	}

	const std::wstring& GetDateLastWrite()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strDateLastWrite;
		}
		return L"\0";
	}

	const std::wstring& GetTimeLastWrite()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strTimeLastWrite;
		}
		return L"\0";
	}

	const std::wstring& GetAttribs()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached) && m_fszAttribs)
		{
			return m_strAttribs;
		}
		SetAttribs();
		return m_strAttribs;
	}

	const std::wstring& GetFileVersion()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strFileVersion;
		}

		if (SetProductVersion() && SetFileVersion())
		{
			WI_SetFlag(m_State, FileState::Version);
		}
		return m_strFileVersion;
	}

	const std::wstring& GetProductVersion()
	{
		if (WI_IsFlagSet(m_State, FileState::Version))
		{
			return m_strProductVersion;
		}

		if (SetProductVersion() && SetFileVersion())
		{
			WI_SetFlag(m_State, FileState::Version);
		}
		return m_strProductVersion;
	}

	const std::wstring& GetLanguage()
	{
		if (WI_IsFlagSet(m_State, FileState::Language))
		{
			return m_strLanguage;
		}
		SetLanguage(m_wLanguage);
		return m_strLanguage;
	}

	const std::wstring& GetCodePage()
	{
		if (WI_IsFlagSet(m_State, FileState::Language))
		{
			return m_strCodePage;
		}
		SetCodePage(m_CodePage);
		return m_strCodePage;
	}

	const std::wstring& GetFlags()
	{
		if (WI_IsFlagSet(m_State, FileState::Version) && m_fszFlags)
		{
			return m_strFlags;
		}
		SetFlags();
		return m_strFlags;
	}

	bool SetSize(bool bHex = false);
	bool SetAttribs();

	bool SetDateCreation(bool fLocal = true);
	bool SetDateLastAccess(bool fLocal = true);
	bool SetDateLastWrite(bool fLocal = true);

	bool SetTimeCreation(bool fLocal = true);
	bool SetTimeLastAccess(bool fLocal = true);
	bool SetTimeLastWrite(bool fLocal = true);

	bool SetFileVersion();
	bool GetFileVersion(LPDWORD pdwHigh, LPDWORD pdwLow);
	bool GetFileVersion(LPWORD pwHighMS, LPWORD pwLowMS, LPWORD pwHighLS, LPWORD pwLowLS);

	bool SetProductVersion();
	bool GetProductVersion(LPDWORD pdwMS, LPDWORD pdwLS);
	bool GetProductVersion(LPWORD pwHighMS, LPWORD pwLowMS, LPWORD pwHighLS, LPWORD pwLowLS);

	bool SetLanguage(UINT Language);
	bool SetCodePage(bool bNumeric = false);

	bool SetOS();
	bool SetType();
	bool SetFlags();

	bool SetCRC(bool bHex = true);

	bool TouchFileTime(FILETIME* lpTime);

	void _SetCRC(DWORD dwCRC)
	{
		m_dwCRC = dwCRC;
	}
	// execute helpers
	bool  ReadVersionInfo();
};