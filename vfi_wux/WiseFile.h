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
		Invalid			= 0x00000,		// invalid state
		Valid			= 0x00001,		// constructed only
		Attached		= 0x00002,		// attached to a file successfully
		Date			= 0x00004,
		DateString		= 0x00008,
		Size			= 0x00010,		// size information is valid
		SizeString		= 0x00020,
		Version			= 0x00040,		// version information has been read
		VersionValid	= 0x00080,
		VersionString	= 0x00100,
		Language		= 0x00200,		// language information is read
		LanguageString	= 0x00400,
		CRC_Pending		= 0x00800,
		CRC_Working		= 0x01000,		// working on CRC
		CRC_Complete	= 0x02000,		// CRC complete
		CRC_Error		= 0x04000,		// error generating CRC
		CRC_String		= 0x08000,
		PendingDeletion = 0x10000,		// pending deletion from list
	}; 
	DEFINE_ENUM_FLAG_OPERATORS(FileState);

class CWiseFile
{
public:

private:
	// Member Variables
	// Set when Attach() called
	std::wstring		m_strFullPath;
	std::wstring		m_strPath;
	std::wstring		m_strName;
	std::wstring		m_strExt;
	std::wstring		m_strAttribs;
	std::wstring		m_strSize;
	std::wstring		m_strDateCreated;
	std::wstring		m_strTimeCreated;
	std::wstring		m_strDateLastAccess;
	std::wstring		m_strTimeLastAccess;
	std::wstring		m_strDateLastWrite;
	std::wstring		m_strTimeLastWrite;

	// Set after ReadVersionInfo and SetVersionStrings is called
	std::wstring		m_strFlags;
	std::wstring		m_strOS;
	std::wstring		m_strType;
	std::wstring		m_strFileVersion;
	std::wstring		m_strProductVersion;

	// Set after ReadVersionInfo and SetLanguage is called
	std::wstring		m_strLanguage;
	std::wstring		m_strCodePage;

	// set when CRC is calculated and complete
	std::wstring		m_strCRC;

	// from WIN32_FIND_DATA
	DWORD		m_dwAttribs;
	SYSTEMTIME	m_stUTCCreation;
	SYSTEMTIME	m_stUTCLastAccess;
	SYSTEMTIME	m_stUTCLastWrite;

	// set when Attach() called
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

	// Version Flags
	// todo incorporate into FileState::
	bool m_fDebugStripped;

public:
	// construction, destruction
	CWiseFile();
	CWiseFile(const std::wstring& strFileSpec);
	CWiseFile(CWiseFile& wf) = delete;
	CWiseFile const& operator=(CWiseFile& wf) = delete;
	~CWiseFile();

	// initialization, release
	bool Attach(const std::wstring& strFileSpec) const;
	bool Detach();

	// handlers for listview
	const std::wstring& GetFieldString(int iField, bool fOptions);

	// currently unused
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
		if (WI_IsFlagSet(m_State, FileState::SizeString))
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
		if (WI_IsFlagSet(m_State, FileState::CRC_String))
		{
			return m_strCRC;
		}
		SetCRC(true);
		return m_strCRC;
	}

	const std::wstring& GetOS()
	{
		if (WI_IsFlagSet(m_State, FileState::VersionString))
		{
			return m_strOS;
		}

		if (WI_IsFlagSet(m_State, FileState::VersionValid))
		{
			SetVersionStrings();
		}
		return m_strOS;
	}

	const std::wstring& GetType()
	{
		if (WI_IsFlagSet(m_State, FileState::VersionString))
		{
			return m_strType;
		}

		if (WI_IsFlagSet(m_State, FileState::VersionValid))
		{
			SetVersionStrings();
		}
		return m_strType;
	}

	const std::wstring& GetDateLastAccess() const
	{
		if (WI_IsFlagSet(m_State, FileState::DateString))
		{
			return m_strDateLastAccess;
		}
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			SetDateTimeStrings();
			WI_SetFlag(m_State, FileState::DateString);
		}
		return m_strDateLastAccess;
	}

	const std::wstring& GetTimeLastAccess() const
	{
		if (WI_IsFlagSet(m_State, FileState::DateString))
		{
			return m_strTimeLastAccess;
		}
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			SetDateTimeStrings();
			WI_SetFlag(m_State, FileState::DateString);
		}
		return m_strTimeLastAccess;
	}

	const std::wstring& GetDateCreated() const
	{
		if (WI_IsFlagSet(m_State, FileState::DateString))
		{
			return m_strDateCreated;
		}
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			SetDateTimeStrings();
			WI_SetFlag(m_State, FileState::DateString);
		}
		return m_strDateCreated;
	}

	const std::wstring& GetTimeCreated() const
	{
		if (WI_IsFlagSet(m_State, FileState::DateString))
		{
			return m_strTimeCreated;
		}
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			SetDateTimeStrings();
			WI_SetFlag(m_State, FileState::DateString);
		}
		return m_strTimeCreated;
	}

	const std::wstring& GetDateLastWrite() const
	{
		if (WI_IsFlagSet(m_State, FileState::DateString))
		{
			return m_strDateLastWrite;
		}
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			SetDateTimeStrings();
			WI_SetFlag(m_State, FileState::DateString);
		}
		return m_strDateLastWrite;
	}

	const std::wstring& GetTimeLastWrite() const
	{
		if (WI_IsFlagSet(m_State, FileState::DateString))
		{
			return m_strTimeLastWrite;
		}
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			SetDateTimeStrings();
			WI_SetFlag(m_State, FileState::DateString);
		}
		return m_strTimeLastWrite;
	}

	const std::wstring& GetAttribs()
	{
		if (WI_IsFlagSet(m_State, FileState::Attached))
		{
			return m_strAttribs;
		}
		//SetAttribs();
		return m_strAttribs;
	}

	const std::wstring& GetFileVersion()
	{
		if (WI_IsFlagSet(m_State, FileState::VersionString))
		{
			return m_strFileVersion;
		}

		if (WI_IsFlagSet(m_State, FileState::VersionValid))
		{
			SetVersionStrings();
		}

		return m_strFileVersion;
	}

	const std::wstring& GetProductVersion()
	{
		if (WI_IsFlagSet(m_State, FileState::VersionString))
		{
			return m_strProductVersion;
		}

		if (WI_IsFlagSet(m_State, FileState::VersionValid))
		{
			SetVersionStrings();
		}
		return m_strProductVersion;
	}

	const std::wstring& GetLanguage()
	{
		if (WI_IsFlagSet(m_State, FileState::LanguageString))
		{
			return m_strLanguage;
		}
		SetLanguage(m_wLanguage, false);
		return m_strLanguage;
	}

	const std::wstring& GetCodePage()
	{
		if (WI_IsFlagSet(m_State, FileState::LanguageString))
		{
			return m_strCodePage;
		}
		SetLanguage(m_wLanguage, false);
		return m_strCodePage;
	}

	const std::wstring& GetFlags()
	{
		if (WI_IsFlagSet(m_State, FileState::VersionString))
		{
			return m_strFlags;
		}
		SetVersionStrings();
		return m_strFlags;
	}

	bool SetSize(bool bHex = false);
	bool SetAttribs();

	bool SetDateTimeStrings(bool fLocal = true);

	bool SetVersionStrings();
	bool GetFileVersion(LPDWORD pdwHigh, LPDWORD pdwLow);
	bool GetFileVersion(LPWORD pwHighMS, LPWORD pwLowMS, LPWORD pwHighLS, LPWORD pwLowLS);

	bool GetProductVersion(LPDWORD pdwMS, LPDWORD pdwLS);
	bool GetProductVersion(LPWORD pwHighMS, LPWORD pwLowMS, LPWORD pwHighLS, LPWORD pwLowLS);

	bool SetLanguage(UINT Language, bool bNumeric = false);

	bool SetCRC(bool bHex = true);
	void _SetCRC(DWORD dwCRC)
	{
		m_dwCRC = dwCRC;
	}

	bool TouchFileTime(FILETIME* lpTime);
	bool  ReadVersionInfo();
};