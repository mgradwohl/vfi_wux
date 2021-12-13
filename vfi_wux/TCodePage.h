// Visual File Information
// Copyright (c) Microsoft Corporation
// All rights reserved. 
// 
// MIT License
// 
// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the ""Software"", 
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

#include "pch.h"
//#include <mlang.h>

class TCodePage  
{
public:
	TCodePage()
	{
		m_wCodePage = 0;
		m_pszCodePage = NULL;
		m_pML = NULL;

		if (SUCCEEDED(CoInitialize(NULL)))
		{
			if (FAILED(CoCreateInstance(CLSID_CMultiLanguage, NULL, CLSCTX_INPROC_SERVER, IID_IMultiLanguage2, (void**)&m_pML)))
			{
				m_pML = NULL;
			}
		}
	}

	TCodePage(WORD wCodePage)
	{
		m_wCodePage = wCodePage;
		m_pszCodePage = NULL;
		CoUninitialize();
	}

	bool GetCodePageName(WORD wLanguage, UINT CodePage, LPWSTR pszBuf)
	{
		if (NULL == pszBuf)
			return false;
		
		if (m_pML)
		{
			ZeroMemory(&m_mcp, sizeof(MIMECPINFO));
			if (S_OK  == m_pML->GetCodePageInfo(CodePage, wLanguage, &m_mcp))
			{
				wsprintf( pszBuf, L"%s (%lu)", m_mcp.wszDescription, CodePage);
				return true;
			}
		}

		CPINFOEX cp;
		ZeroMemory(&cp, sizeof(CPINFOEX));
		if (GetCPInfoEx(CodePage, 0, &cp))
		{
			wsprintf( pszBuf, L"%s", cp.CodePageName);
			return true;
		}
		
		switch (CodePage)
		{
			case	37:		wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - U.S./Canada");
			break;

			case	709:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Arabic (ASMO 449+, BCON V4)");
			break;

			case	710:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Arabic (Transparent Arabic)");
			break;

			case	437:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - United States");
			break;

			case	500:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - International");
			break;

			case	708:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Arabic - ASMO");
			break;

			case	720:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Arabic - Transparent ASMO");
			break;

			case	737:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Greek 437G");
			break;

			case	775:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Baltic");
			break;

			case	850:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Multilingual Latin I");
			break;

			case	852:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Latin II");
			break;

			case	855:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Cyrillic");
			break;

			case	857:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Turkish");
			break;

			case	860:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Portuguese");
			break;

			case	861:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Icelandic");
			break;

			case	862:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Hebrew");
			break;

			case	863:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Canadian French");
			break;

			case	864:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Arabic");
			break;

			case	865:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Nordic");
			break;

			case	866:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Russian");
			break;

			case	869:	wsprintf(pszBuf,L"%lu %s",CodePage,L"OEM - Modern Greek");
			break;

			case	870:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Multilingual/ROECE (Latin-2)");
			break;

			case	874:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI/OEM - Thai");
			break;

			case	875:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Modern Greek");
			break;

			case	932:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI/OEM - Japanese Shift-JIS");
			break;

			case	936:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI/OEM - Simplified Chinese GBK");
			break;

			case	949:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI/OEM - Korean");
			break;

			case	950:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI/OEM - Traditional Chinese Big5");
			break;

			case	1200:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Unicode (BMP of ISO 10646)");
			break;

			case	1026:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Turkish (Latin-5)");
			break;

			case	1250:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI - Central Europe");
			break;

			case	1251:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI - Cyrillic");
			break;

			case	1252:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI - Latin I");
			break;

			case	1253:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI - Greek");
			break;

			case	1254:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI - Turkish");
			break;

			case	1255:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI - Hebrew");
			break;

			case	1256:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI - Arabic");
			break;

			case	1257:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI - Baltic");
			break;

			case	1258:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ANSI/OEM - Viet Nam");
			break;

			case	1361:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Korean - Johab");
			break;

			case	10000:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Roman");
			break;

			case	10001:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Japanese");
			break;

			case	10002:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Traditional Chinese Big5");
			break;

			case	10003:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Korean");
			break;

			case	10004:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Arabic");
			break;

			case	10005:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Hebrew");
			break;

			case	10006:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Greek I");
			break;

			case	10007:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Cyrillic");
			break;

			case	10008:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Simplified Chinese GB 2312");
			break;

			case	10010:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Romania");
			break;

			case	10017:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Ukraine");
			break;

			case	10029:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Latin II");
			break;

			case	10079:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Icelandic");
			break;

			case	10081:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Turkish");
			break;

			case	10082:	wsprintf(pszBuf,L"%lu %s",CodePage,L"MAC - Croatia");
			break;

			case	20866:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Russian - KOI8");
			break;

			case	29001:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Europa 3");
			break;

			case	65000:	wsprintf(pszBuf,L"%lu %s",CodePage,L"UTF-7");
			break;

			case	65001:	wsprintf(pszBuf,L"%lu %s",CodePage,L"UTF-8");
			break;

			case	20000:	wsprintf(pszBuf,L"%lu %s",CodePage,L"CNS - Taiwan");
			break;

			case	20001:	wsprintf(pszBuf,L"%lu %s",CodePage,L"TCA - Taiwan");
			break;

			case	20002:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Eten - Taiwan");
			break;

			case	20003:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM5550 - Taiwan");
			break;

			case	20004:	wsprintf(pszBuf,L"%lu %s",CodePage,L"TeleText - Taiwan");
			break;

			case	20005:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Wang - Taiwan");
			break;

			case	20105:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IA5 IRV International Alphabet No. 5");
			break;

			case	20106:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IA5 German");
			break;

			case	20107:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IA5 Swedish");
			break;

			case	20108:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IA5 Norwegian");
			break;

			case	20127:	wsprintf(pszBuf,L"%lu %s",CodePage,L"US-ASCII");
			break;

			case	20261:	wsprintf(pszBuf,L"%lu %s",CodePage,L"T.61");
			break;

			case	20269:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 6937 Non-Spacing Accent");
			break;

			case	20273:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Germany");
			break;

			case	20277:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Denmark/Norway");
			break;

			case	20278:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Finland/Sweden");
			break;

			case	20280:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Italy");
			break;

			case	20284:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Latin America/Spain");
			break;

			case	20285:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - United Kingdom");
			break;

			case	20290:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Japanese Katakana Extended");
			break;

			case	20297:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - France");
			break;

			case	20420:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Arabic");
			break;

			case	20423:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Greek");
			break;

			case	20424:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Hebrew");
			break;

			case	20833:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Korean Extended");
			break;

			case	20838:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Thai");
			break;

			case	20871:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Icelandic");
			break;

			case	20880:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Cyrillic,(Russian)");
			break;

			case	20905:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Turkish");
			break;

			case	21025:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Cyrillic (Serbian, Bulgarian)");
			break;

			case	21027:	wsprintf(pszBuf,L"%lu %s",CodePage,L"Ext Alpha Lowercase");
			break;

			case	28591:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-1 Latin I");
			break;

			case	28592:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-2 Central Europe");
			break;

			case	28593:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-3 Turkish");
			break;

			case	28594:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-4 Baltic");
			break;

			case	28595:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-5 Cyrillic");
			break;

			case	28596:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-6 Arabic");
			break;

			case	28597:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-7 Greek");
			break;

			case	28598:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-8 Hebrew");
			break;

			case	28599:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO 8859-9 Latin 5");
			break;

			case	50220:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO-2022 Japanese with no halfwidth Katakana");
			break;

			case	50221:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO-2022 Japanese with halfwidth Katakana");
			break;

			case	50222:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO-2022 Japanese JIS X 0201-1989");
			break;

			case	50225:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO-2022 Korean");
			break;

			case	50227:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO-2022 Simplified Chinese");
			break;

			case	50229:	wsprintf(pszBuf,L"%lu %s",CodePage,L"ISO-2022 Traditional Chinese");
			break;

			case	52936:	wsprintf(pszBuf,L"%lu %s",CodePage,L"HZ-GB2312 Simplified Chinese");
			break;

			case	51932:	wsprintf(pszBuf,L"%lu %s",CodePage,L"EUC-Japanese");
			break;

			case	51949:	wsprintf(pszBuf,L"%lu %s",CodePage,L"EUC-Korean");
			break;

			case	51936:	wsprintf(pszBuf,L"%lu %s",CodePage,L"EUC-Simplified Chinese");
			break;

			case	51950:	wsprintf(pszBuf,L"%lu %s",CodePage,L"EUC-Traditional Chinese");
			break;

			case	50930:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Japanese (Katakana) Extended and Japanese");
			break;

			case	50931:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - US/Canada and Japanese");
			break;

			case	50933:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Korean Extended and Korean");
			break;

			case	50935:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Simplified Chinese");
			break;

			case	50937:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - US/Canada and Traditional Chinese");
			break;

			case	50939:	wsprintf(pszBuf,L"%lu %s",CodePage,L"IBM EBCDIC - Japanese (Latin) Extended and Japanese");
			break;

			default:	wsprintf( pszBuf, L"%lu Unknown", CodePage);
			break;
		}

		return true;
	}

	size_t GetStringLength()
	{
		return wcslen(m_pszCodePage);
	}

	virtual ~TCodePage()
	{
		m_wCodePage = 0;
		delete [] m_pszCodePage;
	}

	operator LPCTSTR()
	{
		return NULL;
	}

	operator WORD() const
	{
		return m_wCodePage;
	}

	void Set(WORD wCodePage)
	{
		m_wCodePage = wCodePage;
	}

	WORD Get(void) const
	{
		return m_wCodePage;
	}

	const TCodePage& operator=(const TCodePage& cp)
	{
		m_wCodePage = (WORD)cp;
		return *this;
	}

	const TCodePage& operator=(const WORD w)
	{
		m_wCodePage = w;
		return *this;
	}

	bool operator==(const TCodePage& cp)
	{
		return m_wCodePage == (WORD)cp;
	}

	bool operator!=(const TCodePage& cp)
	{
		return m_wCodePage != (WORD)cp;
	}

	bool operator==(const WORD w)
	{
		return m_wCodePage == w;
	}

	bool operator!=(const WORD w)
	{
		return m_wCodePage != w;
	}

private:
	IMultiLanguage2* m_pML;

	MIMECPINFO m_mcp;
	WORD m_wCodePage;
	LPWSTR m_pszCodePage;
};
