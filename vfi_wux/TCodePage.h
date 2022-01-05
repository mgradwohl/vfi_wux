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

class TCodePage  
{
public:
	TCodePage()
	{
		m_mcp = {};
		m_wCodePage = 0;
		m_pszCodePage.clear();
		m_pML = NULL;

		if (SUCCEEDED(CoInitialize(NULL)))
		{
			if (FAILED(CoCreateInstance(CLSID_CMultiLanguage, NULL, CLSCTX_INPROC_SERVER, IID_IMultiLanguage2, (void**)&m_pML)))
			{
				m_pML = NULL;
			}
		}
	}

	bool GetCodePageName(WORD wLanguage, UINT CodePage, std::wstring& pszBuf)
	{
		if (m_pML)
		{
			ZeroMemory(&m_mcp, sizeof(MIMECPINFO));
			if (S_OK  == m_pML->GetCodePageInfo(CodePage, wLanguage, &m_mcp))
			{
				pszBuf = std::format(L"{} {})", m_mcp.wszDescription, CodePage);
				return true;
			}
		}

		CPINFOEX cp;
		ZeroMemory(&cp, sizeof(CPINFOEX));
		if (GetCPInfoEx(CodePage, 0, &cp))
		{
			pszBuf = std::format(L"{}", cp.CodePageName);
			return true;
		}
		
		switch (CodePage)
		{
			case	37:		pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - U.S./Canada");
			break;

			case	709:	pszBuf = std::format(L"{} {}",CodePage,L"Arabic (ASMO 449+, BCON V4)");
			break;

			case	710:	pszBuf = std::format(L"{} {}",CodePage,L"Arabic (Transparent Arabic)");
			break;

			case	437:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - United States");
			break;

			case	500:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - International");
			break;

			case	708:	pszBuf = std::format(L"{} {}",CodePage,L"Arabic - ASMO");
			break;

			case	720:	pszBuf = std::format(L"{} {}",CodePage,L"Arabic - Transparent ASMO");
			break;

			case	737:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Greek 437G");
			break;

			case	775:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Baltic");
			break;

			case	850:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Multilingual Latin I");
			break;

			case	852:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Latin II");
			break;

			case	855:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Cyrillic");
			break;

			case	857:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Turkish");
			break;

			case	860:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Portuguese");
			break;

			case	861:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Icelandic");
			break;

			case	862:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Hebrew");
			break;

			case	863:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Canadian French");
			break;

			case	864:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Arabic");
			break;

			case	865:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Nordic");
			break;

			case	866:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Russian");
			break;

			case	869:	pszBuf = std::format(L"{} {}",CodePage,L"OEM - Modern Greek");
			break;

			case	870:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Multilingual/ROECE (Latin-2)");
			break;

			case	874:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI/OEM - Thai");
			break;

			case	875:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Modern Greek");
			break;

			case	932:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI/OEM - Japanese Shift-JIS");
			break;

			case	936:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI/OEM - Simplified Chinese GBK");
			break;

			case	949:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI/OEM - Korean");
			break;

			case	950:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI/OEM - Traditional Chinese Big5");
			break;

			case	1200:	pszBuf = std::format(L"{} {}",CodePage,L"Unicode (BMP of ISO 10646)");
			break;

			case	1026:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Turkish (Latin-5)");
			break;

			case	1250:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI - Central Europe");
			break;

			case	1251:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI - Cyrillic");
			break;

			case	1252:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI - Latin I");
			break;

			case	1253:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI - Greek");
			break;

			case	1254:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI - Turkish");
			break;

			case	1255:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI - Hebrew");
			break;

			case	1256:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI - Arabic");
			break;

			case	1257:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI - Baltic");
			break;

			case	1258:	pszBuf = std::format(L"{} {}",CodePage,L"ANSI/OEM - Viet Nam");
			break;

			case	1361:	pszBuf = std::format(L"{} {}",CodePage,L"Korean - Johab");
			break;

			case	10000:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Roman");
			break;

			case	10001:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Japanese");
			break;

			case	10002:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Traditional Chinese Big5");
			break;

			case	10003:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Korean");
			break;

			case	10004:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Arabic");
			break;

			case	10005:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Hebrew");
			break;

			case	10006:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Greek I");
			break;

			case	10007:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Cyrillic");
			break;

			case	10008:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Simplified Chinese GB 2312");
			break;

			case	10010:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Romania");
			break;

			case	10017:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Ukraine");
			break;

			case	10029:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Latin II");
			break;

			case	10079:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Icelandic");
			break;

			case	10081:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Turkish");
			break;

			case	10082:	pszBuf = std::format(L"{} {}",CodePage,L"MAC - Croatia");
			break;

			case	20866:	pszBuf = std::format(L"{} {}",CodePage,L"Russian - KOI8");
			break;

			case	29001:	pszBuf = std::format(L"{} {}",CodePage,L"Europa 3");
			break;

			case	65000:	pszBuf = std::format(L"{} {}",CodePage,L"UTF-7");
			break;

			case	65001:	pszBuf = std::format(L"{} {}",CodePage,L"UTF-8");
			break;

			case	20000:	pszBuf = std::format(L"{} {}",CodePage,L"CNS - Taiwan");
			break;

			case	20001:	pszBuf = std::format(L"{} {}",CodePage,L"TCA - Taiwan");
			break;

			case	20002:	pszBuf = std::format(L"{} {}",CodePage,L"Eten - Taiwan");
			break;

			case	20003:	pszBuf = std::format(L"{} {}",CodePage,L"IBM5550 - Taiwan");
			break;

			case	20004:	pszBuf = std::format(L"{} {}",CodePage,L"TeleText - Taiwan");
			break;

			case	20005:	pszBuf = std::format(L"{} {}",CodePage,L"Wang - Taiwan");
			break;

			case	20105:	pszBuf = std::format(L"{} {}",CodePage,L"IA5 IRV International Alphabet No. 5");
			break;

			case	20106:	pszBuf = std::format(L"{} {}",CodePage,L"IA5 German");
			break;

			case	20107:	pszBuf = std::format(L"{} {}",CodePage,L"IA5 Swedish");
			break;

			case	20108:	pszBuf = std::format(L"{} {}",CodePage,L"IA5 Norwegian");
			break;

			case	20127:	pszBuf = std::format(L"{} {}",CodePage,L"US-ASCII");
			break;

			case	20261:	pszBuf = std::format(L"{} {}",CodePage,L"T.61");
			break;

			case	20269:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 6937 Non-Spacing Accent");
			break;

			case	20273:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Germany");
			break;

			case	20277:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Denmark/Norway");
			break;

			case	20278:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Finland/Sweden");
			break;

			case	20280:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Italy");
			break;

			case	20284:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Latin America/Spain");
			break;

			case	20285:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - United Kingdom");
			break;

			case	20290:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Japanese Katakana Extended");
			break;

			case	20297:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - France");
			break;

			case	20420:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Arabic");
			break;

			case	20423:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Greek");
			break;

			case	20424:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Hebrew");
			break;

			case	20833:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Korean Extended");
			break;

			case	20838:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Thai");
			break;

			case	20871:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Icelandic");
			break;

			case	20880:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Cyrillic,(Russian)");
			break;

			case	20905:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Turkish");
			break;

			case	21025:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Cyrillic (Serbian, Bulgarian)");
			break;

			case	21027:	pszBuf = std::format(L"{} {}",CodePage,L"Ext Alpha Lowercase");
			break;

			case	28591:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-1 Latin I");
			break;

			case	28592:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-2 Central Europe");
			break;

			case	28593:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-3 Turkish");
			break;

			case	28594:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-4 Baltic");
			break;

			case	28595:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-5 Cyrillic");
			break;

			case	28596:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-6 Arabic");
			break;

			case	28597:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-7 Greek");
			break;

			case	28598:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-8 Hebrew");
			break;

			case	28599:	pszBuf = std::format(L"{} {}",CodePage,L"ISO 8859-9 Latin 5");
			break;

			case	50220:	pszBuf = std::format(L"{} {}",CodePage,L"ISO-2022 Japanese with no halfwidth Katakana");
			break;

			case	50221:	pszBuf = std::format(L"{} {}",CodePage,L"ISO-2022 Japanese with halfwidth Katakana");
			break;

			case	50222:	pszBuf = std::format(L"{} {}",CodePage,L"ISO-2022 Japanese JIS X 0201-1989");
			break;

			case	50225:	pszBuf = std::format(L"{} {}",CodePage,L"ISO-2022 Korean");
			break;

			case	50227:	pszBuf = std::format(L"{} {}",CodePage,L"ISO-2022 Simplified Chinese");
			break;

			case	50229:	pszBuf = std::format(L"{} {}",CodePage,L"ISO-2022 Traditional Chinese");
			break;

			case	52936:	pszBuf = std::format(L"{} {}",CodePage,L"HZ-GB2312 Simplified Chinese");
			break;

			case	51932:	pszBuf = std::format(L"{} {}",CodePage,L"EUC-Japanese");
			break;

			case	51949:	pszBuf = std::format(L"{} {}",CodePage,L"EUC-Korean");
			break;

			case	51936:	pszBuf = std::format(L"{} {}",CodePage,L"EUC-Simplified Chinese");
			break;

			case	51950:	pszBuf = std::format(L"{} {}",CodePage,L"EUC-Traditional Chinese");
			break;

			case	50930:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Japanese (Katakana) Extended and Japanese");
			break;

			case	50931:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - US/Canada and Japanese");
			break;

			case	50933:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Korean Extended and Korean");
			break;

			case	50935:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Simplified Chinese");
			break;

			case	50937:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - US/Canada and Traditional Chinese");
			break;

			case	50939:	pszBuf = std::format(L"{} {}",CodePage,L"IBM EBCDIC - Japanese (Latin) Extended and Japanese");
			break;

			default:	pszBuf = std::format(L"Unknown {} ", CodePage);
			break;
		}

		return true;
	}

	size_t GetStringLength()
	{
		return m_pszCodePage.length();
	}

	virtual ~TCodePage()
	{
		m_wCodePage = 0;
		m_pszCodePage.clear();
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
	std::wstring m_pszCodePage;
};
