////////////////////////////////////////////////////////////////////////////////
//
// Copyright (c) 2004 Sergey Solozhentsev
// Author: 	Sergey Solozhentsev e-mail: salos@mail.ru
// Product:	WTL Helper
// File:      	stdafx.h
// Created:	16.11.2004 8:55
// 
//   Using this software in commercial applications requires an author
// permission. The permission will be granted to everyone excluding the cases
// when someone simply tries to resell the code.
//   This file may be redistributed by any means PROVIDING it is not sold for
// profit without the authors written consent, and providing that this notice
// and the authors name is included.
//   This file is provided "as is" with no expressed or implied warranty. The
// author accepts no liability if it causes any damage to you or your computer
// whatsoever.
//
////////////////////////////////////////////////////////////////////////////////

#pragma once

#ifndef STRICT
#define STRICT
#endif

// Modify the following defines if you have to target a platform prior to the ones specified below.
// Refer to MSDN for the latest info on corresponding values for different platforms.
#ifndef WINVER				// Allow use of features specific to Windows 95 and Windows NT 4 or later.
#define WINVER 0x0500		// Change this to the appropriate value to target Windows 98 and Windows 2000 or later.
#endif

#ifndef _WIN32_WINNT		// Allow use of features specific to Windows NT 4 or later.
#define _WIN32_WINNT 0x0501	// Change this to the appropriate value to target Windows 2000 or later.
#endif						

#ifndef _WIN32_IE			// Allow use of features specific to IE 4.0 or later.
#define _WIN32_IE 0x0500	// Change this to the appropriate value to target IE 5.0 or later.
#endif

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

// turns off ATL's hiding of some common and often safely ignored warning messages
#define _ATL_ALL_WARNINGS
#define ISOLATION_AWARE_ENABLED 1

#define  _WTL_NO_CSTRING
#include <atlstr.h>
#include <atlbase.h>
#include <atlcom.h>
#include <atlsafe.h>
#include <atlcoll.h>
#include <atlapp.h>
#include <atlwin.h>
#include <atlctrls.h>
#include <atlmisc.h>
#include <atlddx.h>

using namespace ATL;

#include "Settings.h"
#include "CollSerialize.h"
#include <windows.h>
#include <objbase.h>

#include "imports.h"

//wParam - BOOL is modified
#define WTLH_SETMODIFIED	(WM_USER + 500)
//wParam = HWND of the toolbar
#define WTLH_SETACTIVE		(WM_USER + 501)


#define FIRST_TOOLBAR_COMMAND	20000



#define IfFailGo(x) { hr=(x); if (FAILED(hr)) goto Error; }
#define IfFailGoCheck(x, p) { hr=(x); if (FAILED(hr)) goto Error; if(p == NULL) {hr = E_FAIL; goto Error; } }

#if defined(_FOR_VS2008)
class DECLSPEC_UUID("D193ECB0-D5E8-483a-922E-E0F5AD020262") WtlHelperLib;
#elif defined(_FOR_VS2005)
class DECLSPEC_UUID("E2CBD4CB-D05B-480f-B9A3-76F31CB0973E") WtlHelperLib;
#else
class DECLSPEC_UUID("730B33D4-7AEB-476B-9018-703AB582E457") WtlHelperLib;
#endif

enum eWTLVersion
{
	eWTL71 = 0,
	eWTL75 = 1,
	eWTL80 = 2,
	eWTL81 = 3,
};

class CAddInModule : public CAtlDllModuleT< CAddInModule >,
	public CSettings<CAddInModule>
{
public:
	CAddInModule() : m_eWTLVersion(eWTL75)
	{
		m_hInstance = NULL;
	}

	DECLARE_LIBID(__uuidof(WtlHelperLib))

	inline HINSTANCE GetResourceInstance()
	{
		return m_hInstance;
	}

	inline void SetResourceInstance(HINSTANCE hInstance)
	{
		m_hInstance = hInstance;
	}
	CAtlArray<CString> m_DialogClasses;

	bool IsMacro(CString MacroName);
	bool LoadAll();
	void Destroy();

	CString m_RegMessages;
	CString m_CommonFile;
	eWTLVersion m_eWTLVersion;

	BEGIN_SETTINGS_MAP()
		SETTINGS_VARIABLE_RO(m_RegMessages)
		SETTINGS_VARIABLE_RO(m_CommonFile)
	ALT_SETTINGS_MAP(1)
		SETTINGS_OBJECT_RO(m_DialogClasses, CAtlArraySerializer<CString>)
		SETTINGS_OBJECT_RO(m_MacroMasks, CAtlArraySerializer<CString>)
	ALT_SETTINGS_MAP(2)
		SETTINGS_ENUM_OPT_RO(m_eWTLVersion)
	END_SETTINGS_MAP()
	
private:
	HINSTANCE m_hInstance;
	CAtlArray<CString> m_MacroMasks;
	bool IsMatchMask(LPCTSTR Name, LPCTSTR Mask);
};

template<class T>
class CSmartAtlArray : public CAtlArray<T>
{
public:
	CSmartAtlArray(const CSmartAtlArray<T>& ar)
	{
		Copy(ar);
	}
	CSmartAtlArray() : CAtlArray<T>(){};
	CSmartAtlArray& operator =(const CSmartAtlArray<T> ar)
	{
		Copy(ar);
		return *this;
	}
};
extern CAddInModule _AtlModule;

#ifdef _DEBUG
//profile macros
#define START_PROFILE(shortname) \
	LARGE_INTEGER start_##shortname, end_##shortname;\
	QueryPerformanceCounter(&start_##shortname);

#define END_PROFILE(shortname, str)\
	QueryPerformanceCounter(&end_##shortname);\
	__int64 res_##shortname = *(__int64*)&(end_##shortname) - *(__int64*)&(start_##shortname);\
	LARGE_INTEGER freq_##shortname; QueryPerformanceFrequency(&freq_##shortname);\
	__int64 ifreq_##shortname = *(__int64*)&(freq_##shortname);\
	double millisec_##shortname = (double)(res_##shortname * 1000) / (double)ifreq_##shortname;\
	CString str_##shortname; str_##shortname.Format(_T("\r\n%s - %f ms\r\n\r\n"), str, millisec_##shortname);\
	OutputDebugString(str_##shortname);
#else

#define	START_PROFILE(shortname)
#define END_PROFILE(str, shortname)

#endif

#ifndef _countof
#define _countof(_Array) sizeof(_Array) / sizeof((_Array)[0])
#endif

#define STR_WTL_WIZARDS     _T("Software\\SaloS\\WTL Wizards")
#define STR_WTL_HELPER_KEY  _T("Software\\SaloS\\WtlHelper")

#define STR_WTL_HELPER_W    L"WTL Helper"
#define STR_WTL_HELPER_A    "WTL Helper"
#ifdef UNICODE
#define STR_WTL_HELPER      STR_WTL_HELPER_W
#else
#define STR_WTL_HELPER      STR_WTL_HELPER_A
#endif

