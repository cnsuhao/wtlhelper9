// AddIn.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"

#if defined(_FOR_VS2008)
#include "AddIn9.h"
#elif defined(_FOR_VS2005)
#include "AddIn8.h"
#else
#include "AddIn.h"
#endif

#include "RegArchive.h"
#include "TextFile.h"
#include "dialog/FunctionPage.h"
#include "XMLSettingsArchive.h"
#include "Dialog/DDXVariable.h"
#include "Dialog/ReflectionDlg.h"

CAddInModule _AtlModule;
extern LPCTSTR g_MainDlgName;

void UnregisterCommands();
BOOL RegisterXmlFiles(void);
void UnRegisterXmlFiles(void);

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	_AtlModule.SetResourceInstance(hInstance);
	return _AtlModule.DllMain(dwReason, lpReserved); 
}


// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
	return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
	RegisterXmlFiles();
	// registers object, typelib and all interfaces in typelib
	HRESULT hr = _AtlModule.DllRegisterServer();
	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	UnregisterCommands();
	UnRegisterXmlFiles();

	HRESULT hr = _AtlModule.DllUnregisterServer();
	return hr;
}

void UnregisterCommands()
{
#if defined(_FOR_VS2008)
	CString strVSKey = _T("Software\\Microsoft\\VisualStudio\\9.0\\Setup\\VS");
	CString strCmd = _T(" /command WtlHelper.Connect9.Uninstall");
#elif defined(_FOR_VS2005)
	CString strVSKey = _T("Software\\Microsoft\\VisualStudio\\8.0\\Setup\\VS");
	CString strCmd = _T(" /command WtlHelper.Connect8.Uninstall");
#else
	CString strVSKey = _T("Software\\Microsoft\\VisualStudio\\7.1\\Setup\\VS");
	CString strCmd = _T(" /setup");
#endif

	CString strParam = _T("EnvironmentPath");
	CRegKey RegKey;
	LONG Res;
	//MessageBox(NULL, strVSKey, strCmd, 0);
	Res = RegKey.Open(HKEY_LOCAL_MACHINE, strVSKey, KEY_READ);
	if (Res == ERROR_SUCCESS)
	{
		CString strPath;
		ULONG nChars = FILENAME_MAX;
		Res = RegKey.QueryStringValue(strParam, strPath.GetBuffer(FILENAME_MAX), &nChars);
		strPath.ReleaseBuffer();
		if (!strPath.IsEmpty())
		{
			//MessageBox(NULL, strPath, strCmd, 0);
			SHELLEXECUTEINFO sei;
			ZeroMemory(&sei, sizeof(sei));
			sei.cbSize = sizeof(sei);
			sei.lpFile = strPath.LockBuffer();
			sei.lpParameters = strCmd.LockBuffer();
			sei.nShow = SW_MINIMIZE;
			sei.fMask = SEE_MASK_NOCLOSEPROCESS;

			BOOL bRes = ShellExecuteEx(&sei);

			strPath.UnlockBuffer();
			strCmd.UnlockBuffer();

			if (bRes && sei.hProcess)
			{
				WaitForSingleObject(sei.hProcess, INFINITE);
				CloseHandle(sei.hProcess);
			}
		}
		
	}
}

BOOL RegisterXmlFiles(void)
{
	BOOL bResult = FALSE;
	TCHAR szWorkPath[MAX_PATH] = { 0 };
	GetModuleFileName(_AtlModule.GetResourceInstance(), szWorkPath, _countof(szWorkPath));
	PathRemoveFileSpec(szWorkPath);

	TCHAR szCommonFile[MAX_PATH] = { 0 };
	lstrcpyn(szCommonFile, szWorkPath, _countof(szCommonFile));
	PathAppend(szCommonFile, _T("common.xml")); 

	TCHAR szMessagesFile[MAX_PATH] = { 0 };
	lstrcpyn(szMessagesFile, szWorkPath, _countof(szMessagesFile));
	PathAppend(szMessagesFile, _T("Messages.xml")); 

	if (PathFileExists(szCommonFile) && PathFileExists(szMessagesFile))
	{
		CRegKey key;
		if (ERROR_SUCCESS == key.Create(HKEY_LOCAL_MACHINE, STR_WTL_HELPER_KEY _T("\\common")))
		{
			key.SetValue(szCommonFile, _T("CommonFile"));
			key.SetValue(szMessagesFile, _T("RegMessages"));
			bResult = TRUE;
		}
	}
	return bResult;
}

DWORD RegDeleteKeyRecursive(HKEY RootKey, LPCTSTR pSubKey); 
void UnRegisterXmlFiles(void)
{
	RegDeleteKeyRecursive(HKEY_LOCAL_MACHINE, STR_WTL_HELPER_KEY);
}

DWORD RegDeleteKeyRecursive(HKEY hStartKey , LPCTSTR pKeyName )
{
#define MAX_KEY_LENGTH 256 

	DWORD   dwRtn = 0, dwSubKeyLength = 0;
	LPTSTR  pSubKey = NULL;
	TCHAR   szSubKey[MAX_KEY_LENGTH] = { 0 }; // (256) this should be dynamic.
	HKEY    hKey = NULL;

	// Do not allow NULL or empty key name 
	if ( pKeyName &&  lstrlen(pKeyName))
	{
		if( (dwRtn=RegOpenKeyEx(hStartKey,pKeyName,
			0, KEY_ENUMERATE_SUB_KEYS | DELETE, &hKey )) == ERROR_SUCCESS)
		{
			while (dwRtn == ERROR_SUCCESS )
			{
				dwSubKeyLength = MAX_KEY_LENGTH;
				dwRtn=RegEnumKeyEx(
					hKey,
					0,       // always index zero
					szSubKey,
					&dwSubKeyLength,
					NULL,
					NULL,
					NULL,
					NULL
					);

				if(dwRtn == ERROR_NO_MORE_ITEMS)
				{
					dwRtn = RegDeleteKey(hStartKey, pKeyName);
					break;
				}
				else if(dwRtn == ERROR_SUCCESS)
				{
					dwRtn=RegDeleteKeyRecursive(hKey, szSubKey);
				}
			}
			RegCloseKey(hKey);
			// Do not save return code because error has already occurred
		}
	}
	else
	{
		dwRtn = ERROR_BADKEY;
	}

	return dwRtn;
}



bool CAddInModule::IsMacro(CString MacroName)
{
	for (size_t i = 0; i < m_MacroMasks.GetCount(); i++)
	{
		if (IsMatchMask(MacroName, m_MacroMasks[i]))
			return true;
	}
	return false;
}

bool CAddInModule::IsMatchMask(LPCTSTR Name, LPCTSTR Mask)
{
	if (!Mask)
		return true;
	if (!Name)
	{
		if (Mask && Mask[0] && lstrcmp(Mask, _T("*")))
			return false;
		else
		{
			return true;
		}
	}

	while(*Name && *Mask)
	{
		switch(*Mask) 
		{
		case _T('?'):
			{
				Name++;
				Mask++;
			}
			break;
		case _T('*'):
			{
				Mask++;
				if (!*Mask)
					return true;
				do 
				{
					if (IsMatchMask(Name, Mask))
						return true;
					Name++;
				} while(*Name);
				return false;
			}
			break;
		default:
			{
				if (*Name != *Mask)
					return false;
				Name++;
				Mask++;
			}
		}
	}
	return ((!*Name && !*Mask) || (*Mask == _T('*')));
}

bool CAddInModule::LoadAll()
{
	CRegArchive RegArchive;
	if (RegArchive.Open(HKEY_CURRENT_USER, STR_WTL_HELPER_KEY))
	{
		LoadSettings(RegArchive, g_MainDlgName, 2);
		RegArchive.Close();
	}
	if (!m_DialogClasses.IsEmpty() || !m_MacroMasks.IsEmpty())
		return true;

	BOOL bFind = FALSE;
	if (RegArchive.Open(HKEY_LOCAL_MACHINE, STR_WTL_HELPER_KEY)) {
		bFind = TRUE;
	} else {
		if ( RegisterXmlFiles() ) {
			bFind = RegArchive.Open(HKEY_LOCAL_MACHINE, STR_WTL_HELPER_KEY);
		}
	}

	if (FALSE == bFind)
	{
#define OUT_MSG _T("Cannot open Load files Common.xml and messages.xml!!!") 
		MessageBox(::GetActiveWindow(), OUT_MSG, _T("Error"), MB_OK|MB_ICONWARNING);
		return false;
	}

	if (_AtlModule.LoadSettings(RegArchive, _T("Common")))
	{
		bool bTrue = true;
		CXMLSettingsArchive XMLArchive;
		if (XMLArchive.Open(m_CommonFile))
		{
			if (!LoadSettings(XMLArchive, _T("Common"), 1))
				bTrue = false;
			XMLArchive.StartObject(_T("Common"));
			if (!CFunctionManager::LoadSpecFunctions(&XMLArchive, _T("SpecFunctions")))
				bTrue = false;
			if (!CAtlArrayObjectSerializer<DDXContols>::LoadVariable(_T("ControlClasses"), XMLArchive, CDDXVariable::m_ControlClasses))
				bTrue = false;
			if (!CAtlArrayObjectSerializer<DDXMemberTypes>::LoadVariable(_T("MemberTypes"), XMLArchive, CDDXVariable::m_MemberTypes))
				bTrue = false;
			if (!CAtlArraySerializer<CString>::LoadVariable(_T("ReflectMessages"), XMLArchive, CReflectionDlg::m_sReflectMessages))
				bTrue = false;

			XMLArchive.EndObject();

			XMLArchive.Close();
		}
		else
			bTrue = false;
		if (XMLArchive.Open(m_RegMessages))
		{
			if (!CMessageManager::m_sHandlerManager.LoadSettings(&XMLArchive, _T("HandlerManager")))
				bTrue = false;
			XMLArchive.Close();
		}
		else
			bTrue = false;

		if (!bTrue)
		{
			MessageBox(NULL, _T("Cannot load some settings"), NULL, 0);
		}
	}
	RegArchive.Close();

	return true;
}

void CAddInModule::Destroy()
{
	m_DialogClasses.RemoveAll();
	m_MacroMasks.RemoveAll();
}
