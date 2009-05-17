////////////////////////////////////////////////////////////////////////////////
//
// Copyright (c) 2004 Sergey Solozhentsev
// Author: 	Sergey Solozhentsev e-mail: salos@mail.ru
// Product:	WTL Helper
// File:      	Connect.h
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

// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols
#include "VSElements.h"
#include "common.h"
#include <atlcoll.h>
#include "ResourceManager.h"

#if defined(_FOR_VS2008)

using namespace Microsoft_VisualStudio_CommandBars;
typedef CComPtr<CommandBar> UICommandBarPtr;
typedef CComPtr<CommandBarControl> UICommandBarControlPtr;
typedef CComPtr<_CommandBars> UICommandBarsPtr;
typedef CComPtr<CommandBarPopup> UICommandBarPopupControlPtr;
typedef CComPtr<CommandBarControls> UICommandBarControlsPtr;
typedef CommandBarControl UICommandBarControl;

#elif defined(_FOR_VS2005)

using namespace Microsoft_VisualStudio_CommandBars;
typedef CComPtr<CommandBar> UICommandBarPtr;
typedef CComPtr<CommandBarControl> UICommandBarControlPtr;
typedef CComPtr<_CommandBars> UICommandBarsPtr;
typedef CComPtr<CommandBarPopup> UICommandBarPopupControlPtr;
typedef CComPtr<CommandBarControls> UICommandBarControlsPtr;
typedef CommandBarControl UICommandBarControl;

#else

using namespace Office;
typedef CComPtr<CommandBar> UICommandBarPtr;
typedef CComPtr<CommandBarControl> UICommandBarControlPtr;
typedef CComPtr<_CommandBars> UICommandBarsPtr;
typedef CComPtr<CommandBarPopup> UICommandBarPopupControlPtr;
typedef CComPtr<CommandBarControls> UICommandBarControlsPtr;
typedef CommandBarControl UICommandBarControl;

#endif

enum eWizardsErrors
{
	eSuccess = 0,
	eWizardsNotInstalled = 1,
	eWTLDLGNotInstalled = 2,
	eWrongVersion = 3,
	eError = -1
};


#define CMDBAR_TOOLS		0
#define CMDBAR_CLASSVIEW	1
#define CMDBAR_RESEDITORS	2
#define CMDBAR_MENUEDITORS	3
#define CMDBAR_RESVIEW		4
#define CMDBAR_NO			-1


struct WtlHelperCmdBar
{
	LPCWSTR lpName;
	int nPos;
	EnvDTE::vsCommandBarType CmdType;
	UICommandBarPtr  pCmdBar;
};

struct CommandStruct
{
	_bstr_t Name;
	_bstr_t ButtonText;
	_bstr_t ToolTip;
	_variant_t Bindings;
	VARIANT_BOOL bMSOButton;
	long lBitmapId;
	int iCmdBar;
	long lPos;
	CommandStruct(LPCWSTR lpName, LPCWSTR lpBText, LPCWSTR lpToolTip, LPCWSTR lpBindings, 
		int cmdBar = CMDBAR_NO, long Pos = -1, bool bMSO = true, long lBitmap = 59)
		: Name(lpName), ButtonText(lpBText), ToolTip(lpToolTip), Bindings(lpBindings), 
		iCmdBar(cmdBar), lPos(Pos),
		bMSOButton(bMSO ? VARIANT_TRUE : VARIANT_FALSE), lBitmapId(lBitmap)
	{
	}
};

// CConnect
class ATL_NO_VTABLE CConnect : 
	public CComObjectRootEx<CComSingleThreadModel>,
#if defined(_FOR_VS2008)
	public CComCoClass<CConnect, &CLSID_Connect9>,
#elif defined(_FOR_VS2005)
	public CComCoClass<CConnect, &CLSID_Connect8>,
#else
	public CComCoClass<CConnect, &CLSID_Connect>,
#endif

	public IDispatchImpl<EnvDTE::IDTCommandTarget, &EnvDTE::IID_IDTCommandTarget, &EnvDTE::LIBID_EnvDTE, 7, 0>,
	public IDispatchImpl<AddInDesignerObjects::_IDTExtensibility2, &AddInDesignerObjects::IID__IDTExtensibility2, &AddInDesignerObjects::LIBID_AddInDesignerObjects, 1, 0>
{
private:
	UICommandBarPtr m_pAddinCommandBar;
	UICommandBarPtr m_pClassViewAddinBar;
	UICommandBarPtr m_pResViewAddinBar;
	UICommandBarPtr m_pMenuDesignerBar;
	UICommandBarPtr m_pResViewContextBar;
public:
	CConnect()
	{
	}

	virtual ~CConnect()
	{}

public:
	static HRESULT WINAPI UpdateRegistry(BOOL bRegister)
	{
		OLECHAR szPath[MAX_PATH] = { 0 };
		GetModuleFileNameW(_AtlModule.GetResourceInstance(), szPath, _countof(szPath));
		PathRemoveFileSpecW(szPath); 
		OLECHAR szShortName[MAX_PATH] = { 0 };
		GetShortPathNameW(szPath, szShortName, _countof(szShortName));
		_ATL_REGMAP_ENTRY rm[] = { 
			{ OLESTR("ModulePath"), szShortName },
			{ NULL,NULL },
		};
#if defined(_FOR_VS2008)
		return _AtlModule.UpdateRegistryFromResource(IDR_ADDIN9, bRegister, rm);
#elif defined(_FOR_VS2005)
		return _AtlModule.UpdateRegistryFromResource(IDR_ADDIN8, bRegister, rm);
#else
		return _AtlModule.UpdateRegistryFromResource(IDR_ADDIN, bRegister, rm);
#endif
	}
	// DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)

	DECLARE_NOT_AGGREGATABLE(CConnect)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY(AddInDesignerObjects::IDTExtensibility2)
		COM_INTERFACE_ENTRY(EnvDTE::IDTCommandTarget)
		COM_INTERFACE_ENTRY2(IDispatch, AddInDesignerObjects::IDTExtensibility2)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}
	HRESULT GetActiveProject(CComPtr<EnvDTE::Project> & pProj);
	HRESULT GetSelectedProject(CComPtr<EnvDTE::Project>& pProj);
	HRESULT GetClasses(EnvDTE::Project * pProj, CSmartAtlArray<VSClass*>& Classes);
	HRESULT AddClassToVector(CComPtr<EnvDTE::CodeElement> pItem, CSmartAtlArray<VSClass*>& Classes, VSClass* pParent = NULL);
	HRESULT AddNamespaceToVector(CComPtr<EnvDTE::CodeElement> pItem, CSmartAtlArray<VSClass*>& Classes);
	CString  GetActiveClass(CComPtr<EnvDTE::TextPoint>& pCurPoint);
	int GetActiveClass(CSmartAtlArray<VSClass*>& Classes, CString ClassName);
	void GetBases(CComPtr<EnvDTE::CodeElement> pElem, CAtlArray<CString>& Bases);
	void GetAllBases(VSClass* pClass, CAtlArray<CString>& Bases);
	bool ShowWTLHelper(CComPtr<EnvDTE::Project> pProj, CString ActiveClass, int ActivePage);
	CComPtr<EnvDTE::CodeClass> GetSelectedClass(CComPtr<EnvDTE::Project> & pProj);
	// return resource type of current resource window or empty for no resource windows
	CString GetActiveResourceType(CString& ResourceID, CString& FileName);
	bool IsPossibleResourceType(CString ResourceType);
	bool IsDialogSelected();
	bool IsMultipleSelected();
	bool GetActiveControlID(CAtlArray<CString>& ActiveControls);
	int FindClassByDlgID(ClassVector Classes, CString DialogID);
	void SaveResourceDocuments();
	HRESULT AddDialogHandler(CString RCFile, CString DialogID);
	HRESULT AddIDHandler(CString RCFile, CString ResID, CString ResType);
	CComPtr<EnvDTE::CodeElement> GetClassFromMember(CComPtr<EnvDTE::CodeElement> pMember);
	EnvDTE::wizardResult CreateDialogClass(CComPtr<EnvDTE::Project> pProj, CString DialogID, bool bShow = false);
	eWizardsErrors GetWTLDLGWizardPath(CString& Path, LPCTSTR lpMinVersion);
	// return S_FALSE if showed main WTL Helper dialog, S_OK if it is need to continue or E_FAIL if any error
	HRESULT PrepareDlgClass(CComPtr<EnvDTE::Project> & pProj, CSmartAtlArray<VSClass*>& Classes, int& iDlgClass, CAtlArray<CString>& Controls, CResourceManager* pResManager);
public:
	//IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * Application, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );
	
	//IDTCommandTarget implementation:
	STDMETHOD(QueryStatus)(BSTR CmdName, EnvDTE::vsCommandStatusTextWanted NeededText, EnvDTE::vsCommandStatus *StatusOption, VARIANT *CommandText);
	STDMETHOD(Exec)(BSTR CmdName, EnvDTE::vsCommandExecOption ExecuteOption, VARIANT *VariantIn, VARIANT *VariantOut, VARIANT_BOOL *Handled);
	STDMETHOD(DoAddin)();
	STDMETHOD(DoAddFunction)();
	STDMETHOD(DoAddVariable)();
	STDMETHOD(DoOptions)();
	STDMETHOD(DoClassViewHandler)();
	STDMETHOD(DoClassViewVariable)();
	STDMETHOD(DoResViewDDX)();
	STDMETHOD(DoResViewHandler)();
	STDMETHOD(DoResViewReflect)();
	STDMETHOD(DoResViewDialog)();
	STDMETHOD(DoResViewContextDDX)();
	STDMETHOD(DoResViewContextHandler)();
	STDMETHOD(DoResViewContextDialog)();
	STDMETHOD(CreateCommand)(EnvDTE::Commands* pCommands, const CommandStruct* pCmd);

#if defined(_FOR_VS2008)
	STDMETHOD(UninstallAddin)();
#elif defined(_FOR_VS2005)
	STDMETHOD(UninstallAddin)();
#else
#endif

	CComPtr<EnvDTE::_DTE> m_pDTE;
	CComPtr<EnvDTE::AddIn> m_pAddInInstance;
	EnvDTE::CodeModel* m_pCodeModel;
private:
	
};

#if defined(_FOR_VS2008)
OBJECT_ENTRY_AUTO(__uuidof(Connect9), CConnect)
#elif defined(_FOR_VS2005)
OBJECT_ENTRY_AUTO(__uuidof(Connect8), CConnect)
#else
OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
#endif
