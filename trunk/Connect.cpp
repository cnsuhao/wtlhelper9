////////////////////////////////////////////////////////////////////////////////
//
// Copyright (c) 2004 Sergey Solozhentsev
// Author: 	Sergey Solozhentsev e-mail: salos@mail.ru
// Product:	WTL Helper
// File:      	Connect.cpp
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

// Connect.cpp : Implementation of CConnect
#include "stdafx.h"

#if defined(_FOR_VS2008)
#include "AddIn9.h"
#elif defined(_FOR_VS2005)
#include "AddIn8.h"
#else
#include "AddIn.h"
#endif

#include "Connect.h"
#include "dialog/FunctionPage.h"
#include "dialog/GeneralPage.h"
#include "dialog/VariablePage.h"
#include "dialog/AddMemberVarGL.h"
#include "Options/OptionsDlg.h"
#include "Dialog/WtlHelperDlg.h"
#include "common.h"
#include <msiquery.h>
#include <msi.h>
#include "WtlHelperRes/resource.h"
#include "CustomProjectSettings.h"

extern CAddInModule _AtlModule;
VSFunction* g_pSelectedFunction;

#if defined(_FOR_VS2008)
#define AddinString() L"WtlHelper.Connect9."
#elif defined(_FOR_VS2005)
#define AddinString() L"WtlHelper.Connect8."
#else
#define AddinString() L"WtlHelper.Connect."
#endif

#define CMD(Command) AddinString()L#Command

WtlHelperCmdBar WtlHelperCmdBars[]=
{
	{L"Tools", 1, EnvDTE::vsCommandBarTypeMenu, },
	{L"Class View Item", 10, EnvDTE::vsCommandBarTypeMenu, },
	{L"Resource Editors", 1, EnvDTE::vsCommandBarTypeMenu, },
	{L"Menu Designer", 1, EnvDTE::vsCommandBarTypeMenu, },
	{L"Resource View", 1, EnvDTE::vsCommandBarTypeMenu, }
};

const CommandStruct VariableCmd(L"AddVariable", L"Add Variable", L"Add variable to class", 
								L"Text Editor::Ctrl+Shift+W, V");
const CommandStruct FunctionCmd(L"AddFunction", L"Add Function", L"Add function to class", 
								L"Text Editor::Ctrl+Shift+W, F");

#if defined(_FOR_VS2008)
const CommandStruct UninstallCmd(L"Uninstall", L"Uninstall", NULL, L"", CMDBAR_NO, -1,
								 false, -1);
#elif defined(_FOR_VS2005)
const CommandStruct UninstallCmd(L"Uninstall", L"Uninstall", NULL, L"", CMDBAR_NO, -1,
								 false, -1);
#else
#endif


// WTL Helper command

const CommandStruct HelperCmd(L"WtlHelper", L"WTL Helper", L"Helps to add message handlers",
							  L"Text Editor::Ctrl+Alt+W", CMDBAR_TOOLS, 1, false, IDB_BITMAP_WIZARD);

// Options command
const CommandStruct OptionCmd(L"Options", L"Options", L"Set options for WTL Helper", 
							  L"Text Editor::Ctrl+Shift+W, O", CMDBAR_TOOLS, 2, false, IDB_BITMAP_OPTION);

//Class view context menu
const CommandStruct ClassViewHandlerCmd(L"ClassView_Handlers", L"Add handler or function",
										L"Show WTL Helper main window on Functions page", 
										L"", CMDBAR_CLASSVIEW, 1, false, IDB_BITMAP_MSG);
const CommandStruct ClassViewVarCmd(L"ClassView_Variables", L"Add DDX variable", 
									L"Show WTL Helper main window on Variables page", L"", 
									CMDBAR_CLASSVIEW, 2, false, IDB_BITMAP_DDXVAR);

// Resource view context menu
const CommandStruct ResViewDDXCmd(L"ResView_DDX", L"Add DDX entry", 
								  L"Add entry to DDX map for the specified control", L"", 
								  CMDBAR_RESEDITORS, 1, false, IDB_BITMAP_DDXVAR);
const CommandStruct ResViewDDXMultCmd(L"ResView_DDX_Mult", L"Add multiple DDX entries", 
									  L"Add multiple entries to DDX map for the specified controls", 
									  L"", CMDBAR_RESEDITORS, 2, false, IDB_BITMAP_DDXVAR);
const CommandStruct ResViewHandlerCmd(L"ResView_Handler", L"Add handler", 
									  L"Add command or notify handler for the specified item", 
									  L"", CMDBAR_RESEDITORS, 3, false, IDB_BITMAP_MSG);
const CommandStruct ResViewHandlerMultCmd(L"ResView_Handler_Mult", L"Add multiple handlers", 
										  L"Add command or notify handlers for the specified items", 
										  L"", CMDBAR_RESEDITORS, 4, false, IDB_BITMAP_MSG);
const CommandStruct ResViewDialogCmd(L"ResView_Dialog", L"Create dialog class", 
									 L"Create dialog class using WTL Dialog Wizard", L"", 
									 CMDBAR_RESEDITORS, 5, false, IDB_BITMAP_DIALOG);
const CommandStruct ResViewReflectCmd(L"ResView_Reflect", L"Add custom reflection handler", 
									  L"Add custom reflection handler only for WTL 7.5 and above", 
									  L"", CMDBAR_RESEDITORS, 6, false, IDB_BITMAP_REFLECT);

// Resource view for menu editor
const CommandStruct MenuDesignHandlerCmd(L"MenuDesign_Handler", L"Add handler", 
										 L"Add command handler for the specified menu item", 
										 L"", CMDBAR_MENUEDITORS, 1, false, IDB_BITMAP_MSG);

// Resource file view 
const CommandStruct ResViewContextDialogCmd(L"ResViewContext_Dialog", L"Create dialog class", 
											L"Create dialog class using WTL Dialog Wizard", 
											L"", CMDBAR_RESVIEW, 1, false, IDB_BITMAP_DIALOG);
const CommandStruct ResViewContextHandlerCmd(L"ResViewContext_Handler", L"Add handler", 
											 L"Add command or notify handler for the specified item", 
											 L"", CMDBAR_RESVIEW, 2, false, IDB_BITMAP_MSG);
const CommandStruct ResViewContextDDXCmd(L"ResViewContext_DDX", L"Add DDX entry", 
										 L"Add entry to DDX map for the specified control", 
										 L"", CMDBAR_RESVIEW, 3, false, IDB_BITMAP_DDXVAR);

// List of all commands
const CommandStruct* WtlHelperCommands[]=
{
	&VariableCmd,
	&FunctionCmd,

#if defined(_FOR_VS2008)
	&UninstallCmd,
#elif defined(_FOR_VS2005)
	&UninstallCmd,
#else
#endif

	&HelperCmd,
	&OptionCmd,
	&ClassViewHandlerCmd,
	&ClassViewVarCmd,
	&ResViewDDXCmd,
	&ResViewDDXMultCmd,
	&ResViewHandlerCmd,
	&ResViewHandlerMultCmd,
	&ResViewDialogCmd,
	&ResViewReflectCmd,
	&MenuDesignHandlerCmd,
	&ResViewContextDialogCmd,
	&ResViewContextHandlerCmd,
	&ResViewContextDDXCmd
};


// When run, the Add-in wizard prepared the registry for the Add-in.
// At a later time, if the Add-in becomes unavailable for reasons such as:
//   1) You moved this project to a computer other than which is was originally created on.
//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
//   3) Registry corruption.
// you will need to re-register the Add-in by building the MyAddin21Setup project 
// by right clicking the project in the Solution Explorer, then choosing install.

CComPtr<EnvDTE::Project> g_pActiveProject;

// CConnect
STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/ )
{
	HRESULT hr = S_OK;
	pApplication->QueryInterface(__uuidof(EnvDTE::_DTE), (LPVOID*)&m_pDTE);
	pAddInInst->QueryInterface(__uuidof(EnvDTE::AddIn), (LPVOID*)&m_pAddInInstance);
	CCustomProjectSettings::m_sCustomProperties.Init(m_pDTE);

	CComPtr<EnvDTE::Commands> pCommands;
	UICommandBarsPtr pCommandBars;
	CComPtr<EnvDTE::Command> pCreatedCommand;
	UICommandBarPtr pHostCmdBar;

	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, the Add-in or its commands may become unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) You add new commands or modify commands already defined.
	// You will need to re-register the Add-in by building the WtlHelperSetup project,
	// right-clicking the project in the Solution Explorer, and then choosing install.
	// Alternatively, you could execute the ReCreateCommands.reg file the Add-in Wizard generated in 
	// the project directory, or run 'devenv /setup' from a command prompt.
	//if (ConnectMode == AddInDesignerObjects::ext_cm_Startup || ConnectMode == AddInDesignerObjects::ext_cm_AfterStartup)
	IfFailGoCheck(m_pDTE->get_Commands(&pCommands), pCommands);

#if defined(_FOR_VS2008)
	IfFailGoCheck(m_pDTE->get_CommandBars((IDispatch**)&pCommandBars), pCommandBars);
#elif defined(_FOR_VS2005)
	IfFailGoCheck(m_pDTE->get_CommandBars((IDispatch**)&pCommandBars), pCommandBars);
#else
	IfFailGoCheck(m_pDTE->get_CommandBars(&pCommandBars), pCommandBars);
#endif

	if(ConnectMode == 5) //5 == ext_cm_UISetup
	{
		// remove old menus if exists
		for (size_t i = 0; i < _countof(WtlHelperCmdBars); i++)
		{
			pHostCmdBar = NULL;
			if (FAILED(pCommandBars->get_Item(_variant_t(WtlHelperCmdBars[i].lpName), 
				&pHostCmdBar)) || pHostCmdBar == NULL)
			{
				continue;
			}
			
			UICommandBarControlsPtr pControls;
			UICommandBarPtr pHelperCmdBar;
			pHostCmdBar->get_Controls(&pControls);
			UICommandBarControlPtr pControl;
			hr = pControls->get_Item(_variant_t(L"WTL Helper"), &pControl);
			while (pControl != NULL)
			{
				UICommandBarPopupControlPtr pPopup; 
				hr = pControl->QueryInterface(&pPopup);
				if (pPopup != NULL)
				{
					pPopup->get_CommandBar(&pHelperCmdBar);
					if(pHelperCmdBar != NULL)
					{
						hr = pCommands->RemoveCommandBar(pHelperCmdBar);
					}
					else
					{
						pPopup->Delete();
					}
				}

				pControls->get_Item(_variant_t(L"WTL Helper"), &pControl);
			}
		}
	}
	//ATLASSERT(false);

	if(ConnectMode != 5)
	{
		//create new menus
		for (size_t i = 0; i < _countof(WtlHelperCmdBars); i++)
		{
			pHostCmdBar = NULL;
			if (FAILED(pCommandBars->get_Item(_variant_t(WtlHelperCmdBars[i].lpName), 
				&pHostCmdBar)) || pHostCmdBar == NULL)
			{
				continue;
			}

			UICommandBarControlsPtr pBarControls;
			pHostCmdBar->get_Controls(&pBarControls);
			UICommandBarPopupControlPtr pPopup;

			_variant_t EmptyParam;
			EmptyParam.Clear();
			hr = pBarControls->Add(_variant_t(msoControlPopup), _variant_t(1), EmptyParam, 
				_variant_t(WtlHelperCmdBars[i].nPos), _variant_t(true), 
				(UICommandBarControl**)&pPopup);
			if (FAILED(hr))
				break;
			pPopup->put_Caption(_bstr_t(L"WTL Helper"));
			pPopup->put_Visible(VARIANT_TRUE);
			pPopup->get_CommandBar(&WtlHelperCmdBars[i].pCmdBar);
		}

		//create commands
		for (size_t i = 0; i < _countof(WtlHelperCommands); i++)
		{
			CreateCommand(pCommands, WtlHelperCommands[i]);
		}

		CFunctionPage::FillFuncButtons();
	}

	return S_OK;

Error:
	return hr;
}

HRESULT CConnect::CreateCommand(EnvDTE::Commands* pCommands, const CommandStruct* pCmd)
{
	CComPtr<EnvDTE::Command> pCreatedCommand;
	UICommandBarControlPtr pCommandBarControl;

	HRESULT hr = pCommands->AddNamedCommand(m_pAddInInstance, pCmd->Name, pCmd->ButtonText, pCmd->ToolTip, pCmd->bMSOButton, pCmd->lBitmapId, NULL, EnvDTE::vsCommandStatusSupported+EnvDTE::vsCommandStatusEnabled, &pCreatedCommand);
	if(SUCCEEDED(hr) && (pCreatedCommand != NULL))
	{
		if ((pCmd->iCmdBar != CMDBAR_NO) && (WtlHelperCmdBars[pCmd->iCmdBar].pCmdBar != NULL))
		{
			//Add a button to the tools menu bar.
#if defined(_FOR_VS2008)
			if (FAILED(pCreatedCommand->AddControl(WtlHelperCmdBars[pCmd->iCmdBar].pCmdBar,
				pCmd->lPos, (IDispatch**)&pCommandBarControl)) || (pCommandBarControl == NULL))
#elif defined(_FOR_VS2005)
			if (FAILED(pCreatedCommand->AddControl(WtlHelperCmdBars[pCmd->iCmdBar].pCmdBar,
				pCmd->lPos, (IDispatch**)&pCommandBarControl)) || (pCommandBarControl == NULL))
#else
			if (FAILED(pCreatedCommand->AddControl(WtlHelperCmdBars[pCmd->iCmdBar].pCmdBar,
				pCmd->lPos, &pCommandBarControl)) || (pCommandBarControl == NULL))
#endif
			{
				return E_FAIL;
			}
		}

		if (pCmd->Bindings.bstrVal && pCmd->Bindings.bstrVal[0])
		{
			if (pCreatedCommand->put_Bindings(pCmd->Bindings) != S_OK)
			{
				return S_FALSE;
			}
		}
	}
	else
	{
		CComPtr<EnvDTE::Command> pCurCommand = NULL;
		_bstr_t commandname = AddinString();
		commandname += pCmd->Name;
		pCommands->Item(_variant_t(commandname), 0, &pCurCommand);
		if (pCurCommand == NULL)
		{
			return E_FAIL;
		}

		if ((pCmd->iCmdBar != CMDBAR_NO) && (WtlHelperCmdBars[pCmd->iCmdBar].pCmdBar != NULL))
		{
#if defined(_FOR_VS2008)
			if (FAILED(pCurCommand->AddControl(WtlHelperCmdBars[pCmd->iCmdBar].pCmdBar, 
				pCmd->lPos,	(IDispatch**)&pCommandBarControl)) || (pCommandBarControl == NULL))
#elif defined(_FOR_VS2005)
			if (FAILED(pCurCommand->AddControl(WtlHelperCmdBars[pCmd->iCmdBar].pCmdBar, 
				pCmd->lPos,	(IDispatch**)&pCommandBarControl)) || (pCommandBarControl == NULL))
#else
			if (FAILED(pCurCommand->AddControl(WtlHelperCmdBars[pCmd->iCmdBar].pCmdBar, 
				pCmd->lPos, &pCommandBarControl)) || (pCommandBarControl == NULL))
#endif
			{
				return E_FAIL;
			}
		}

		if (pCurCommand->put_Bindings(pCmd->Bindings) != S_OK)
		{
			return S_FALSE;	
		}

	}

	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY ** /*custom*/ )
{
	HRESULT hr = S_OK;

	// AddInDesignerObjects::ext_DisconnectMode == 2when unload after (ConnectMode == 5) //5 == ext_cm_UISetup
	if(RemoveMode != 2)
	{
		for (size_t i = 0; i < _countof(WtlHelperCmdBars); i++)
		{
			WtlHelperCmdBars[i].pCmdBar.Release();
		}

		m_pDTE.Release();
		m_pAddInInstance.Release();
		_AtlModule.Destroy();
	}
	return hr;
}

STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::QueryStatus(BSTR bstrCmdName, EnvDTE::vsCommandStatusTextWanted NeededText, EnvDTE::vsCommandStatus *pStatusOption, VARIANT *pvarCommandText)
{
	if(NeededText == EnvDTE::vsCommandStatusTextWantedNone)
	{
		if(!_wcsicmp(bstrCmdName, CMD(WtlHelper)))
		{
			*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
		}
		else
		{
			int a = 0;
		}
		if(!_wcsicmp(bstrCmdName, CMD(AddFunction)))
		{
			*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
		}
		if(!_wcsicmp(bstrCmdName, CMD(AddVariable)))
		{
			*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
		}
		if(!_wcsicmp(bstrCmdName, CMD(Options)))
		{
			*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
		}
		if (!_wcsicmp(bstrCmdName, CMD(ClassView_Handlers)))
		{
			CComPtr<EnvDTE::Project> pProj;
			if (GetSelectedClass(pProj) != NULL)
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
		if (!_wcsicmp(bstrCmdName, CMD(ClassView_Variables)))
		{
			CComPtr<EnvDTE::Project> pProj;
			if (GetSelectedClass(pProj) != NULL)
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResView_DDX)))
		{
			CString ResourceID;
			CString RCFileName;
			if (GetActiveResourceType(ResourceID, RCFileName) == _T("Dialog") && !IsMultipleSelected())
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResView_DDX_Mult)))
		{
			CString ResourceID;
			CString RCFileName;
			if (GetActiveResourceType(ResourceID, RCFileName) == _T("Dialog") && IsMultipleSelected())
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}

		if (!_wcsicmp(bstrCmdName, CMD(ResView_Handler_Mult)))
		{
			CString ResourceID;
			CString RCFileName;
			if (IsPossibleResourceType(GetActiveResourceType(ResourceID, RCFileName)) && IsMultipleSelected())
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}

		if (!_wcsicmp(bstrCmdName, CMD(ResView_Handler)) ||
			!_wcsicmp(bstrCmdName, CMD(MenuDesign_Handler)))
		{
			CString ResourceID;
			CString RCFileName;
			if (IsPossibleResourceType(GetActiveResourceType(ResourceID, RCFileName)) && !IsMultipleSelected())
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResView_Dialog)))
		{
			CString ResourceID;
			CString RCFileName;
			if (GetActiveResourceType(ResourceID, RCFileName) == _T("Dialog") && IsDialogSelected())
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResViewContext_Handler)))
		{
			if (IsDialogSelected())
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResViewContext_DDX)))
		{
			if (IsDialogSelected())
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResViewContext_Dialog)))
		{
			if (IsDialogSelected())
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResView_Reflect)))
		{
			CString ResourceID;
			CString RCFileName;
			if (GetActiveResourceType(ResourceID, RCFileName) == _T("Dialog") 
				&& !IsMultipleSelected()
				&& _AtlModule.m_eWTLVersion >= eWTL75)
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
			}
			else
			{
				*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusInvisible+EnvDTE::vsCommandStatusSupported);
			}
		}
#if defined(_FOR_VS2008)
		if (!_wcsicmp(bstrCmdName, CMD(Uninstall)))
		{
			*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
		}
#elif defined(_FOR_VS2005)
		if (!_wcsicmp(bstrCmdName, CMD(Uninstall)))
		{
			*pStatusOption = (EnvDTE::vsCommandStatus)(EnvDTE::vsCommandStatusEnabled+EnvDTE::vsCommandStatusSupported);
		}
#else
#endif
	}
	
	return S_OK;
}

STDMETHODIMP CConnect::Exec(BSTR bstrCmdName, EnvDTE::vsCommandExecOption ExecuteOption, VARIANT * /*pvarVariantIn*/, VARIANT * /*pvarVariantOut*/, VARIANT_BOOL *pvbHandled)
{
	*pvbHandled = VARIANT_FALSE;
	if(ExecuteOption == EnvDTE::vsCommandExecOptionDoDefault)
	{
		if(!_wcsicmp(bstrCmdName, CMD(WtlHelper)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoAddin();
		}
		if (!_wcsicmp(bstrCmdName, CMD(AddFunction)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoAddFunction();
		}
		if (!_wcsicmp(bstrCmdName, CMD(AddVariable)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoAddVariable();
		}
		if (!_wcsicmp(bstrCmdName, CMD(Options)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoOptions();
		}
		if (!_wcsicmp(bstrCmdName, CMD(ClassView_Handlers)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoClassViewHandler();
		}
		if (!_wcsicmp(bstrCmdName, CMD(ClassView_Variables)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoClassViewVariable();
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResView_DDX)) ||
			!_wcsicmp(bstrCmdName, CMD(ResView_DDX_Mult)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoResViewDDX();
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResView_Handler)) ||
			!_wcsicmp(bstrCmdName, CMD(ResView_Handler_Mult)) ||
			!_wcsicmp(bstrCmdName, CMD(MenuDesign_Handler)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoResViewHandler();
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResView_Dialog)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoResViewDialog();
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResViewContext_Dialog)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoResViewContextDialog();
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResViewContext_Handler)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoResViewContextHandler();
		}
		if (!_wcsicmp(bstrCmdName, CMD(ResViewContext_DDX)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoResViewContextDDX();
		}

		if (!_wcsicmp(bstrCmdName, CMD(ResView_Reflect)))
		{
			*pvbHandled = VARIANT_TRUE;
			DoResViewReflect();
		}
#if defined(_FOR_VS2008)
		if (!_wcsicmp(bstrCmdName, CMD(Uninstall)))
		{
			UninstallAddin();
		}
#elif defined(_FOR_VS2005)
		if (!_wcsicmp(bstrCmdName, CMD(Uninstall)))
		{
			UninstallAddin();
		}
#else
#endif
	}
	return S_OK;
}

bool CConnect::ShowWTLHelper(CComPtr<EnvDTE::Project> pProj, CString ActiveClass, int ActivePage)
{
	_bstr_t str;
	_AtlModule.LoadAll();

	if (pProj == NULL)
	{
		if (GetActiveProject(pProj) != S_OK)
		{
			return false;
		}
	}
	g_pActiveProject = pProj;

	CSmartAtlArray<VSClass*> Classes;
	if (GetClasses(pProj, Classes) != S_OK)
	{
		return false;
	}
	if (!Classes.GetCount())
	{
		MessageBox(NULL, _T("No classes to work with!"), _T("No classes"), MB_OK | MB_ICONINFORMATION);
		return false;
	}
	SortClasses(Classes);
	size_t iCurClass =  GetActiveClass(Classes, ActiveClass);

	SaveResourceDocuments();

	CSmartAtlArray<InsDelPoints> Modifications;
	CResourceManager ResManager;
	CAtlArray<CString> ResFiles;

	ResManager.LoadResources(pProj);

	// set modifications count
	Modifications.SetCount(Classes.GetCount());

	CWtlHelperDlg dlg;
	dlg.m_pClassVector = &Classes;
	dlg.m_pModifications = &Modifications;
	dlg.m_piCurrentClass = (int*)&iCurClass;
	dlg.m_pResManager = &ResManager;
	dlg.m_iActivePage = ActivePage;

	if (dlg.DoModal() == IDOK)
	{
		UpdateClasses(&Modifications, &Classes);
		pProj->Save(NULL);
	}

	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}
	Classes.RemoveAll();
#ifdef DEBUG
	// MessageBox(NULL, _T("End Work"), NULL, 0);
#endif
	return true;
}

STDMETHODIMP CConnect::DoAddin()
{
	g_pSelectedFunction = NULL;
	CComPtr<EnvDTE::Project> spDummy;
	ShowWTLHelper(spDummy, CString(), -1);
	return S_OK;
}

HRESULT CConnect::DoAddFunction()
{
	CComPtr<EnvDTE::Project> pProj;
	if (GetActiveProject(pProj) != S_OK)
	{
		return S_FALSE;
	}

	CSmartAtlArray<VSClass*> Classes;
	if (GetClasses(pProj, Classes) != S_OK)
	{
		return S_FALSE;
	}

	VSClass* pCurClass = NULL;
	pCurClass = Classes[GetActiveClass(Classes, CString())];

	CAddMemberVarGl dlg;
	dlg.SetForFunction();
	dlg.m_pClassVector = &Classes;
	dlg.m_pCurClass = pCurClass;

	if ((dlg.DoModal() == IDOK) && (dlg.m_pCurClass))
	{
		VSFunction* pFunc = new VSFunction;
		pFunc->Type = dlg.m_Type;
		int n1 = dlg.m_Body.Find(_T('('));
		if (n1 == -1)
		{
			MessageBox(NULL, _T("It's not a function"), NULL, 0);
			return 0;
		}
		pFunc->Name = dlg.m_Body.Left(n1);
		pFunc->Name.Trim();
		if (dlg.m_bPublic)
		{
			pFunc->Access = EnvDTE::vsCMAccessPublic;
		}
		else
		{
			if (dlg.m_bProtected)
			{
				pFunc->Access = EnvDTE::vsCMAccessProtected;
			}
			else
			{
				pFunc->Access = EnvDTE::vsCMAccessPrivate;
			}
		}

		CString Params = dlg.m_Body.Mid(n1+1, dlg.m_Body.GetLength() - n1 - 2);
		int StartPos = 0;
		VSParameter* pParam;
		Params.Trim();
		if (!Params.IsEmpty())
		{
			while(StartPos >=0)
			{
				//определение параметров не совсем правильное
				//может быть ситуация int a b, int c
				CString Param1;
				int n2 = Params.Find(_T(','), StartPos);
				if (n2 != -1)
				{
					Param1 = Params.Mid(StartPos, n2 - StartPos);
					StartPos = n2+1;
				}
				else
				{
					Param1 = Params.Mid(StartPos);
					StartPos = -1;
				}
				Param1.Trim();
				int n3 = -1;
				int i = Param1.GetLength()-1;
				CString Delim = _T(" \t\r\n");

				//находим первый разделитель с конца
				while(i >= 0)
				{
					if (Delim.Find(Param1[i]) != -1)
					{
						n3 = i;
						break;
					}
					i--;
				}

				if (n3 == -1)
				{
					MessageBox(NULL, _T("Wrong parameter list"), NULL, 0);
					return 0;
				}

				pParam = new VSParameter;
				pParam->Name  = Param1.Mid(n3+1);
				pParam->Name.Trim();
				pParam->Type = Param1.Left(n3);
				pParam->Type.Trim();
				pFunc->Parameters.Add(pParam);

			};
		}
		pFunc->bConst = dlg.m_bConst;
		pFunc->bStatic = dlg.m_bStatic;
		pFunc->bVirtual = dlg.m_bVirtual;

		dlg.m_pCurClass->InsertFunction(pFunc, CString());
		delete pFunc;
	}

	return S_OK;
}

HRESULT CConnect::DoAddVariable()
{
	CComPtr<EnvDTE::Project> pProj;
	if (GetActiveProject(pProj) != S_OK)
	{
		return S_FALSE;
	}

	CSmartAtlArray<VSClass*> Classes;
	if (GetClasses(pProj, Classes) != S_OK)
	{
		return S_FALSE;
	}
	VSClass* pCurClass = NULL;

	pCurClass = Classes[GetActiveClass(Classes, CString())];

	CAddMemberVarGl dlg;
	dlg.SetForVar();
	dlg.m_pClassVector = &Classes;
	dlg.m_pCurClass = pCurClass;
	VARIANT_BOOL bStatic = VARIANT_FALSE;
	VARIANT_BOOL bConst = VARIANT_FALSE;

	if ((dlg.DoModal() == IDOK) && (dlg.m_pCurClass))
	{
		VSVariable* pVar = new VSVariable;
		pVar->Type = dlg.m_Type;
		pVar->Name = dlg.m_Body;
		if (dlg.m_bPublic)
		{
			pVar->Access = EnvDTE::vsCMAccessPublic;
		}
		else
		{
			if (dlg.m_bProtected)
			{
				pVar->Access = EnvDTE::vsCMAccessProtected;
			}
			else
			{
				pVar->Access = EnvDTE::vsCMAccessPrivate;
			}
		}
		while(pVar->Name[pVar->Name.GetLength()-1] == _T(';'))
		{
			pVar->Name.Delete(pVar->Name.GetLength()-1);
		}

		pVar->bConst = dlg.m_bConst;
		pVar->bStatic = dlg.m_bStatic;

		dlg.m_pCurClass->InsertVariable(pVar);
		delete pVar;
	}

	return S_OK;
}

HRESULT CConnect::DoOptions()
{
	g_pSelectedFunction = NULL;

	COptionsDlg dlg;
	_variant_t vtSolutionNewStyle;
	dlg.m_FuncOpt.m_bSolutionEnabled = 
		CCustomProjectSettings::m_sCustomProperties.GetSolutionVariableValue(_T("SolutionNewStyle"), 
		vtSolutionNewStyle);
	if(vtSolutionNewStyle.vt != VT_EMPTY)
	{
		int bSolutionNewStyle;
		if(vtSolutionNewStyle.vt == VT_BSTR)
		{
			bSolutionNewStyle = StrToIntW(vtSolutionNewStyle.bstrVal);
		}
		dlg.m_FuncOpt.m_bSolutionNewStyle = bSolutionNewStyle;
	}
	if(dlg.DoModal() == IDOK)
	{
		//CMessageManager::m_sHandlerManager.SetSolutionNewStyle(dlg.m_FuncOpt.m_bSolutionNewStyle);
		if(dlg.m_FuncOpt.m_bSolutionNewStyle != BST_INDETERMINATE)
		{
			CString str;
			str.Format(_T("%d"), dlg.m_FuncOpt.m_bSolutionNewStyle);
			CCustomProjectSettings::m_sCustomProperties.SetSolutionVariableValue(
				_T("SolutionNewStyle"), _variant_t(str));
		}
		else
		{
			CCustomProjectSettings::m_sCustomProperties.DeleteSolutionVariable(_T("SolutionNewStyle"));
		}
	}
	return S_OK;
}

HRESULT CConnect::DoClassViewHandler()
{
	g_pSelectedFunction = NULL;

	CComPtr<EnvDTE::Project> pProj;
	CComPtr<EnvDTE::CodeClass> pClass = GetSelectedClass(pProj);
	CString ClassName;
	if (pClass != NULL)
	{
		_bstr_t bsName;
		pClass->get_FullName(bsName.GetAddress());
		ClassName = (LPCTSTR)bsName;
	}
	ShowWTLHelper(pProj, ClassName, 0);
	return S_OK;
}

HRESULT CConnect::DoClassViewVariable()
{
	g_pSelectedFunction = NULL;

	CComPtr<EnvDTE::Project> pProj;
	CComPtr<EnvDTE::CodeClass> pClass = GetSelectedClass(pProj);
	CString ClassName;
	if (pClass != NULL)
	{
		_bstr_t bsName;
		pClass->get_FullName(bsName.GetAddress());
		ClassName = (LPCTSTR)bsName;
	}
	ShowWTLHelper(pProj, ClassName, 1);
	return S_OK;
}

HRESULT CConnect::DoResViewDDX()
{
	g_pSelectedFunction = NULL;

	CComPtr<EnvDTE::Project> pProj;
	int iDlgClass = -1;
	CAtlArray<CString> Controls;
	CSmartAtlArray<VSClass*> Classes;
	
	CSmartAtlArray<InsDelPoints> Modifications;
	CResourceManager ResManager;
	HRESULT hRes = PrepareDlgClass(pProj, Classes, iDlgClass, Controls, &ResManager);
	if (hRes == S_FALSE)
	{
		bool Res = ShowWTLHelper(pProj, Classes[iDlgClass]->Name, 1);
		for (size_t i = 0; i < Classes.GetCount(); i++)
		{
			delete Classes[i];
		}
		return Res ? S_OK : E_FAIL;
	}
	if (FAILED(hRes))
		return hRes;

	int iCurControl = -1;
	Classes[iDlgClass]->RetriveItems();
	Modifications.SetCount(Classes.GetCount());

	CDDXManager DDXManager;
	DDXManager.SetGlobalParams(NULL, &Modifications, &ResManager);
	DDXManager.Init(Classes[iDlgClass], &Modifications[iDlgClass]);
	for (size_t i = 0; i < Controls.GetCount() ;i++)
	{
		CString ControlID = Controls[i];
		if (ControlID == _T("IDC_STATIC"))
			continue;

		const CAtlArray<ResControl>& SelControls = DDXManager.GetControls();
		for (size_t j = 0; j < SelControls.GetCount(); j++)
		{
			if (SelControls[j].m_ID == ControlID)
			{
				iCurControl = (int)j;
				break;
			}
		}
		CString strCtrlId, strMemberName;
		VSMapEntry* pMapEntry = DDXManager.AddVariable(iCurControl, strCtrlId, strMemberName);
	}

	UpdateClasses(&Modifications, &Classes);
	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}
	Classes.RemoveAll();

	return S_OK;
}

HRESULT CConnect::DoResViewHandler()
{
	g_pSelectedFunction = NULL;

	CString ResourceID, RCFileName;
	CString ResourceType = GetActiveResourceType(ResourceID, RCFileName);
	if (ResourceType == _T("Dialog"))
	{
		return AddDialogHandler(RCFileName, ResourceID);
	}
	if (ResourceType == _T("Menu") ||
		ResourceType == _T("Toolbar") ||
		ResourceType == _T("Accelerator"))
	{
		return AddIDHandler(RCFileName, ResourceID, ResourceType);
	}

	return S_OK;
}

HRESULT CConnect::DoResViewReflect()
{
	int iDlgClass;
	CAtlArray<CString> Controls;
	CSmartAtlArray<VSClass*> Classes;
	CComPtr<EnvDTE::Project> pProj;

	CString DialogID, RCFile;
	CString ResourceType = GetActiveResourceType(DialogID, RCFile);
	
	CResourceManager ResManager;
	CSmartAtlArray<InsDelPoints> Modifications;
	HRESULT hRes = PrepareDlgClass(pProj, Classes, iDlgClass, Controls, &ResManager);
	if (hRes == S_FALSE)
	{
		bool bRes = ShowWTLHelper(pProj, Classes[iDlgClass]->Name, 0);
		for (size_t i = 0; i < Classes.GetCount(); i++)
		{
			delete Classes[i];
		}
		Classes.RemoveAll();
		return bRes ? S_OK : E_FAIL;
	}
	if (FAILED(hRes))
		return hRes;

	int iCurControl = -1;
	Modifications.SetCount(Classes.GetCount());

	CMessageManager MessageManager;
	MessageManager.Init(NULL, &ResManager, NULL);
	MessageManager.SetClass(Classes[iDlgClass], &Modifications[iDlgClass]);
	const CResDialog* pCurDlg = ResManager.GetDialog(DialogID);

	for (size_t i = 0; i < Controls.GetCount() ;i++)
	{
		CString ControlID = Controls[i];
		if (ControlID == _T("IDC_STATIC"))
			continue;

		const ResControl* pControl = ResManager.GetDlgControl(pCurDlg, ControlID);
		HandlerStruct Handler;
		MessageStruct Mes;
		eControlType eType = ResManager.GetControlType(pControl->m_Type);
		if (eType == eCTCommand)
		{
			Mes.Type = CUSTOM_COMMAND_REFLECTION_HANDLER;
		}
		else if (eType == eCTNotify)
		{
			Mes.Type = CUSTOM_NOTIFY_REFLECTION_HANDLER;
		}
		else
		{
			Mes.Type = CUSTOM_REFLECTION_HANDLER;
		}
		MessageManager.InsertReflectionHandler(&Mes, Handler, pControl->m_ID);
	}
	UpdateClasses(&Modifications, &Classes);

	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}
	Classes.RemoveAll();
	return S_OK;
}

HRESULT CConnect::DoResViewDialog()
{
	g_pSelectedFunction = NULL;

	CString DialogID;
	CString RCFileName;
	int iDlgClass = -1;
	if (GetActiveResourceType(DialogID, RCFileName) != _T("Dialog"))
		return S_OK;

	bool bDialogSelected = IsDialogSelected();

	CAtlArray<CString> Controls;
	if (!GetActiveControlID(Controls))
		return S_OK;

	_AtlModule.LoadAll();

	CComPtr<EnvDTE::Project> pProj;
	if (GetActiveProject(pProj) != S_OK)
	{
		return E_FAIL;
	}

	CSmartAtlArray<VSClass*> Classes;
	if (GetClasses(pProj, Classes) != S_OK)
	{
		return E_FAIL;
	}
	
	iDlgClass = FindClassByDlgID(Classes, DialogID);
	if (iDlgClass == -1)
	{
		CreateDialogClass(pProj, DialogID);
	}
	else
	{
		MessageBox(NULL, _T("Class for this dialog is already created!"), _T("Already created"), MB_ICONINFORMATION);
	}
	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}
	return S_OK;
}

HRESULT CConnect::DoResViewContextDDX()
{
	g_pSelectedFunction = NULL;

	CSmartAtlArray<VSClass*> Classes;
	CAtlArray<CString> Controls;
	CComPtr<EnvDTE::Project> pProj;

	int iDlgClass = -1;
	
	if (!GetActiveControlID(Controls) || !Controls.GetCount())
		return S_OK;

	_AtlModule.LoadAll();

	if (GetSelectedProject(pProj) != S_OK)
	{
		return E_FAIL;
	}

	
	if (GetClasses(pProj, Classes) != S_OK)
	{
		return E_FAIL;
	}
	if (!Classes.GetCount())
	{
		MessageBox(NULL, _T("No classes to work with!"), _T("No classes"), MB_OK | MB_ICONINFORMATION);
		return E_FAIL;
	}

	CString DialogID = Controls[0];
	iDlgClass = FindClassByDlgID(Classes, DialogID);
	if (iDlgClass == -1)
	{
		if (CreateDialogClass(pProj, DialogID) == EnvDTE::wizardResultSuccess)
		{
			for (size_t i = 0; i < Classes.GetCount(); i++)
			{
				delete Classes[i];
			}
			Classes.RemoveAll();
			if (GetClasses(pProj, Classes) != S_OK)
			{
				return E_FAIL;
			}
			iDlgClass = FindClassByDlgID(Classes, DialogID);
		}

		if (iDlgClass == -1)
			return S_OK;
	}

	bool Res = ShowWTLHelper(pProj, Classes[iDlgClass]->Name, 1);
		
	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}

	return S_OK;
}

HRESULT CConnect::DoResViewContextHandler()
{
	g_pSelectedFunction = NULL;

	CSmartAtlArray<VSClass*> Classes;
	CAtlArray<CString> Controls;
	CComPtr<EnvDTE::Project> pProj;

	int iDlgClass = -1;

	if (!GetActiveControlID(Controls) || !Controls.GetCount())
		return S_OK;

	_AtlModule.LoadAll();

	if (GetSelectedProject(pProj) != S_OK)
	{
		return E_FAIL;
	}

	if (GetClasses(pProj, Classes) != S_OK)
	{
		return E_FAIL;
	}
	if (!Classes.GetCount())
	{
		MessageBox(NULL, _T("No classes to work with!"), _T("No classes"), MB_OK | MB_ICONINFORMATION);
		return false;
	}

	CString DialogID = Controls[0];
	iDlgClass = FindClassByDlgID(Classes, DialogID);
	if (iDlgClass == -1)
	{
		if (CreateDialogClass(pProj, DialogID) == EnvDTE::wizardResultSuccess)
		{
			for (size_t i = 0; i < Classes.GetCount(); i++)
			{
				delete Classes[i];
			}
			Classes.RemoveAll();
			if (GetClasses(pProj, Classes) != S_OK)
			{
				return E_FAIL;
			}
			iDlgClass = FindClassByDlgID(Classes, DialogID);
		}

		if (iDlgClass == -1)
			return S_OK;
	}

	bool bRes = ShowWTLHelper(pProj, Classes[iDlgClass]->Name, 0);
	
	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}
	Classes.RemoveAll();
	return S_OK;
}

HRESULT CConnect::DoResViewContextDialog()
{
	g_pSelectedFunction = NULL;

	int iDlgClass = -1;
	CAtlArray<CString> Controls;
	CSmartAtlArray<VSClass*> Classes;

	if (!GetActiveControlID(Controls) || !Controls.GetCount())
		return S_OK;

	_AtlModule.LoadAll();

	CComPtr<EnvDTE::Project> pProj;
	if (GetSelectedProject(pProj) != S_OK)
	{
		return E_FAIL;
	}

	
	if (GetClasses(pProj, Classes) != S_OK)
	{
		return E_FAIL;
	}

	CString DialogID = Controls[0];
	iDlgClass = FindClassByDlgID(Classes, DialogID);
	if (iDlgClass == -1)
	{
		CreateDialogClass(pProj, DialogID);
	}
	else
	{
		MessageBox(NULL, _T("Class for this dialog is already created!"), _T("Already created"), MB_ICONINFORMATION);
	}
	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}
	return S_OK;
}

CComPtr<EnvDTE::CodeClass> CConnect::GetSelectedClass(CComPtr<EnvDTE::Project> & pProj)
{
	HRESULT hr = E_FAIL;
	CComPtr<EnvDTE::SelectedItems> pSelItems;
	CComPtr<EnvDTE::CodeClass> pClass;
	if (m_pDTE->get_SelectedItems(&pSelItems) == S_OK)
	{
		ATLASSERT(pSelItems != NULL);
		long Count;
		pSelItems->get_Count(&Count);
		if (Count)
		{
			CComPtr<EnvDTE::SelectedItem> pSelItem;
			pSelItems->Item(_variant_t(1), &pSelItem);
			ATLASSERT(pSelItem != NULL);
			CComPtr<EnvDTE::ProjectItem> pProjItem;
			pSelItem->get_ProjectItem(&pProjItem);
			if (pProjItem != NULL)
			{
				pProjItem->get_ContainingProject(&pProj);
			}
		}
				
		CComPtr<EnvDTE::SelectionContainer> pSelContainer;
		if (pSelItems->get_SelectionContainer(&pSelContainer) == S_OK)
		{
			if (pSelContainer != NULL)
			{
				CComPtr<IDispatch> pObject;
				hr = pSelContainer->Item(_variant_t(1), &pObject);
				if (SUCCEEDED(hr))
				{
					hr = pObject->QueryInterface(&pClass);
				}
			}
		}
	}
	return pClass;
}

HRESULT CConnect::GetActiveProject(CComPtr<EnvDTE::Project> & pProj)
{
	HRESULT hr = E_FAIL;
	CComPtr<EnvDTE::Document> pCurDoc;
	m_pDTE->get_ActiveDocument(&pCurDoc);
	if (pCurDoc != NULL)
	{
		CComPtr<EnvDTE::ProjectItem> pProjItem;
		pCurDoc->get_ProjectItem(&pProjItem);
		if (pProjItem == NULL)
		{
			return E_FAIL;
		}
		return pProjItem->get_ContainingProject(&pProj);
	}
	else
	{
		_variant_t vt, vt2;

		if (FAILED(m_pDTE->get_ActiveSolutionProjects(&vt)))
		{
			return S_FALSE;
		}

		if (vt.vt != (VT_ARRAY | VT_VARIANT))
		{
			return S_FALSE;
		}
		SAFEARRAY* pArray = vt.parray;
		if (pArray->cDims != 1 && !(pArray->fFeatures & FADF_DISPATCH ))
		{
			return S_FALSE;
		}

		if (pArray[0].rgsabound[0].cElements > 0)
		{
			vt2 = ((_variant_t*)pArray[0].pvData)[0];
			hr = vt2.pdispVal->QueryInterface(&pProj);
			if (pProj == NULL)
			{
				return S_FALSE;
			}
		}
		else
			return S_FALSE;
	}

	return S_OK;
}

HRESULT CConnect::GetSelectedProject(CComPtr<EnvDTE::Project>& pProj)
{
	CComPtr<EnvDTE::SelectedItems> pSelItems;
	if (m_pDTE->get_SelectedItems(&pSelItems) == S_OK)
	{
		ATLASSERT(pSelItems != NULL);
		long Count;
		pSelItems->get_Count(&Count);
		if (Count)
		{
			CComPtr<EnvDTE::SelectedItem> pSelItem;
			pSelItems->Item(_variant_t(1), &pSelItem);
			ATLASSERT(pSelItem != NULL);
			CComPtr<EnvDTE::ProjectItem> pProjItem;
			pSelItem->get_ProjectItem(&pProjItem);
			if (pProjItem != NULL)
			{
				return pProjItem->get_ContainingProject(&pProj);
			}
		}
	}
	return E_FAIL;
}

CComPtr<EnvDTE::CodeElement> CConnect::GetClassFromMember(CComPtr<EnvDTE::CodeElement> pMember)
{
	HRESULT hr = E_FAIL;
	CComPtr<EnvDTE::CodeElement> pParrentElem;
	CComPtr<EnvDTE::CodeElements> pColection = NULL;
	if (FAILED(pMember->get_Collection(&pColection)) || (pColection == NULL))
	{
		return pParrentElem;
	}

	CComPtr<IDispatch> pPar;
	pColection->get_Parent(&pPar);
	hr = pPar->QueryInterface(&pParrentElem);
	return pParrentElem;
}

EnvDTE::wizardResult CConnect::CreateDialogClass(CComPtr<EnvDTE::Project> pProj, CString DialogID, bool bShow /* = false */)
{
	CString WizardPath;
	LPCTSTR lpVersion = _T("1.1");
	eWizardsErrors Res = GetWTLDLGWizardPath(WizardPath, lpVersion);
	switch(Res) 
	{
	case eWizardsNotInstalled:
		{
			MessageBox(NULL, _T("WTL Wizards are not installed!\r\nYou cannot create new dialog class"), _T("No WTL Wizards"), MB_ICONWARNING);
			return EnvDTE::wizardResultFailure;
		}
	case eWTLDLGNotInstalled:
		{
			MessageBox(NULL, _T("WTL Dialog Wizard is not installed!\r\nYou cannot create new dialog class"), _T("No WTL Wizards"), MB_ICONWARNING);
			return EnvDTE::wizardResultFailure;
		}
	case eWrongVersion:
		{
			CString Mes;
			Mes.Format(_T("You need at least version %s of WTL Wizards to create new dialog class"), lpVersion);
			MessageBox(NULL, Mes, _T("No WTL Wizards"), MB_ICONWARNING);
			return EnvDTE::wizardResultFailure;
		}
	case eError:
		{
			MessageBox(NULL, _T("Unknow error!\r\nYou cannot create new dialog class"), _T("No WTL Wizards"), MB_ICONERROR);
			return EnvDTE::wizardResultFailure;
		}
	}
	CComSafeArray<VARIANT> ContextParams;
	_bstr_t Str;
	ContextParams.Add(_variant_t(EnvDTE::vsWizardAddItem));
	pProj->get_Name(Str.GetAddress());
	ContextParams.Add(_variant_t(Str));
	CComPtr<EnvDTE::ProjectItems> pProjItems;
	pProj->get_ProjectItems(&pProjItems);
	pProj->get_FullName(Str.GetAddress());
	CString FullName = (LPCTSTR)Str;
	int Pos = FullName.ReverseFind(_T('\\'));
	FullName.Delete(Pos, FullName.GetLength() - Pos);
	ContextParams.Add(_variant_t((LPDISPATCH)pProjItems));
	ContextParams.Add(_variant_t(FullName));
	ContextParams.Add(_variant_t((char*)NULL));
	ContextParams.Add(_variant_t(L"c:\\Program Files\\Microsoft Visual Studio .NET 2003\\Vc7"));
	ContextParams.Add(_variant_t(false));
	EnvDTE::wizardResult Result;
	//generate vsz file
	CStringA VSZData = "VSWIZARD 7.0\r\nWizard=VsWizard.VsWizardEngine.";
	_bstr_t DTEVersion;
	m_pDTE->get_Version(DTEVersion.GetAddress());
	CStringA strVersion((char*)DTEVersion);
	int iVLen = strVersion.GetLength();
	if ((iVLen >= 2) && (strVersion[iVLen - 1] == '0') && (strVersion[iVLen - 2] != '.'))
	{
		strVersion.Delete(iVLen - 1);
	}
	
	VSZData += strVersion;
	VSZData += "\r\nParam=\"WIZARD_NAME = WTLDLG\"\r\n";
	CStringA WizPath(WizardPath);
	VSZData += "Param=\"ABSOLUTE_PATH = " + WizPath + "\"\r\n";
	VSZData += CStringA("Param=\"DIALOG_ID = ") + CStringA(DialogID) + "\"\r\n";
	VSZData += CStringA("Param=\"SHOW_AFTER_CREATE = ") + (bShow ? "true\"" : "false\"");
	VSZData += "\r\n";
	//create vsz file in temporary directory
	TCHAR Buf[FILENAME_MAX];
	GetTempPath(FILENAME_MAX, Buf);
	CString FileName = Buf;
	if (FileName[FileName.GetLength() - 1] == _T('\\'))
		FileName += _T("WTLDLG.vsz");
	else
		FileName += _T("\\WTLDLG.vsz");
	CAtlFile File;
	if (FAILED(File.Create(FileName, FILE_ALL_ACCESS, 0, CREATE_ALWAYS)))
		return EnvDTE::wizardResultFailure;
	File.Write(VSZData, (DWORD)VSZData.GetLength());
	File.Flush();
	File.Close();

	HRESULT hr = m_pDTE->LaunchWizard(_bstr_t(FileName), &ContextParams.m_psa, &Result);
	DeleteFile(FileName);

	return Result;
}

eWizardsErrors CConnect::GetWTLDLGWizardPath(CString& Path, LPCTSTR lpMinVersion)
{
	CRegKey Key;
	if (Key.Open(HKEY_LOCAL_MACHINE, _T("Software\\SaloS\\WTL Wizards"), KEY_READ) == ERROR_SUCCESS)
	{
		ULONG Len;
		CString RegPath;
		LONG Res = Key.QueryStringValue(_T("InstallPath"), NULL, &Len);
		if (Res == ERROR_MORE_DATA || Res == ERROR_SUCCESS)
		{
			Key.QueryStringValue(_T("InstallPath"), RegPath.GetBuffer(Len + 1), &Len);
			RegPath.ReleaseBuffer();

			TCHAR VersionBuf[256];
			Len = 256;
			Res = Key.QueryStringValue(_T("Version"), VersionBuf, &Len);
			if (Res == ERROR_SUCCESS)
			{
				if (lpMinVersion)
				{
					if (CompareVersions(VersionBuf, lpMinVersion) < 0)
						return eWrongVersion;
				}
				Path = RegPath + _T("WTLDLG");
				return eSuccess;
			}
		}
	}

	LPCTSTR lpProductCode = TEXT("{C1EBF990-BE54-4e43-BA92-D0731263B016}");
	UINT Res;

	Res = MsiQueryProductState(lpProductCode);
	if (Res != INSTALLSTATE_DEFAULT)
	{
		return eWizardsNotInstalled;
	}
	TCHAR buf[FILENAME_MAX];
	DWORD dwSize = FILENAME_MAX;

	Res = MsiGetProductInfo(lpProductCode, INSTALLPROPERTY_VERSIONSTRING, buf, &dwSize);
	if (Res != ERROR_SUCCESS)
	{
		return eError;
	}
	buf[dwSize] = 0;
	if (lpMinVersion)
	{
		if (CompareVersions(buf, lpMinVersion) < 0)
			return eWrongVersion;
	}

	dwSize = FILENAME_MAX;
	Res = MsiGetComponentPath(lpProductCode, _T("{6802889B-ADEB-45eb-885D-C55144ADE507}"), buf, &dwSize);
	if (Res == INSTALLSTATE_ABSENT || Res == INSTALLSTATE_UNKNOWN)
	{
		return eWTLDLGNotInstalled;
	}
	buf[dwSize] = 0;
	LPCTSTR lpEndPath = _tcsstr(buf, _T("\\HTML"));
	if (lpEndPath != NULL)
	{
		Path = CString(buf, (int)(lpEndPath - buf));
	}

	return eSuccess;
}

CString CConnect::GetActiveClass(CComPtr<EnvDTE::TextPoint>& pCurPoint)
{
	CComPtr<EnvDTE::Document> pDoc = NULL;
	if (FAILED(m_pDTE->get_ActiveDocument(&pDoc)) || (pDoc == NULL))
	{
		return CString();
	}

	CComPtr<EnvDTE::ProjectItem> pItem = NULL;
	if (FAILED(pDoc->get_ProjectItem(&pItem)) || (pItem == NULL))
	{
		return CString();
	}

	CComPtr<EnvDTE::FileCodeModel> pFileModel = NULL;
	if (FAILED(pItem->get_FileCodeModel(&pFileModel)) || pFileModel == NULL)
	{
		return CString();
	}

	CComPtr<EnvDTE::TextSelection> pSel;
	CComPtr<IDispatch> pDisp;
	if (FAILED(pDoc->get_Selection(&pDisp)) || pDisp==NULL)
	{
		return CString();
	}
	pDisp->QueryInterface(&pSel);
	if (pSel != NULL)
	{
		CComPtr<EnvDTE::VirtualPoint> pCurPos = NULL;
		if (FAILED(pSel->get_ActivePoint(&pCurPos)) || (pCurPos == NULL))
		{
			return CString();
		}

		CComPtr<EnvDTE::CodeElement> pElement = NULL;
		if (FAILED(pFileModel->CodeElementFromPoint(pCurPos, EnvDTE::vsCMElementClass, &pElement)) || (pElement == NULL))
		{
			pFileModel->CodeElementFromPoint(pCurPos, EnvDTE::vsCMElementFunction, &pElement);
		}
		if (pElement != NULL)
		{
			VSElement Elem(EnvDTE::vsCMElementOther);
			EnvDTE::vsCMElement ElemType;
			pElement->get_Kind(&ElemType);

			if (ElemType == EnvDTE::vsCMElementClass)
			{
				Elem.pElement = pElement;
				Elem.RetriveName(true);
				pCurPoint = pCurPos;
			}

			if (ElemType == EnvDTE::vsCMElementFunction)
			{
				CComPtr<EnvDTE::CodeElement> pParrentElem = GetClassFromMember(pElement);
				
				Elem.pElement = pParrentElem;
				Elem.RetriveName(true);
			}
			return Elem.Name;
		}
	}

	if (pFileModel != NULL)
	{
		CComPtr<EnvDTE::CodeElements> Elements;
		pFileModel->get_CodeElements(&Elements);
		long Count;
		Elements->get_Count(&Count);
		CString ClassName;
		for (long i = 1; i <= Count; i++)
		{
			CComPtr<EnvDTE::CodeElement> pElem;
			Elements->Item(_variant_t(i), &pElem);
			
			EnvDTE::vsCMElement ElemType;
			pElem->get_Kind(&ElemType);
#ifdef _DEBUG
			_bstr_t bsName;
			pElem->get_Name(bsName.GetAddress());
#endif
			
			if (ElemType == EnvDTE::vsCMElementClass)
			{
				_bstr_t Name;
				pElem->get_FullName(Name.GetAddress());
				ClassName = (LPCTSTR)Name;
				break;
			}
			if (ElemType == EnvDTE::vsCMElementFunction)
			{
				CComPtr<EnvDTE::CodeElement> pClass = GetClassFromMember(pElem);
				_bstr_t Name;
				pClass->get_FullName(Name.GetAddress());
				ClassName = (LPCTSTR)Name;
				break;
			}
		}
		return ClassName;
	}
	return CString();
}

int CConnect::GetActiveClass(CSmartAtlArray<VSClass*>& Classes, CString ClassName)
{
	CComPtr<EnvDTE::TextPoint> pCurPtr;
	CString CurrentClass;
	if (ClassName.IsEmpty())
	{
		CurrentClass = GetActiveClass(pCurPtr);
	}
	else
	{
		CurrentClass = ClassName;
	}
	int iCurClass = 0;

	for (size_t j = 0; j < Classes.GetCount(); j++)
	{
		if (Classes[j]->Name == CurrentClass)
		{
			if (pCurPtr != NULL)
			{
				//probably nested classes
				VSClass* pCurClass = Classes[j];
				bool bNestedClass = false;
				for (size_t k = 0; k < pCurClass->NestedClasses.GetCount(); k++)
				{
					VSClass* pNestedClass = pCurClass->NestedClasses[k];
					CComPtr<EnvDTE::EditPoint> pStart, pEnd;
					pNestedClass->GetStartPoint(&pStart);
					pNestedClass->GetEndPoint(&pEnd);
					VARIANT_BOOL bInside = VARIANT_FALSE;
					pCurPtr->GreaterThan(pStart, &bInside);
					if (bInside == VARIANT_TRUE)
					{
						pCurPtr->LessThan(pEnd, &bInside);
						if (bInside == VARIANT_TRUE)
						{
							//found that active is nested class
							//change name and go to next step because
							//nested classes are after parent class
							CurrentClass = pNestedClass->Name;
							bNestedClass = true;
							break;
						}
					}
				}
				if (bNestedClass)
					continue;
			}
			iCurClass = (int)j;
			break;
		}
	}
	return iCurClass;
}

CString CConnect::GetActiveResourceType(CString& ResourceID, CString& FileName)
{
	CComPtr<EnvDTE::Window> pCurWindow;
	m_pDTE->get_ActiveWindow(&pCurWindow);
	if (pCurWindow == NULL)
		return CString();

	_bstr_t WindowCaption;
	pCurWindow->get_Caption(WindowCaption.GetAddress());
	CString strCaption = WindowCaption;
	CString ResourceType;
	int OpenBracket = strCaption.Find(_T('('));
	int CloseBracket = strCaption.ReverseFind(_T(')'));
	if (OpenBracket != -1 && CloseBracket != -1)
	{
		FileName = strCaption.Left(OpenBracket);
		FileName.Trim();
		strCaption = strCaption.Mid(OpenBracket + 1, CloseBracket - OpenBracket-1);
		strCaption.Trim();
		int iDash = strCaption.ReverseFind(_T('-'));
		if (iDash == -1)
		{
			ResourceType = strCaption;
		}
		else
		{
			ResourceType = strCaption.Mid(iDash + 1);
			ResourceType.Trim();
			ResourceID = strCaption.Left(iDash);
			ResourceID.Trim();
		}
	}
	return ResourceType;
}

bool CConnect::IsPossibleResourceType(CString ResourceType)
{
	if (ResourceType == _T("Dialog") || ResourceType == _T("Menu") ||
		ResourceType == _T("Accelerator") || ResourceType == _T("Toolbar"))
	{
		return true;
	}
	return false;
}

bool CConnect::IsDialogSelected()
{
	CComPtr<EnvDTE::SelectedItems> pSelItems;
	CComPtr<EnvDTE::SelectionContainer> pSelContainer;
	m_pDTE->get_SelectedItems(&pSelItems);

	CComPtr<IDispatch> pObject;
	pSelItems->get_SelectionContainer(&pSelContainer);
	if (pSelContainer == NULL)
		return false;

	long Count;
	pSelContainer->get_Count(&Count);
	bool bRes = false;
	if (Count)
	{
		pSelContainer->Item(_variant_t(1), &pObject);
		if (pObject == NULL)
		{
			return false;
		}
		CComPtr<ITypeInfo> pTypeInfo;
		pObject->GetTypeInfo(1, GetSystemDefaultLCID(), &pTypeInfo);
		if (pTypeInfo != NULL)
		{
			LPTYPEATTR pTypeAttr;
			pTypeInfo->GetTypeAttr(&pTypeAttr);
			if (pTypeAttr->guid == __uuidof(ResEditPKGLib::IDlgEditor) ||
				pTypeAttr->guid == __uuidof(ResEditPKGLib::IDlgRes))
			{
				bRes = true;
			}
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
	}
	return bRes;
}

bool CConnect::IsMultipleSelected()
{
	CComPtr<EnvDTE::SelectedItems> pSelItems;
	CComPtr<EnvDTE::SelectionContainer> pSelContainer;
	m_pDTE->get_SelectedItems(&pSelItems);

	pSelItems->get_SelectionContainer(&pSelContainer);
	long Count;
	pSelContainer->get_Count(&Count);
	return (Count > 1);
}

bool CConnect::GetActiveControlID(CAtlArray<CString>& ActiveControls)
{
	CString ResID;
	CComPtr<EnvDTE::SelectedItems> pSelItems;
	CComPtr<EnvDTE::SelectionContainer> pSelContainer;
	m_pDTE->get_SelectedItems(&pSelItems);

	CComPtr<IDispatch> pObject;
	pSelItems->get_SelectionContainer(&pSelContainer);
	long Count;
	pSelContainer->get_Count(&Count);
	for (long i = 1; i <= Count; i++)
	{
		pSelContainer->Item(_variant_t(i), &pObject);
		if (pObject == NULL)
		{
			return false;
		}
		_variant_t varRes;
		HRESULT hr = GetDispatchProperty(pObject, L"ID", &varRes);
		if (hr == S_OK)
		{
			_bstr_t bsID = varRes;
			ResID = (LPCTSTR)bsID;
			ActiveControls.Add(ResID);
		}
	}
	return (Count > 0);
}

int CConnect::FindClassByDlgID(ClassVector Classes, CString DialogID)
{
	HRESULT hr = E_FAIL; 
	int iDlgClass = -1;
	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		CComPtr<VCCodeModelLibrary::VCCodeClass> pVCClass;
		pVCClass = Classes[i]->pElement;
		CComPtr<EnvDTE::CodeElements> pEnums;
		hr = pVCClass->get_Enums(&pEnums);
		long Count;
		pEnums->get_Count(&Count);
		for (long j = 1; j <= Count; j++ )
		{
			CComPtr<EnvDTE::CodeElement> pElem;
			pEnums->Item(_variant_t(j), &pElem);
			CComPtr<VCCodeModelLibrary::VCCodeEnum> pEnum; 
			hr = pElem->QueryInterface(&pEnum);
			if (pEnum != NULL)
			{
				CComPtr<EnvDTE::CodeElements> pMembers; 
				hr = pEnum->get_Members(&pMembers);
				CComPtr<EnvDTE::CodeVariable> pMember;
				pElem.Release();
				pMembers->Item(_variant_t(L"IDD"), &pElem);
				hr = pElem->QueryInterface(&pMember);
				if (pMember != NULL)
				{
					_bstr_t Value;
					_variant_t vt;
					pMember->get_InitExpression(&vt);
					Value = vt;
					if ((LPCTSTR)Value == DialogID)
					{
						iDlgClass = (int)i;
						break;
					}
				}
			}
		}
	}
	return iDlgClass;
}

void CConnect::SaveResourceDocuments()
{
	// save resource files
	// information about resources is got from files directly
	// so this step is important
	CComPtr<EnvDTE::Documents> pDocuments;
	m_pDTE->get_Documents(&pDocuments);
	if (pDocuments != NULL)
	{
		long Count;
		pDocuments->get_Count(&Count);
		for (long i = 1; i <= Count; i++)
		{
			CComPtr<EnvDTE::Document> pDoc;
			pDocuments->Item(_variant_t(i), &pDoc);
			_bstr_t str;
			pDoc->get_FullName(str.GetAddress());
			CString Name = str;
			if (Name.Right(3) == _T(".rc"))
			{
				VARIANT_BOOL bSaved = VARIANT_FALSE;
				pDoc->get_Saved(&bSaved);
				if (bSaved == VARIANT_FALSE)
				{
					EnvDTE::vsSaveStatus SaveStatus;
					pDoc->Save(NULL, &SaveStatus);
				}
			}
		}
	}
}

HRESULT CConnect::AddDialogHandler(CString RCFile, CString DialogID)
{
	CAtlArray<CString> Controls;
	CSmartAtlArray<VSClass*> Classes;
	CComPtr<EnvDTE::Project> pProj;
	int iDlgClass;

	CSmartAtlArray<InsDelPoints> Modifications;
	CResourceManager ResManager;
	HRESULT hRes = PrepareDlgClass(pProj, Classes, iDlgClass, Controls, &ResManager);
	if (hRes == S_FALSE)
	{
		bool bRes = ShowWTLHelper(pProj, Classes[iDlgClass]->Name, 0);
		for (size_t i = 0; i < Classes.GetCount(); i++)
		{
			delete Classes[i];
		}
		Classes.RemoveAll();
		return bRes ? S_OK : E_FAIL;
	}
	
	if (FAILED(hRes))
		return hRes;

	int iCurControl = -1;
	Modifications.SetCount(Classes.GetCount());

	CMessageManager MessageManager;
	MessageManager.Init(NULL, &ResManager, NULL);
	const CResDialog* pCurDlg = ResManager.GetDialog(DialogID);
		
	for (size_t i = 0; i < Controls.GetCount() ;i++)
	{
		CString ControlID = Controls[i];
		if (ControlID == _T("IDC_STATIC"))
			continue;
		
		const ResControl* pControl = ResManager.GetDlgControl(pCurDlg, ControlID);
		MessageManager.InsertDialogControlMessageHandler(pControl, Classes, Modifications, iDlgClass);
	}
	UpdateClasses(&Modifications, &Classes);
	
	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}
	Classes.RemoveAll();
	return S_OK;
}

HRESULT CConnect::AddIDHandler(CString RCFile, CString ResID, CString ResType)
{
	CAtlArray<CString> Controls;
	if (!GetActiveControlID(Controls))
		return S_OK;

	_AtlModule.LoadAll();

	CComPtr<EnvDTE::Project> pProj;
	if (GetActiveProject(pProj) != S_OK)
	{
		return false;
	}

	CSmartAtlArray<VSClass*> Classes;
	if (GetClasses(pProj, Classes) != S_OK)
	{
		return false;
	}
	if (!Classes.GetCount())
	{
		MessageBox(NULL, _T("No classes to work with!"), _T("No classes"), MB_OK | MB_ICONINFORMATION);
		return false;
	}

	SaveResourceDocuments();

	CSmartAtlArray<InsDelPoints> Modifications;
	CResourceManager ResManager;

	CComPtr<EnvDTE::ProjectItem> pRCFile = FindItem(pProj, _bstr_t(RCFile), CComPtr<EnvDTE::ProjectItem>());
	if (pRCFile != NULL)
	{
		ResManager.LoadResources(pRCFile);
	}

	int iCurControl = -1;
	Modifications.SetCount(Classes.GetCount());
	CMessageManager MessageManager;
	CAtlArray<CString> ResIds;
	if (ResType == _T("Menu"))
	{
		ResManager.GetMenuIds(ResID, ResIds);
	}
	if (ResType == _T("Toolbar"))
	{
		ResManager.GetToolbarIds(ResID, ResIds);
	}
	if (ResType == _T("Accelerator"))
	{
		ResManager.GetAcceleratorIds(ResID, ResIds);
	}
	MessageManager.Init(NULL, &ResManager, &ResIds);

	int iDefClass = -1;
	// select default class
	CString DefClassName = _T("CMainFrame");
	CComPtr<EnvDTE::Globals> pGlobals;
	pProj->get_Globals(&pGlobals);
	if (pGlobals != NULL)
	{
		CString ParamName = _T("WTLHELPER.")+ ResType + _T(".") + ResID;
		_variant_t vt;
		HRESULT hr = pGlobals->get_VariableValue(_bstr_t(ParamName), &vt);
		_bstr_t Value = vt;
		if (Value.length())
		{
			DefClassName = (LPCTSTR)Value;
		}
	}

	// find class in class list
	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		if (Classes[i]->Name == DefClassName)
		{
			iDefClass = (int)i;
			break;
		}
	}
	if (iDefClass == -1)
		iDefClass = 0;

	bool bInsRes = false;
	for (size_t i = 0; i < Controls.GetCount() ;i++)
	{
		CString ControlID = Controls[i];
		bInsRes = MessageManager.InsertIDMessageHandler(Controls[i], Classes, Modifications, iDefClass);
	}
	UpdateClasses(&Modifications, &Classes);
	if (pGlobals != NULL && bInsRes)
	{
		CString ParamName = _T("WTLHELPER.")+ ResType + _T(".") + ResID;
		_variant_t vt = Classes[iDefClass]->Name;

		HRESULT hr = pGlobals->put_VariableValue(_bstr_t(ParamName), vt);
		pGlobals->put_VariablePersists(_bstr_t(ParamName), VARIANT_TRUE);
	}

	for (size_t i = 0; i < Classes.GetCount(); i++)
	{
		delete Classes[i];
	}
	Classes.RemoveAll();

	return S_OK;
}

HRESULT CConnect::AddClassToVector(CComPtr<EnvDTE::CodeElement> pItem, CSmartAtlArray<VSClass*>& Classes, VSClass* pParentClass /*  = NULL */)
{
	HRESULT hr = E_FAIL;
	VSClass* pClass = new VSClass;
	pClass->pElement = pItem;
	pClass->pCodeModel = m_pCodeModel;
	pClass->RetriveName(true);
	if (pParentClass)
	{
		pParentClass->NestedClasses.Add(pClass);
	}
	Classes.Add(pClass);
	CComPtr<VCCodeModelLibrary::VCCodeClass> pVCClass;
	hr = pItem->QueryInterface(&pVCClass);
	if (pVCClass != NULL)
	{
		CComPtr<EnvDTE::CodeElements> pNestedClasses;
		hr = pVCClass->get_Classes(&pNestedClasses);
		long NestedCount;
		pNestedClasses->get_Count(&NestedCount);
		for (long j = 1; j <= NestedCount; j++ )
		{
			CComPtr<EnvDTE::CodeElement> pNestedClass;
			pNestedClasses->Item(_variant_t(j), &pNestedClass);
			AddClassToVector(pNestedClass, Classes, pClass);
		}
	}
	return S_OK;
}

HRESULT CConnect::AddNamespaceToVector(CComPtr<EnvDTE::CodeElement> pItem, CSmartAtlArray<VSClass*>& Classes)
{
	HRESULT hr = E_FAIL;
	CComPtr<VCCodeModelLibrary::VCCodeNamespace> pNamespace; 
	hr = pItem->QueryInterface(&pNamespace);
	if (pNamespace != NULL)
	{
		CComPtr<EnvDTE::CodeElements> pClasses, pNamespaces;
		hr = pNamespace->get_Classes(&pClasses);
		long Count;
		pClasses->get_Count(&Count);
		for (long i = 1; i <= Count; i++)
		{
			CComPtr<EnvDTE::CodeElement> pElem;
			pClasses->Item(_variant_t(i), &pElem);
			AddClassToVector(pElem, Classes);
		}
		hr = pNamespace->get_Namespaces(&pNamespaces);
		pNamespaces->get_Count(&Count);
		for (long i = 1; i <= Count; i++)
		{
			CComPtr<EnvDTE::CodeElement> pElem;
			pNamespaces->Item(_variant_t(i), &pElem);
			AddNamespaceToVector(pElem, Classes);
		}
	}
	return S_OK;
}

HRESULT CConnect::GetClasses(EnvDTE::Project * pProj, CSmartAtlArray<VSClass*>& Classes)
{
	HRESULT hr = E_FAIL;
	CComPtr<EnvDTE::CodeElements> pElements = NULL;

	if (pProj == NULL)
	{
		return E_FAIL;
	}

	if (FAILED(pProj->get_CodeModel(&m_pCodeModel)))
	{
		return E_FAIL;
	}

#if defined(_FOR_VS2008)
	CComPtr<EnvDTE80::CodeModel2> pCodeModel2; 
	hr = m_pCodeModel->QueryInterface(&pCodeModel2);
	pCodeModel2->Synchronize();
#elif defined(_FOR_VS2005)
	CComPtr<EnvDTE80::CodeModel2> pCodeModel2; 
	hr = m_pCodeModel->QueryInterface(&pCodeModel2);
	pCodeModel2->Synchronize();
#else
#endif

	if (FAILED(m_pCodeModel->get_CodeElements(&pElements)))
	{
		return E_FAIL;
	}

	long Count;
	pElements->get_Count(&Count);
	for (long i = 1; i <= Count; i++)
	{
		_variant_t vt = i;
		CComPtr<EnvDTE::CodeElement> pItem;
		if (FAILED(pElements->Item(vt, &pItem)) || (pItem == NULL))
		{
			return E_FAIL;
		}

		EnvDTE::vsCMElement type;
		pItem->get_Kind(&type);
		if (type == EnvDTE::vsCMElementClass)
		{
			AddClassToVector(pItem, Classes);
		}
		if (type == EnvDTE::vsCMElementNamespace)
		{
			AddNamespaceToVector(pItem, Classes);
		}
	}

	return S_OK;
}

HRESULT CConnect::PrepareDlgClass(CComPtr<EnvDTE::Project> & pProj, CSmartAtlArray<VSClass*>& Classes, int& iDlgClass, CAtlArray<CString>& Controls, CResourceManager* pResManager)
{
	CString DialogID;
	CString RCFileName;
	
	if (GetActiveResourceType(DialogID, RCFileName) != _T("Dialog"))
		return E_FAIL;

	if (!GetActiveControlID(Controls))
		return E_FAIL;

	bool bDialogSelected = IsDialogSelected();
	_AtlModule.LoadAll();

	if (GetActiveProject(pProj) != S_OK)
	{
		return E_FAIL;
	}

	if (GetClasses(pProj, Classes) != S_OK)
	{
		return E_FAIL;
	}
	if (!Classes.GetCount())
	{
		MessageBox(NULL, _T("No classes to work with!"), _T("No classes"), MB_OK | MB_ICONINFORMATION);
		return E_FAIL;
	}

	iDlgClass = FindClassByDlgID(Classes, DialogID);
	if (iDlgClass == -1)
	{
		if (CreateDialogClass(pProj, DialogID) == EnvDTE::wizardResultSuccess)
		{
			for (size_t i = 0; i < Classes.GetCount(); i++)
			{
				delete Classes[i];
			}
			Classes.RemoveAll();
			if (GetClasses(pProj, Classes) != S_OK)
			{
				return E_FAIL;
			}
			iDlgClass = FindClassByDlgID(Classes, DialogID);
		}

		if (iDlgClass == -1)
			return E_FAIL;
	}
	if (bDialogSelected)
	{
		/*bool bRes = ShowWTLHelper(pProj, Classes[iDlgClass]->Name, 0);
		for (size_t i = 0; i < Classes.GetCount(); i++)
		{
			delete Classes[i];
		}
		Classes.RemoveAll();*/
		return S_FALSE;
	}

	SaveResourceDocuments();

	CComPtr<EnvDTE::ProjectItem> pRCFile = FindItem(pProj, _bstr_t(RCFileName), CComPtr<EnvDTE::ProjectItem>());
	if (pRCFile != NULL && pResManager)
	{
		pResManager->LoadResources(pRCFile);
	}
	return S_OK;
}

#if defined(_FOR_VS2005) || defined(_FOR_VS2008)

HRESULT CConnect::UninstallAddin()
{
	HRESULT hr = S_OK;
	CComPtr<EnvDTE::Commands> pCommands;
	UICommandBarsPtr pCommandBars;
	UICommandBarPtr pHostCmdBar;
		
	m_pDTE->get_Commands(&pCommands);

	//removing Commands
	CString strBase = AddinString();
	for(size_t i =0; i < _countof(WtlHelperCommands); i++)
	{
		CString str = strBase + (LPCTSTR)WtlHelperCommands[i]->Name;
		CComPtr<EnvDTE::Command> pCmd;
		pCommands->Item(_variant_t(str), 0, &pCmd);
		if (pCmd != NULL)
		{
			HRESULT hr = pCmd->Delete();
		}
	}

	m_pAddInInstance->Remove();
	m_pDTE->Quit();

	return hr;
}

#endif

//////////////////////////////////////////////////////////////////////////

void UpdateClasses(CSmartAtlArray<InsDelPoints>* pModifications, ClassVector* pClassVector)
{
	HRESULT hr = E_FAIL;
	VSFunction* pLastFunction = NULL;
	for (size_t i = 0; i < (*pModifications).GetCount(); i++)
	{
		//сначала удаляем, затем добавляем
		CSmartAtlArray<VSElement*>& CurDelPoint = (*pModifications)[i].DeletePoint;
		CSmartAtlArray<InsertionPoint*>& CurInsPoint = (*pModifications)[i].InsertPoints;
		for (size_t l = 0; l < CurDelPoint.GetCount(); l++)
		{
			(CurDelPoint[l])->Remove();
			delete (CurDelPoint[l]);
		}
		CurDelPoint.RemoveAll();

		int Step = INSERT_STEP_ENVDTE;

		for (size_t j = 0; j < CurInsPoint.GetCount(); j++)
		{
			if ((CurInsPoint[j])->Insert((*pClassVector)[i], Step) == S_OK)
			{
				if ((CurInsPoint[j])->pElement)
				{
					if ((CurInsPoint[j])->pElement->ElementType == EnvDTE::vsCMElementFunction)
					{
						pLastFunction = (VSFunction*)(CurInsPoint[j])->pElement;
					}
				}
				if (CurInsPoint[j]->Type == INSERT_POINT_HANDLER)
				{
					VSFunction* pFunc = ((InsertPointHandler*)(CurInsPoint[j]))->pFunction;
					if (pFunc)
						pLastFunction = pFunc;
				}
			}
		}

		for (int k = INSERT_STEP_GLOBAL; k <= INSERT_STEP_MAP_ENTRY; k++)
		{
			for (size_t j = 0; j < CurInsPoint.GetCount(); j++)
			{
				(CurInsPoint[j])->Insert((*pClassVector)[i], k);
			}
		}

		// удаление объектов
		for (size_t j = 0; j < CurInsPoint.GetCount(); j++)
		{
			delete (CurInsPoint[j]);
		}
		CurInsPoint.RemoveAll();
	}
	if (!g_pSelectedFunction)
	{
		g_pSelectedFunction = pLastFunction;
	}
	if (g_pSelectedFunction)
	{
		CComPtr<EnvDTE::ProjectItem> pProjItem;
		CComPtr<VCCodeModelLibrary::VCCodeFunction> pFunc;
		pFunc = g_pSelectedFunction->pElement;
		if(pFunc != NULL)
		{
			hr = pFunc->get_ProjectItem(&pProjItem);

			CComPtr<EnvDTE::Window> pWindow;
			HRESULT hRes = pProjItem->Open(_bstr_t(EnvDTE::vsViewKindPrimary), &pWindow);
			if (SUCCEEDED(hRes))
			{
				CComPtr<EnvDTE::EditPoint> pEditPoint;
				if (g_pSelectedFunction->GetStartPoint(&pEditPoint) == S_OK)
				{
					VARIANT_BOOL bShow;
					pEditPoint->TryToShow(EnvDTE::vsPaneShowCentered, _variant_t(0), &bShow);
				}
			}
		}
	}
	pModifications->RemoveAll();
}

CComPtr<EnvDTE::ProjectItem> FindItem(EnvDTE::Project * pProject, _bstr_t ItemName, CComPtr<EnvDTE::ProjectItem> pPrevElem)
{
	CComPtr<EnvDTE::ProjectItems> pItems;
	if (pPrevElem == NULL)
	{
		pProject->get_ProjectItems(&pItems);
	}
	else
	{
		pPrevElem->get_ProjectItems(&pItems);
	}
	if (pItems == NULL)
		return CComPtr<EnvDTE::ProjectItem>(NULL);
	long Count;
	pItems->get_Count(&Count);
	if (Count == 0)
		return CComPtr<EnvDTE::ProjectItem>(NULL);
	for (long i = 1; i <= Count; i++)
	{
		CComPtr<EnvDTE::ProjectItem> pItem;
		pItems->Item(_variant_t(i), &pItem);
		_bstr_t IName;
		pItem->get_Name(IName.GetAddress());

		if (!_wcsicmp(IName, ItemName))
		{
			return pItem;
		}

		CComPtr<EnvDTE::ProjectItem> pItem2 = FindItem(pProject, ItemName, pItem);
		if (pItem2 != NULL)
			return pItem2;
	}
	return CComPtr<EnvDTE::ProjectItem>(NULL);
}

CComPtr<EnvDTE::CodeElement> FindDefine(CComPtr<EnvDTE::ProjectItem> pItem, LPCWSTR Name)
{
	HRESULT hr = E_FAIL;
	CComPtr<EnvDTE::FileCodeModel> pFileCodeModel;
	CComPtr<VCCodeModelLibrary::VCFileCodeModel> pVCFileCodeModel;
	hr = pItem->get_FileCodeModel(&pFileCodeModel);
	ATLASSERT(hr == S_OK);
	if (hr != S_OK)
		return CComPtr<EnvDTE::CodeElement>();

	pVCFileCodeModel = pFileCodeModel;
	ATLASSERT(pVCFileCodeModel != NULL);

	CComPtr<EnvDTE::CodeElements> pDefines;
	hr = pVCFileCodeModel->get_Macros(&pDefines);
	ATLASSERT(hr == S_OK);
	CComPtr<EnvDTE::CodeElement> pDefine;
	if ((pDefines->Item(_variant_t(Name), &pDefine) == S_OK) && (pDefine != NULL))
	{
		return pDefine;
	}
	else
	{
		return CComPtr<EnvDTE::CodeElement>();
	}
}

CComPtr<EnvDTE::CodeElement> FindDefine(VSClass* pClass, LPCWSTR Name, bool bStdAfx /* = false */)
{
	CComPtr<EnvDTE::Project> pProject;
	CComPtr<EnvDTE::ProjectItem> pProjectItem;
	CComPtr<EnvDTE::CodeElement> pDefine;

	pClass->pElement->get_ProjectItem(&pProjectItem);
	pDefine = FindDefine(pProjectItem, Name);
	if (pDefine != NULL)
		return pDefine;
	if (bStdAfx)
	{
		pProjectItem->get_ContainingProject(&pProject);
		CComPtr<EnvDTE::ProjectItem> pStdAfxFile = FindItem(pProject, _bstr_t(L"stdafx.h"), CComPtr<EnvDTE::ProjectItem>(NULL));
		return FindDefine(pStdAfxFile, Name);
	}
	return pDefine;
}

CComPtr<EnvDTE::CodeElement> FindInclude(CComPtr<EnvDTE::ProjectItem> pItem, LPCWSTR Name)
{
	CComPtr<EnvDTE::FileCodeModel> pFileCodeModel;
	CComPtr<VCCodeModelLibrary::VCFileCodeModel> pVCFileCodeModel;
	HRESULT hr;
	hr = pItem->get_FileCodeModel(&pFileCodeModel);
	ATLASSERT(hr == S_OK);
	if (hr != S_OK)
		return CComPtr<EnvDTE::CodeElement>();

	pVCFileCodeModel = pFileCodeModel;
	ATLASSERT(pVCFileCodeModel != NULL);

	CComPtr<EnvDTE::CodeElements> pIncludes;
	hr = pVCFileCodeModel->get_Includes(&pIncludes);
	ATLASSERT(hr == S_OK);
	CComPtr<EnvDTE::CodeElement> pInclude;
	if ((pIncludes->Item(_variant_t(Name), &pInclude) == S_OK) && (pInclude != NULL))
	{
		return pInclude;
	}
	else
	{
		return CComPtr<EnvDTE::CodeElement>();
	}
}

CComPtr<EnvDTE::CodeElement> FindInclude(VSClass* pClass, LPCWSTR Name, bool bStdAfx /* = false */)
{
	CComPtr<EnvDTE::Project> pProject;
	CComPtr<EnvDTE::ProjectItem> pProjectItem;
	CComPtr<EnvDTE::CodeElement> pInclude;

	pClass->pElement->get_ProjectItem(&pProjectItem);
	pInclude = FindInclude(pProjectItem, Name);
	if (pInclude != NULL)
		return pInclude;
	if (bStdAfx)
	{
		pProjectItem->get_ContainingProject(&pProject);
		CComPtr<EnvDTE::ProjectItem> pStdAfxFile = FindItem(pProject, _bstr_t(L"stdafx.h"), CComPtr<EnvDTE::ProjectItem>(NULL));
		return FindInclude(pStdAfxFile, Name);
	}
	return pInclude;
}

void RemoveCommands()
{
	HRESULT hr = E_FAIL;
	CComPtr<EnvDTE::_DTE> pDTE;
	pDTE.CoCreateInstance(L"VisualStudio.DTE.8.0");
	
	if (pDTE == NULL)
		return;

	CComPtr<EnvDTE::Commands> pCommands;
	UICommandBarsPtr pCommandBars;
	UICommandBarPtr pHelperCmdBar, pHostCmdBar;
	CComPtr<EnvDTE::Command> pCmd;

	pDTE->get_Commands(&pCommands);
	
#if defined(_FOR_VS2008)
	pDTE->get_CommandBars((IDispatch**)&pCommandBars);
#elif defined(_FOR_VS2005)
	pDTE->get_CommandBars((IDispatch**)&pCommandBars);
#else
	pDTE->get_CommandBars(&pCommandBars);
#endif

	if (pCommandBars != NULL)
	{
		pCommandBars->get_Item(_variant_t(L"Tools"), &pHostCmdBar);
		if (pHostCmdBar != NULL)
		{
			
			UICommandBarControlsPtr pControls;
			pHostCmdBar->get_Controls(&pControls);
			if (pControls != NULL)
			{
				UICommandBarControlPtr pControl;
				pControls->get_Item(_variant_t(L"WTL Helper"), &pControl);
				if (pControl != NULL)
				{
					UICommandBarPopupControlPtr pPopup; 
					hr = pControl->QueryInterface(&pPopup);
					if (pPopup != NULL)
					{
						pPopup->get_CommandBar(&pHelperCmdBar);
						pCommands->RemoveCommandBar(pHelperCmdBar);
					}
				}
			}
		}
		
		/*pCommandBars->get_Item(_variant_t(L"Class View Item.WTL Helper"), &pHelperCmdBar);
		if (pHelperCmdBar != NULL)
		{
			pCommands->RemoveCommandBar(pHelperCmdBar);
		}

		pCommandBars->get_Item(_variant_t(L"Resource Editors.WTL Helper"), &pHelperCmdBar);
		if (pHelperCmdBar != NULL)
		{
			pCommands->RemoveCommandBar(pHelperCmdBar);
		}

		pCommandBars->get_Item(_variant_t(L"Menu Designer.WTL Helper"), &pHelperCmdBar);
		if (pHelperCmdBar != NULL)
		{
			pCommands->RemoveCommandBar(pHelperCmdBar);
		}

		pCommandBars->get_Item(_variant_t(L"Resource View.WTL Helper"), &pHelperCmdBar);
		if (pHelperCmdBar != NULL)
		{
			pCommands->RemoveCommandBar(pHelperCmdBar);
		}*/

		/*if (pCommandBars->get_Item(_variant_t(L"Class View Item"), &pClassViewBar) == S_OK)
		{
			pCommands->AddCommandBar(_bstr_t(L"WTL Helper"), EnvDTE::vsCommandBarTypeMenu, pClassViewBar, 10, (IDispatch**)&m_pClassViewAddinBar);
		}
		if (pCommandBars->get_Item(_variant_t(L"Resource Editors"), &pResourceBar) == S_OK)
		{
			pCommands->AddCommandBar(_bstr_t(L"WTL Helper"), EnvDTE::vsCommandBarTypeMenu, pResourceBar, 1, (IDispatch**)&m_pResViewAddinBar);
		}
		if (pCommandBars->get_Item(_variant_t(L"Menu Designer"), &pMenuDesigner) == S_OK)
		{
			pCommands->AddCommandBar(_bstr_t(L"WTL Helper"), EnvDTE::vsCommandBarTypeMenu, pMenuDesigner, 1, (IDispatch**)&m_pMenuDesignerBar);
		}
		if (pCommandBars->get_Item(_variant_t(L"Resource View"), &pResView) == S_OK)
		{
			pCommands->AddCommandBar(_bstr_t(L"WTL Helper"), EnvDTE::vsCommandBarTypeMenu, pResView, 1, (IDispatch**)&m_pResViewContextBar);
		}*/
	}

	CString strBase = AddinString();
	for(size_t i =0; i < sizeof(WtlHelperCommands)/sizeof(const CommandStruct*); i++)
	{
		CString str = strBase + (LPCTSTR)WtlHelperCommands[i]->Name;
		pCommands->Item(_variant_t(str), -1, &pCmd);
		if (pCmd != NULL)
		{
			pCmd->Delete();
		}
	}

	pDTE->Quit();
}