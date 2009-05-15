////////////////////////////////////////////////////////////////////////////////
//
// Copyright (c) 2004 Sergey Solozhentsev
// Author: 	Sergey Solozhentsev e-mail: salos@mail.ru
// Product:	WTL Helper
// File:      	InsertPoint.cpp
// Created:	17.01.2005 9:58
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

#include "stdafx.h"
#include "InsertPoint.h"

//////////////////////////////////////////////////////////////////////////
InsertPointFunc::InsertPointFunc(int iType) : InsertionPoint(iType)
{
}

HRESULT InsertPointFunc::Insert(VSClass* pClass, int Step)
{
	if (Step == INSERT_STEP_ENVDTE)
	{
		return pClass->InsertFunction((VSFunction*)pElement, Body);
	}
	return S_OK;
}

//////////////////////////////////////////////////////////////////////////
InsertPointVariable::InsertPointVariable(int iType) : InsertionPoint(iType)
{
}

HRESULT InsertPointVariable::Insert(VSClass* pClass, int Step)
{
	if (Step == INSERT_STEP_ENVDTE)
	{
		return pClass->InsertVariable((VSVariable*)pElement);
	}
	return S_OK;
}

//////////////////////////////////////////////////////////////////////////
InsertPointMap::InsertPointMap(int iType) : InsertionPoint(iType)
{
}

HRESULT InsertPointMap::Insert(VSClass* pClass, int Step)
{
	if (Step == INSERT_STEP_GLOBAL)
	{
		return pClass->InsertMap((VSMap*)pElement);
	}
	return S_OK;
}

//////////////////////////////////////////////////////////////////////////
InsertPointAltMap::InsertPointAltMap(int iType) : InsertionPoint(iType)
{
}

HRESULT InsertPointAltMap::Insert(VSClass* pClass, int Step)
{
	HRESULT hr = E_FAIL;
	VSMessageMap* pMesMap = (VSMessageMap*)pElement;
	if (Step == INSERT_STEP_ENVDTE && !DefineName.IsEmpty())
	{
		CComPtr<EnvDTE::ProjectItem> pProjItem;
		pClass->pElement->get_ProjectItem(&pProjItem);
		if (pProjItem != NULL)
		{
			CComPtr<EnvDTE::FileCodeModel> pCodeModel;
			pProjItem->get_FileCodeModel(&pCodeModel);
			if (pCodeModel != NULL)
			{
				CComPtr<VCCodeModelLibrary::VCFileCodeModel> pVCCodeModel;
				hr = pCodeModel->QueryInterface(&pVCCodeModel);
				if (pVCCodeModel != NULL)
				{
					CComPtr<VCCodeModelLibrary::VCCodeMacro> spDummy;
					pVCCodeModel->AddMacro(_bstr_t(pElement->Name), _bstr_t(DefineName), vtMissing, &spDummy);
				}
			}
		}
	}
	if (Step == INSERT_STEP_GLOBAL)
	{
		pParentMap->InsertAltMap(pMesMap);
	}
	return S_OK;
}

//////////////////////////////////////////////////////////////////////////
InsertPointMapEntry::InsertPointMapEntry(int iType) : InsertionPoint(iType)
{
}

HRESULT InsertPointMapEntry::Insert(VSClass* pClass, int Step)
{
	if (Step == INSERT_STEP_MAP_ENTRY)
	{
		return pParentMap->InsertMapEntry((VSMapEntry*)pElement);
	}
	return S_OK;
}
//////////////////////////////////////////////////////////////////////////
InsertPointHandler::InsertPointHandler(int iType) : InsertPointMapEntry(iType)
{
}

HRESULT InsertPointHandler::Insert(VSClass* pClass, int Step)
{
	if (Step == INSERT_STEP_ENVDTE && pFunction && (pFunction->pElement == NULL))
	{
		return pClass->InsertFunction(pFunction, Body);
	}
	if (Step == INSERT_STEP_MAP_ENTRY)
	{
		return pParentMap->InsertMapEntry((VSMapEntry*)pElement);
	}
	return S_OK;
}

//////////////////////////////////////////////////////////////////////////
InsertPointDDX::InsertPointDDX(int iType) : InsertPointMapEntry(iType)
{
}

HRESULT InsertPointDDX::Insert(VSClass* pClass, int Step)
{
	if (Step == INSERT_STEP_MAP_ENTRY)
	{
		return pParentMap->InsertMapEntry((VSMapEntry*)pElement);
	}
	if (Step == INSERT_STEP_ENVDTE)
	{
		HRESULT hr = pClass->InsertVariable(pVariable);
		if (hr != S_OK) 
		{
			return hr;
		}
		if (!Initializer.IsEmpty())
		{
			bool bFindConstructor = false;
			for (size_t i = 0; i < pClass->Functions.GetCount(); i++)
			{
				VSFunction* pFunc = pClass->Functions[i];
				if (pFunc->Kind == EnvDTE::vsCMFunctionConstructor)
				{
					if ((pClass->Functions[i]->pElement != NULL))
					{
						bFindConstructor = true;
						hr = AddInitializer(pClass->Functions[i]);
						if (FAILED(hr))
							return hr;
					}
				}
			}
			if (!bFindConstructor)
			{
				VSFunction* pFunc = new VSFunction;
				pFunc->pParent = pClass;
				pFunc->Access = EnvDTE::vsCMAccessPublic;
				pClass->Functions.Add(pFunc);
				CComPtr<EnvDTE::CodeFunction> pCodeFunc;
				_bstr_t ClassName;
				pClass->pElement->get_Name(ClassName.GetAddress());
				pFunc->Name = (LPCTSTR)ClassName;
				pClass->InsertFunction(pFunc, CString(), EnvDTE::vsCMFunctionConstructor, _variant_t(0));
				if (pFunc->pElement != NULL)
				{
					pFunc->RetriveType();
					return AddInitializer(pFunc);
				}
				return S_OK;
			}
		}
	}
	return S_OK;
}

HRESULT InsertPointDDX::AddInitializer(VSFunction* pFunc)
{
	HRESULT hr = E_FAIL;
	CComPtr<VCCodeModelLibrary::VCCodeFunction> pVCFunc;
	hr = pFunc->pElement->QueryInterface(&pVCFunc);
	ATLASSERT(pVCFunc != NULL);
	CString InitString;
	InitString.Format(_T("%s(%s)"), pVariable->Name, Initializer);

	return pVCFunc->AddInitializer(_bstr_t(InitString));
}

//////////////////////////////////////////////////////////////////////////
InsertSpecFunction::InsertSpecFunction(int iType /* = INSERT_POINT_SPEC_FUNC */) :
InsertPointFunc(iType)
{
}

HRESULT InsertSpecFunction::Insert(VSClass* pClass, int Step)
{
	HRESULT hr = E_FAIL;
	if (Step == INSERT_STEP_ENVDTE)
	{
		HRESULT hr = pClass->InsertFunction((VSFunction*)pElement, Body);
		if (hr != S_OK)
			return hr;
		if (pBase)
			return pClass->InsertBase(pBase);
		else
			return S_OK;
	}
	if (Step == INSERT_STEP_GLOBAL)
	{
		if (!OnCreateBody.IsEmpty())
		{
			VSMessageMap* pMap = (VSMessageMap*)pClass->GetMap(_T("MSG"));
			if (!pMap)
				return E_FAIL;
			CAtlArray<VSMapEntry*> MapEntries;
			CString CreateMessage;
			if (bInitDialog)
			{
				CreateMessage = _T("WM_INITDIALOG");
			}
			else
			{
				CreateMessage = _T("WM_CREATE");
			}
			VSMessageMap::FindMapEntryByMessage(CreateMessage, pMap, MapEntries);
			for (size_t i = 0; i < pMap->AltMaps.GetCount(); i++)
			{
				VSMessageMap::FindMapEntryByMessage(CreateMessage, pMap->AltMaps[i], MapEntries);
			}
			if (MapEntries.GetCount())
			{
				VSFunction* pFunc;
				for (size_t i = 0; i < MapEntries.GetCount(); i++)
				{
					pFunc = (VSFunction*)MapEntries[i]->pData;
					CComPtr<VCCodeModelLibrary::VCCodeFunction> pVCFunc;
					hr = pFunc->pElement->QueryInterface(&pVCFunc);
					ATLASSERT(pVCFunc != NULL);
					_bstr_t FuncBody;
					pVCFunc->get_BodyText(FuncBody.GetAddress());
					FuncBody = _bstr_t(OnCreateBody) + FuncBody;
					pVCFunc->put_BodyText(FuncBody);
				}
			}
		}
		if (!OnDestroyBody.IsEmpty())
		{
			VSMessageMap* pMap = (VSMessageMap*)pClass->GetMap(_T("MSG"));
			if (!pMap)
				return E_FAIL;
			CAtlArray<VSMapEntry*> MapEntries;
			VSMessageMap::FindMapEntryByMessage(_T("WM_DESTROY"), pMap, MapEntries);
			for (size_t i = 0; i < pMap->AltMaps.GetCount(); i++)
			{
				VSMessageMap::FindMapEntryByMessage(_T("WM_DESTROY"), pMap->AltMaps[i], MapEntries);
			}
			if (MapEntries.GetCount())
			{
				VSFunction* pFunc;
				for (size_t i = 0; i < MapEntries.GetCount(); i++)
				{
					pFunc = (VSFunction*)MapEntries[i]->pData;
					CComPtr<VCCodeModelLibrary::VCCodeFunction> pVCFunc;
					hr = pFunc->pElement->QueryInterface(&pVCFunc);
					ATLASSERT(pVCFunc != NULL);
					_bstr_t FuncBody;
					pVCFunc->get_BodyText(FuncBody.GetAddress());
					FuncBody = _bstr_t(OnDestroyBody) + FuncBody;
					pVCFunc->put_BodyText(FuncBody);
				}
			}
		}
		return S_OK;
	}
	return S_OK;
}

//////////////////////////////////////////////////////////////////////////
InsertInclude::InsertInclude(int iType /* = INSERT_POINT_INCLUDE */):
InsertionPoint(iType), Pos(-1)
{
}

InsertInclude::~InsertInclude()
{
	if (pElement)
	{
		delete pElement;
	}
}

HRESULT InsertInclude::Insert(VSClass* pClass, int Step)
{
	HRESULT hr = E_FAIL;
	if (Step == INSERT_STEP_ENVDTE)
	{
		CComPtr<EnvDTE::ProjectItem> pProjectItem;
		if (pProjectFile != NULL)
		{
			pProjectItem = pProjectFile;
		}
		else
		{
			pClass->pElement->get_ProjectItem(&pProjectItem);
		}
		CComPtr<EnvDTE::FileCodeModel> pFileCodeModel;
		pProjectItem->get_FileCodeModel(&pFileCodeModel);
		CComPtr<VCCodeModelLibrary::VCFileCodeModel> pVCFileCodeModel;
		hr = pFileCodeModel->QueryInterface(&pVCFileCodeModel);
		
		CComPtr<VCCodeModelLibrary::VCCodeInclude> spDummy;
		pInclude = pVCFileCodeModel->AddInclude(_bstr_t(pElement->Name), Pos, &spDummy);
	}
	if (Step == INSERT_STEP_GLOBAL)
	{
		if (pInclude != NULL && !AdditionalMacro.IsEmpty())
		{
			CComPtr<EnvDTE::TextPoint> pTextPoint;
			CComPtr<EnvDTE::EditPoint> pEditPoint;
			pInclude->get_StartPoint(&pTextPoint);
			pTextPoint->CreateEditPoint(&pEditPoint);
			_bstr_t MacroText = CString(_T("#define ")) + AdditionalMacro + _T("\r\n");
			pEditPoint->Insert(MacroText);
		}
	}
	return S_OK;
}

//////////////////////////////////////////////////////////////////////////
InsertPointDDXSupport::InsertPointDDXSupport(int Type /* = INSERT_POINT_DDXSUPPORT */):
InsertionPoint(Type)
{
}

InsertPointDDXSupport::~InsertPointDDXSupport()
{
	if (pElement)
	{
		delete pElement;
	}
}

HRESULT InsertPointDDXSupport::Insert(VSClass* pClass, int Step)
{
	HRESULT hr = E_FAIL;
	if (Step == INSERT_STEP_ENVDTE)
	{
		CComPtr<EnvDTE::ProjectItem> pProjectItem;
		pClass->pElement->get_ProjectItem(&pProjectItem);
		CComPtr<EnvDTE::FileCodeModel> pFileCodeModel;
		pProjectItem->get_FileCodeModel(&pFileCodeModel);
		CComPtr<VCCodeModelLibrary::VCFileCodeModel> pVCFileCodeModel;
		hr = pFileCodeModel->QueryInterface(&pVCFileCodeModel);
		ATLASSERT(pVCFileCodeModel != NULL);
		CComPtr<EnvDTE::CodeElement> pElem;
		if (pElement)
		{
			ATLASSERT(pElement->ElementType == EnvDTE::vsCMElementIncludeStmt);
			CComPtr<VCCodeModelLibrary::VCCodeInclude> spTmp;
			hr = pVCFileCodeModel->AddInclude(_bstr_t(pElement->Name), CComVariant(), &spTmp);
			hr = spTmp->QueryInterface(&pElement->pElement);
			pElem = pElement->pElement;
		}
		else
		{
			CComPtr<EnvDTE::CodeElements> pIncludes;
			hr = pVCFileCodeModel->get_Includes(&pIncludes);
			HRESULT hr = pIncludes->Item(_variant_t(L"<atlddx.h>"), &pElem);
			if (FAILED(hr))
				return hr;
		}
		if (bUseFloat)
		{
			ATLASSERT(pElem != NULL);
			CComPtr<EnvDTE::TextPoint> pTextPoint;
			pElem->GetStartPoint(EnvDTE::vsCMPartWholeWithAttributes, &pTextPoint);
			ATLASSERT(pTextPoint != NULL);
			pTextPoint->CreateEditPoint(&pFloatPoint);
		}
		if (pBase)
		{
			CComPtr<VCCodeModelLibrary::VCCodeClass> pVCClass;
			hr = pClass->pElement->QueryInterface(&pVCClass);
			ATLASSERT(pVCClass != NULL);
			HRESULT hr = pVCClass->AddBase(_variant_t(pBase->Name), _variant_t(-1), &pBase->pElement);

			return hr;
		}
	}
	if (Step == INSERT_STEP_GLOBAL)
	{
		if (bUseFloat)
		{
			if (pFloatPoint != NULL)
			{
				pFloatPoint->CharLeft(1);
				_bstr_t Macro = L"\r\n#define _ATL_USE_DDX_FLOAT";
				pFloatPoint->Insert(Macro);
			}
			else
			{
				return E_INVALIDARG;
			}
		}
	}
	return S_OK;
}
//////////////////////////////////////////////////////////////////////////
InsertPointReplaceEndMap::InsertPointReplaceEndMap() : InsertionPoint(INSERT_POINT_REPLACE_END_MSG_MAP)
{

}

HRESULT InsertPointReplaceEndMap::Insert(VSClass* pClass, int Step)
{
	if (Step == INSERT_STEP_GLOBAL)
	{
		VSMessageMap* pMesMap = (VSMessageMap*)pElement;
		return pMesMap->ReplacePrefixEnd();
	}
	return S_OK;
}