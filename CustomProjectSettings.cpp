#include "StdAfx.h"
#include "CustomProjectSettings.h"

CCustomProjectSettings CCustomProjectSettings::m_sCustomProperties;

CCustomProjectSettings::CCustomProjectSettings(void)
{
}

CCustomProjectSettings::~CCustomProjectSettings(void)
{
}

EnvDTE::GlobalsPtr CCustomProjectSettings::GetSolutionGlobals()
{
	EnvDTE::GlobalsPtr pGlobals;
	EnvDTE::_SolutionPtr pActiveSolution;
	m_pDte->get_Solution(&pActiveSolution);
	if(pActiveSolution != NULL)
	{
		VARIANT_BOOL bOpen;
		pActiveSolution->get_IsOpen(&bOpen);
		if(bOpen == VARIANT_TRUE)
		{
			pActiveSolution->get_Globals(&pGlobals);
		}
	}
	return pGlobals;
}

EnvDTE::GlobalsPtr CCustomProjectSettings::GetProjectGlobals(EnvDTE::ProjectPtr pProject)
{
	EnvDTE::GlobalsPtr pGlobals;
	pProject->get_Globals(&pGlobals);

	return pGlobals;
}

bool CCustomProjectSettings::GetGlobalsValue(
	EnvDTE::GlobalsPtr pGlobals, IN LPCTSTR lpszVariableName, OUT _variant_t& vtValue)
{
	bool Result = false;
	if(pGlobals != NULL)
	{
		VARIANT_BOOL bVarExists;
		_bstr_t VarName(lpszVariableName);
		pGlobals->get_VariableExists(VarName, &bVarExists);
		if(bVarExists == VARIANT_TRUE)
		{
			pGlobals->get_VariableValue(VarName, &vtValue);
		}
		Result = true;
	}
	return Result;
}

bool CCustomProjectSettings::SetGlobalsValue(
	EnvDTE::GlobalsPtr pGlobals, IN LPCTSTR lpszVariableName, IN _variant_t& vtValue)
{
	bool Result = false;
	if(pGlobals != NULL)
	{
		_bstr_t VarName(lpszVariableName);
		HRESULT hr;
		hr = pGlobals->put_VariableValue(VarName, vtValue);
		hr = pGlobals->put_VariablePersists(VarName, VARIANT_TRUE);
		
		Result = true;
	}
	return Result;
}

bool CCustomProjectSettings::DeleteGlobalsVariable(
	EnvDTE::GlobalsPtr pGlobals, IN LPCTSTR lpszVariableName)
{
	bool Result = false;
	if(pGlobals != NULL)
	{
		_bstr_t VarName(lpszVariableName);
		HRESULT hr = pGlobals->put_VariablePersists(VarName, VARIANT_FALSE);
		pGlobals->put_VariableValue(VarName, _variant_t());
	}
	return Result;
}

void CCustomProjectSettings::Init(EnvDTE::_DTE* dte)
{
	m_pDte = dte;
}

bool CCustomProjectSettings::GetSolutionVariableValue(
	IN LPCTSTR lpszVariableName, OUT _variant_t& vtValue)
{
	
	vtValue.Clear();
	EnvDTE::GlobalsPtr pGlobals = GetSolutionGlobals();
	return GetGlobalsValue(pGlobals, lpszVariableName, vtValue);
}

bool CCustomProjectSettings::SetSolutionVariableValue(
	IN LPCTSTR lpszVariableName, IN _variant_t& vtValue)
{
	EnvDTE::GlobalsPtr pGlobals = GetSolutionGlobals();
	return SetGlobalsValue(pGlobals, lpszVariableName, vtValue);
}

bool CCustomProjectSettings::DeleteSolutionVariable(IN LPCTSTR lpszVariableName)
{
	EnvDTE::GlobalsPtr pGlobals = GetSolutionGlobals();
	return DeleteGlobalsVariable(pGlobals, lpszVariableName);
}

bool CCustomProjectSettings::GetProjectVariableValue(
	IN EnvDTE::ProjectPtr pProject, IN LPCTSTR lpszVariableName, OUT _variant_t& vtValue)
{
	vtValue.Clear();
	EnvDTE::GlobalsPtr pGlobals = GetProjectGlobals(pProject);
	return GetGlobalsValue(pGlobals, lpszVariableName, vtValue);
}

bool CCustomProjectSettings::SetProjectVariableValue(
	IN EnvDTE::ProjectPtr pProject, IN LPCTSTR lpszVariableName, IN _variant_t& vtValue)
{
	EnvDTE::GlobalsPtr pGlobals = GetProjectGlobals(pProject);
	return SetGlobalsValue(pGlobals, lpszVariableName, vtValue);
}

bool CCustomProjectSettings::DeleteProjectVariable(IN EnvDTE::ProjectPtr pProject, IN LPCTSTR lpszVariableName)
{
	EnvDTE::GlobalsPtr pGlobals = GetProjectGlobals(pProject);
	return DeleteGlobalsVariable(pGlobals, lpszVariableName);
}
