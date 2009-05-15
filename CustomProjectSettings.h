#pragma once


class CCustomProjectSettings
{
	EnvDTE::_DTEPtr m_pDte;
	EnvDTE::GlobalsPtr GetSolutionGlobals();
	EnvDTE::GlobalsPtr GetProjectGlobals(EnvDTE::ProjectPtr pProject);
	bool GetGlobalsValue(EnvDTE::GlobalsPtr pGlobals, IN LPCTSTR lpszlpszVariableName, OUT _variant_t& vtValue);
	bool SetGlobalsValue(EnvDTE::GlobalsPtr pGlobals, IN LPCTSTR lpszlpszVariableName, IN _variant_t& vtValue);
	bool DeleteGlobalsVariable(EnvDTE::GlobalsPtr pGlobals, IN LPCTSTR lpszlpszVariableName);
public:
	CCustomProjectSettings(void);
	~CCustomProjectSettings(void);

	void Init(EnvDTE::_DTE* dte);

	bool GetSolutionVariableValue(IN LPCTSTR lpszVariableName, OUT _variant_t& vtValue);
	bool SetSolutionVariableValue(IN LPCTSTR lpszVariableName, IN _variant_t& vtValue);
	bool DeleteSolutionVariable(IN LPCTSTR lpszVariableName);

	bool GetProjectVariableValue(IN EnvDTE::ProjectPtr pProject, IN LPCTSTR lpszVariableName, OUT _variant_t& vtValue);
	bool SetProjectVariableValue(IN EnvDTE::ProjectPtr pProject, IN LPCTSTR lpszVariableName, IN _variant_t& vtValue);
	bool DeleteProjectVariable(IN EnvDTE::ProjectPtr pProject, IN LPCTSTR lpszVariableName);
	
	static CCustomProjectSettings m_sCustomProperties;
};
