#pragma once


class CCustomProjectSettings
{
	CComPtr<EnvDTE::_DTE> m_pDte;
	CComPtr<EnvDTE::Globals> GetSolutionGlobals();
	CComPtr<EnvDTE::Globals> GetProjectGlobals(CComPtr<EnvDTE::Project> pProject);
	bool GetGlobalsValue(CComPtr<EnvDTE::Globals> pGlobals, IN LPCTSTR lpszlpszVariableName, OUT _variant_t& vtValue);
	bool SetGlobalsValue(CComPtr<EnvDTE::Globals> pGlobals, IN LPCTSTR lpszlpszVariableName, IN _variant_t& vtValue);
	bool DeleteGlobalsVariable(CComPtr<EnvDTE::Globals> pGlobals, IN LPCTSTR lpszlpszVariableName);
public:
	CCustomProjectSettings(void);
	~CCustomProjectSettings(void);

	void Init(EnvDTE::_DTE* dte);
	void Clearup(void);

	bool GetSolutionVariableValue(IN LPCTSTR lpszVariableName, OUT _variant_t& vtValue);
	bool SetSolutionVariableValue(IN LPCTSTR lpszVariableName, IN _variant_t& vtValue);
	bool DeleteSolutionVariable(IN LPCTSTR lpszVariableName);

	bool GetProjectVariableValue(IN CComPtr<EnvDTE::Project> pProject, IN LPCTSTR lpszVariableName, OUT _variant_t& vtValue);
	bool SetProjectVariableValue(IN CComPtr<EnvDTE::Project> pProject, IN LPCTSTR lpszVariableName, IN _variant_t& vtValue);
	bool DeleteProjectVariable(IN CComPtr<EnvDTE::Project> pProject, IN LPCTSTR lpszVariableName);
	
	static CCustomProjectSettings m_sCustomProperties;
};
