/********************************************************************
* CreatedOn: 2005-8-18   17:06
* FileName: KSCOMDll.h
* CreatedBy: qiuruifeng <qiuruifeng@kingsoft.net>
* $LastChangedDate$
* $LastChangedRevision$
* $LastChangedBy$
* $HeadURL$
* Purpose:
*********************************************************************/
#pragma once 

#include <Guiddef.h>
#include <Windows.h>

typedef struct _CLASSINFO
{
	CLSID m_CLSID;
	const char *m_pszProgID;
	DWORD m_dwProperty;
}CLASSINFO;

typedef HRESULT (__stdcall _SCOM_KSCOMDllGETCLASSOBJECT)(const CLSID& clsid, const IID &riid , void **ppv);
typedef HRESULT (__stdcall _SCOM_KSCOMDllGETCLASSCOUNT)(int &nReturnSize);
typedef HRESULT (__stdcall _SCOM_KSCOMDllGETCLASSINFO)(CLASSINFO *ClassInfo, int nInSize);
typedef HRESULT (__stdcall _SCOM_KSCOMDllCANUNLOADNOW)(void);

class KSCOMDll
{
public:
	KSCOMDll():m_hModule(0),
		  m_pfnDllGetClassObject(0),
		  m_pfnDllGetClassCount(0),
		  m_pfnDllGetClassInfo(0),
		  m_pfnDllCanUnloadNow(0)
	  {
	  }

	  ~KSCOMDll()
	  {
		  if (m_hModule)
			  FreeLibrary(m_hModule);
	  }

	HRESULT Open(const char* pszModulePath);
	HRESULT Open(const wchar_t* pszModulePath);
	HRESULT Release();
	HRESULT GetClassObject(const CLSID& clsid, const IID& riid, void** ppv) const;
	HRESULT GetClassCount(int &nReturnSize) const;
	HRESULT GetClassInfo(CLASSINFO *ClassInfo, int nInSize) const;
	HRESULT CanUnloadNow(void) const;
	HMODULE GetModuleHandle(void);
	void Swap(KSCOMDll& Other);

private:
	HRESULT GetFuncPtr(HMODULE hModule);
private:
	HMODULE m_hModule;
	_SCOM_KSCOMDllGETCLASSOBJECT*	m_pfnDllGetClassObject;
	_SCOM_KSCOMDllGETCLASSCOUNT*	m_pfnDllGetClassCount;
	_SCOM_KSCOMDllGETCLASSINFO*	m_pfnDllGetClassInfo;
	_SCOM_KSCOMDllCANUNLOADNOW*	m_pfnDllCanUnloadNow;
};



inline HRESULT KSCOMDll::Release()
{	
	if (m_hModule == NULL)
	{
		return S_OK;
	}
	FreeLibrary(m_hModule);
	m_hModule				= NULL;
	m_pfnDllGetClassObject	= NULL;
	m_pfnDllGetClassCount	= NULL;
	m_pfnDllGetClassInfo	= NULL;
	m_pfnDllCanUnloadNow	= NULL;
	return S_OK;	
}

inline HRESULT KSCOMDll::GetClassObject(const CLSID& clsid, const IID& riid,void** ppv) const
{
	return m_pfnDllGetClassObject(clsid, riid, ppv);
}

inline HRESULT KSCOMDll::GetClassCount(int &nReturnSize) const
{
	return m_pfnDllGetClassCount(nReturnSize);
}

inline HRESULT KSCOMDll::GetClassInfo(CLASSINFO *ClassInfo, int nInSize) const
{
	return m_pfnDllGetClassInfo(ClassInfo, nInSize);
}

inline HRESULT KSCOMDll::CanUnloadNow(void) const
{
	return m_pfnDllCanUnloadNow();
}

inline HMODULE KSCOMDll::GetModuleHandle(void)
{
	return m_hModule;
}

inline void KSCOMDll::Swap(KSCOMDll& Other)
{
	HMODULE tmp_hModule = m_hModule;
	_SCOM_KSCOMDllGETCLASSOBJECT*	tmp_pfnDllGetClassObject= m_pfnDllGetClassObject;
	_SCOM_KSCOMDllGETCLASSCOUNT*	tmp_pfnDllGetClassCount	= m_pfnDllGetClassCount;
	_SCOM_KSCOMDllGETCLASSINFO*		tmp_pfnDllGetClassInfo	= m_pfnDllGetClassInfo;
	_SCOM_KSCOMDllCANUNLOADNOW*		tmp_pfnDllCanUnloadNow	= m_pfnDllCanUnloadNow;

	m_hModule				= Other.m_hModule;
	m_pfnDllGetClassObject	= Other.m_pfnDllGetClassObject;
	m_pfnDllGetClassCount	= Other.m_pfnDllGetClassCount;
	m_pfnDllGetClassInfo	= Other.m_pfnDllGetClassInfo;
	m_pfnDllCanUnloadNow	= Other.m_pfnDllCanUnloadNow;

	Other.m_hModule					= tmp_hModule;
	Other.m_pfnDllGetClassObject	= tmp_pfnDllGetClassObject;
	Other.m_pfnDllGetClassCount		= tmp_pfnDllGetClassCount;
	Other.m_pfnDllGetClassInfo		= tmp_pfnDllGetClassInfo;
	Other.m_pfnDllCanUnloadNow		= tmp_pfnDllCanUnloadNow;
}
/*
inline const char* KSCOMDll::GetModulePath(void)
{
return m_strModulePath.c_str();
}
*/
inline HRESULT KSCOMDll::Open(const char* pszModulePath)
{
	if (m_hModule != NULL)
	{
		return ERROR_ALREADY_EXISTS;
	}
	m_hModule = ::LoadLibraryA(pszModulePath);
	
	return GetFuncPtr(m_hModule);
}

inline HRESULT KSCOMDll::Open(const wchar_t* pszModulePath)
{
	if (m_hModule != NULL)
	{
		return ERROR_ALREADY_EXISTS;
	}
	m_hModule = ::LoadLibraryW(pszModulePath);

	return GetFuncPtr(m_hModule);
}

inline HRESULT KSCOMDll::GetFuncPtr(HMODULE hModule)
{
	HRESULT ret = S_OK;

	if (hModule == NULL)
	{
		//#ifdef WIN32
		//		return GetLastError();
		//#else
		//		return ERROR_MOD_NOT_FOUND;
		//#endif
		ret = E_FAIL;//ERROR_MOD_NOT_FOUND;
		//return CO_E_DLLNOTFOUND;
	}
	else if (NULL == (m_pfnDllGetClassObject = (_SCOM_KSCOMDllGETCLASSOBJECT*)GetProcAddress(
		hModule,
		"KSDllGetClassObject")))
	{
		ret = E_FAIL;//E_SCOM_PROC_NOT_FOUND;
	}
	else if (NULL == (m_pfnDllGetClassCount = (_SCOM_KSCOMDllGETCLASSCOUNT*)GetProcAddress(
		hModule, 
		"KSDllGetClassCount")))
	{
		ret = E_FAIL;//E_SCOM_PROC_NOT_FOUND;
	}
	else if (NULL == (m_pfnDllGetClassInfo = (_SCOM_KSCOMDllGETCLASSINFO*)GetProcAddress(
		hModule, 
		"KSDllGetClassInfo")))
	{
		ret = E_FAIL;//E_SCOM_PROC_NOT_FOUND;
	}
	else if (NULL == (m_pfnDllCanUnloadNow = (_SCOM_KSCOMDllCANUNLOADNOW*)GetProcAddress(
		hModule, 
		"KSDllCanUnloadNow")))
	{
		ret = E_FAIL;//E_SCOM_PROC_NOT_FOUND;
	}

	return ret;

}
