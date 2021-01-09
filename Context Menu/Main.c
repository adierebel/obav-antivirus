//
// main.c DLL code including DLLMain()
//
// See the README.txt
//
#define INITGUID
#define COBJMACROS
#include <windows.h>
#include <objbase.h>
#include <initguid.h>
#include <shlobj.h>
#include <tchar.h>
#include "ClassFactory.h"

#include <stdio.h>

/* 
	This CLSID has been generated with GUIDGEN.EXE - if you want to use
	this context menu extension DLL for another purpose you should use 
	GUIDGEN.EXE to create another unique CLSID.
*/
DEFINE_GUID(CLSID_Shell_ContextMenuExt, 
0xA4158F14, 0xE375, 0x4601, 0x80, 0x63, 0x70, 0x4D, 0xD7, 0x11, 0x49, 0x67);
//0x346FD554, 0xEE7E, 0x4f6b, 0x82, 0x40, 0xD0, 0xE7, 0xB7, 0x27, 0x34, 0xBD);

#define SZ_GUID _T("{A4158F14-E375-4601-8063-704DD7114967}")

//#define SZ_GUID _T("{346FD554-EE7E-4f6b-8240-D0E7B72734BD}")

UINT		g_uiRefThisDll = 0;		// Reference count for this DLL
HINSTANCE	g_hInstance;			// Instance handle for this DLL
char*		g_progfile;

BOOL WINAPI DllMain (HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	char teks[MAX_PATH];
	if (dwReason == DLL_PROCESS_ATTACH)
	{
		SHGetSpecialFolderPath( 0, teks, CSIDL_PROGRAM_FILES, 0);
		LocalFree(g_progfile);
		g_progfile = malloc(strlen(teks)+1);
		//invoke LocalAlloc, LMEM_FIXED, eax
		//mov g_progfile,eax
		strcpy(g_progfile, teks);
		g_hInstance = hInstance;
	}
	return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	*ppv = NULL;
	if (!IsEqualCLSID (rclsid, &CLSID_Shell_ContextMenuExt))
	{
		return ResultFromScode (CLASS_E_CLASSNOTAVAILABLE);
	}

	IClassFactory * pClassFactory = CClassFactory_Create();
	if (pClassFactory == NULL)
	{
		return ResultFromScode (E_OUTOFMEMORY);
	}

	HRESULT hr = pClassFactory->lpVtbl->QueryInterface(pClassFactory, riid, ppv);
	pClassFactory->lpVtbl->Release(pClassFactory);
	return hr;
}

STDAPI DllCanUnloadNow (void)
{
	return ResultFromScode((g_uiRefThisDll == 0) ? S_OK : S_FALSE);
}

#define SZ_CLSID				_T("CLSID\\{A4158F14-E375-4601-8063-704DD7114967}")
#define SZ_INPROCSERVER32		_T("CLSID\\{A4158F14-E375-4601-8063-704DD7114967}\\InprocServer32")
#define SZ_DEFAULT				_T("")
#define SZ_THREADINGMODEL		_T("ThreadingModel")
#define SZ_APARTMENT			_T("Apartment")
#define SZ_APPROVED				_T("Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved")
#define SZ_ERRMSG				_T("Unable to add ClassId {A4158F14-E375-4601-8063-704DD7114967} to Registry\nAdministrative Privileges Needed")
#define SZ_ERROR				_T("Error")
#define SZ_DIRCONTEXTMENUEXT	_T("Directory\\shellex\\ContextMenuHandlers\\FindFile")
#define SZ_FILECONTEXTMENUEXT	_T("*\\shellex\\ContextMenuHandlers\\FindFile")
#define SZ_FOLDERCONTEXTMENUEXT	_T("Folder\\shellex\\ContextMenuHandlers\\FindFile")

STDAPI DllRegisterServer(void)
{
	HRESULT hr = E_UNEXPECTED;

	static TCHAR szDescr[] = _T("ContextMenuExt Extension");
		
	TCHAR szFilePath[MAX_PATH];
	GetModuleFileName(g_hInstance, szFilePath, MAX_PATH);

	HKEY hKeyCLSID = NULL;
	if (RegCreateKey(HKEY_CLASSES_ROOT, SZ_CLSID, &hKeyCLSID) != ERROR_SUCCESS)
	{
		return E_UNEXPECTED;
	}
	if (RegSetValueEx(hKeyCLSID, SZ_DEFAULT, 0, REG_SZ,  (const BYTE*)szDescr,
				(lstrlen(szDescr)+1) * sizeof(TCHAR)) != ERROR_SUCCESS)
	{
		return E_UNEXPECTED;
	}
	RegCloseKey(hKeyCLSID);

	HKEY hkeyInprocServer32 = NULL;
	if (RegCreateKey(HKEY_CLASSES_ROOT, SZ_INPROCSERVER32, &hkeyInprocServer32) == ERROR_SUCCESS)
	{
		static TCHAR szApartment[] = SZ_APARTMENT;
		if (RegSetValueEx(hkeyInprocServer32, SZ_DEFAULT, 0, REG_SZ,
				(const BYTE*)szFilePath, (lstrlen(szFilePath)+1) * sizeof(TCHAR)) == ERROR_SUCCESS)
		{
			if (RegSetValueEx(hkeyInprocServer32, SZ_THREADINGMODEL, 0, REG_SZ,
				(const BYTE*)szApartment, (lstrlen(szApartment)+1) * sizeof(TCHAR)) == ERROR_SUCCESS)
			{
				hr = S_OK;
			}
		}
		RegCloseKey(hkeyInprocServer32);
	}

	HKEY hKeyApproved = NULL;
	LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SZ_APPROVED, 0, KEY_SET_VALUE, &hKeyApproved);
	if (lRet == ERROR_ACCESS_DENIED)
	{
		MessageBox(NULL, SZ_ERRMSG, SZ_ERROR, MB_OK);
		hr = E_UNEXPECTED;
	}
	else if (lRet == ERROR_FILE_NOT_FOUND)
	{
		// Desn't exist
	}
	if (hKeyApproved)
	{
		if (RegSetValueEx(hKeyApproved, SZ_GUID, 0, REG_SZ,
			(const BYTE*) szDescr, (lstrlen(szDescr) + 1) * sizeof(TCHAR)) == ERROR_SUCCESS)
		{
			hr = S_OK;
		}
		else
		{
			hr = E_UNEXPECTED;
		}
	}

	HKEY hkeyDirCtx = NULL;
	if (RegCreateKey(HKEY_CLASSES_ROOT, SZ_DIRCONTEXTMENUEXT, &hkeyDirCtx) == ERROR_SUCCESS)
	{
		if (RegSetValueEx(hkeyDirCtx, SZ_DEFAULT, 0, REG_SZ,
				(const BYTE*)SZ_GUID, (lstrlen(SZ_GUID)+1) * sizeof(TCHAR)) == ERROR_SUCCESS)
		{
			hr = S_OK;
		}
			RegCloseKey(hkeyDirCtx);
	}
	HKEY hkeyFileCtx = NULL;
	if (RegCreateKey(HKEY_CLASSES_ROOT, SZ_FILECONTEXTMENUEXT, &hkeyFileCtx) == ERROR_SUCCESS)
	{

		if (RegSetValueEx(hkeyFileCtx, SZ_DEFAULT, 0, REG_SZ,
				(const BYTE*)SZ_GUID, (lstrlen(SZ_GUID)+1) * sizeof(TCHAR)) == ERROR_SUCCESS)
		{
			hr = S_OK;
		}
			RegCloseKey(hkeyFileCtx);
	}
	HKEY hkeyFolderCtx = NULL;
	if (RegCreateKey(HKEY_CLASSES_ROOT, SZ_FOLDERCONTEXTMENUEXT, &hkeyFolderCtx) == ERROR_SUCCESS)
	{
		if (RegSetValueEx(hkeyFolderCtx, SZ_DEFAULT, 0, REG_SZ,
				(const BYTE*)SZ_GUID, (lstrlen(SZ_GUID)+1) * sizeof(TCHAR)) == ERROR_SUCCESS)
		{
			hr = S_OK;
		}
			RegCloseKey(hkeyFolderCtx);
	}

	return hr;
}

STDAPI DllUnregisterServer(void)
{
	HRESULT hr = E_UNEXPECTED;

	RegDeleteKey(HKEY_CLASSES_ROOT, SZ_INPROCSERVER32);

	if (RegDeleteKey(HKEY_CLASSES_ROOT, SZ_CLSID) == ERROR_SUCCESS)
	{
		hr = S_OK;
	}
	HKEY hKeyApproved = NULL;
	LONG lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SZ_APPROVED, 0, KEY_SET_VALUE, &hKeyApproved);
	if (lRet == ERROR_ACCESS_DENIED)
	{
		MessageBox(NULL, SZ_ERRMSG, SZ_ERROR, MB_OK);
		hr = E_UNEXPECTED;
	}
	else if (lRet == ERROR_FILE_NOT_FOUND)
	{
		// Desn't exist
	}
	if (hKeyApproved)
	{
		if (RegDeleteValue(hKeyApproved, SZ_GUID) != ERROR_SUCCESS)
		{
			hr &= E_UNEXPECTED;
		}
		RegCloseKey(hKeyApproved);
	}

	return hr;
}


