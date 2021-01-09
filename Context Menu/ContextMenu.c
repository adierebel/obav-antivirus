// contextmenu.c

/*--------------------------------------------------------------
	The IShellExtInit interface is incorperated into the 
	IContextMenu interface
--------------------------------------------------------------*/
#include <Windows.h>
#include <shlobj.h>
#include <wchar.h>
#include <tchar.h>
#include "ContextMenu.h"

// Forward declarations
DWORD ReadRegistryString(char *szKey, char *szSubKey, char *pszDefault, char * szReturnBuffer);
int   GetNameFromPath(LPTSTR pPath, char * strBuff, int maxCpy);

// Keep a count of DLL references
extern UINT g_uiRefThisDll;
extern HINSTANCE g_hInstance;
extern char*	g_progfile;

// Command from Shell when FindFile search has been selected
#define ID_SEARCH 0x0001

// The virtual table of functions for IContextMenu interface
IContextMenuVtbl icontextMenuVtbl = {
	CContextMenuExt_QueryInterface,
	CContextMenuExt_AddRef,
	CContextMenuExt_Release,
	CContextMenuExt_QueryContextMenu,
	CContextMenuExt_InvokeCommand,
	CContextMenuExt_GetCommandString
};

// The virtual table of functions for IShellExtInit interface
IShellExtInitVtbl ishellInitExtVtbl = {
    CShellInitExt_QueryInterface,
    CShellInitExt_AddRef,
    CShellInitExt_Release,
    CShellInitExt_Initialize
};

//--------------------------------------------------------------
// IContextMenu constructor
//--------------------------------------------------------------
IContextMenu * CContextMenuExt_Create(void)
{
	//MessageBox(NULL, "Ccreate", NULL, MB_OK + MB_ICONSTOP);

	// Create the ContextMenuExtStruct that will contain interfaces and vars
	ContextMenuExtStruct * pCM = malloc(sizeof(ContextMenuExtStruct));
	if(!pCM)
		return NULL;

	// Point to the IContextMenu and IShellExtInit Vtbl's
	pCM->cm.lpVtbl = &icontextMenuVtbl;
	pCM->si.lpVtbl = &ishellInitExtVtbl;

	// increment the class reference count
	pCM->m_ulRef = 1;
	pCM->m_pszSource = NULL;

	g_uiRefThisDll++;

	// Return the IContextMenu virtual table
	return &pCM->cm;
}

//===============================================
// IContextMenu interface routines
//===============================================
STDMETHODIMP CContextMenuExt_QueryInterface(IContextMenu * this, REFIID riid, LPVOID *ppv)
{
	//MessageBox(NULL, "queryinterface", NULL, MB_OK + MB_ICONSTOP);
	// The address of the struct is the same as the address
	// of the IContextMenu Virtual table. 
	ContextMenuExtStruct * pCM = (ContextMenuExtStruct*)this;
	
    if (IsEqualIID (riid, &IID_IUnknown) || IsEqualIID (riid, &IID_IContextMenu))
	{
        *ppv = this;
        pCM->m_ulRef++;
        return NOERROR;
    }
    else if (IsEqualIID (riid, &IID_IShellExtInit))
	{
		// Give the IShellInitExt interface
		*ppv = &pCM->si;
	    pCM->m_ulRef++;
        return NOERROR;
    }
    else
	{
        *ppv = NULL;
        return ResultFromScode (E_NOINTERFACE);
    }
}

STDMETHODIMP_(ULONG) CContextMenuExt_AddRef(IContextMenu * this)
{
	//MessageBox(NULL, "Caddref", NULL, MB_OK + MB_ICONSTOP);
	ContextMenuExtStruct * pCM = (ContextMenuExtStruct*)this;
    return ++pCM->m_ulRef;
}

STDMETHODIMP_(ULONG) CContextMenuExt_Release(IContextMenu * this)
{
	//MessageBox(NULL, "cRelease", NULL, MB_OK + MB_ICONSTOP);
	ContextMenuExtStruct * pCM = (ContextMenuExtStruct*)this;
    if (--pCM->m_ulRef == 0)
	{
		free(this);
		g_uiRefThisDll--;
		return 0;
	}
    return pCM->m_ulRef;
}

STDMETHODIMP CContextMenuExt_GetCommandString(IContextMenu * this, UINT idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
	//MessageBox(NULL, "cGetcommand", NULL, MB_OK + MB_ICONSTOP);
	HRESULT hr = S_OK;
	switch(uFlags)
	{
	case GCS_HELPTEXTA:
		switch(idCmd)
		{
		case ID_SEARCH:
			lstrcpynA((LPSTR)pszName, "Scan File For Viruses", cchMax);
			hr = NOERROR;
			break;
		default:
			hr = E_NOTIMPL;
		}
		break;
	case GCS_HELPTEXTW:
		switch(idCmd)
		{
		case ID_SEARCH:
			lstrcpynW((LPWSTR)pszName, L"Scan File For Viruses", cchMax);
			hr = NOERROR;
			break;
		default:
			hr = E_NOTIMPL;
		}
		break;
	}
	return hr;
}


STDMETHODIMP CContextMenuExt_QueryContextMenu(IContextMenu * this, HMENU hMenu, UINT uiIndexMenu, UINT idCmdFirst,	UINT idCmdLast, UINT uFlags)
{
	//MessageBox(NULL, "cQuery", NULL, MB_OK + MB_ICONSTOP);
	InsertMenu(hMenu, uiIndexMenu++, MF_SEPARATOR | MF_BYPOSITION, 0, NULL);
	int menunya = uiIndexMenu++;
	HBITMAP image = LoadBitmap(g_hInstance, "#8001");
	HANDLE bittmap = CopyImage(image,IMAGE_BITMAP,13,13,LR_COPYFROMRESOURCE);
	/*
	if(g_hInstance==INVALID_HANDLE_VALUE){
		MessageBoxA(0,"asu2","xx",0);
	}	
	HBITMAP image = LoadBitmap(g_hInstance, "#8001");
	if(image==INVALID_HANDLE_VALUE){
		MessageBoxA(0,"asu","xx",0);
	}
	//MENUITEMINFO* mii = malloc(sizeof(MENUITEMINFO));
	//mii->cbSize 		= sizeof(MENUITEMINFO);
	//mii->fMask	 		= MIIM_ID | MIIM_TYPE | MIIM_STATE;
	//mii->wID			= (idCmdFirst + ID_SEARCH);
	//mii->fType			= MF_STRING;
	//mii->dwTypeData	= "Seken wit";
	//mii->fState			= 0x0;
	//mii->hbmpChecked = bittmap;
	//mii->hbmpUnchecked = bittmap;
	//mii->hbmpItem = bittmap;
	InsertMenuItem(hMenu, menunya,TRUE, mii);
	*/
	InsertMenu(hMenu, menunya, MF_STRING | MF_BYPOSITION, (idCmdFirst + ID_SEARCH), _T("Scan With ObAV"));
	SetMenuItemBitmaps(hMenu, menunya, MF_BITMAP | MF_BYPOSITION, bittmap, bittmap);
	InsertMenu(hMenu, uiIndexMenu++, MF_SEPARATOR | MF_BYPOSITION, 0, NULL);
	#define SEVERITY_SUCCESS 0
	return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, (USHORT)(ID_SEARCH + 1));
}


PROCESS_INFORMATION pi;
STARTUPINFO         si;
STDMETHODIMP CContextMenuExt_InvokeCommand(IContextMenu * this, LPCMINVOKECOMMANDINFO lpici)
{
	//MessageBox(NULL, "cInvoke", NULL, MB_OK + MB_ICONSTOP);
	char name[MAX_PATH];
	char appPath[MAX_PATH];
	HRESULT hr = S_OK;

	ContextMenuExtStruct * pCM = (ContextMenuExtStruct*)this;

	switch(LOWORD(lpici->lpVerb))
	{
	case ID_SEARCH:
	{
			lstrcpy(appPath, g_progfile);
			strcat(appPath, "\\ObAV\\ObavScanner.exe");
			if(GetShortPathName(pCM->m_pszSource, name, MAX_PATH)){
				strcat(appPath, " -Scan ");
				strcat(appPath, name);

				memset(&si, 0, sizeof(si));
				memset(&pi, 0, sizeof(pi));

				//MessageBoxA(0, appPath, "x",0);
				si.cb = sizeof(si);
				CreateProcess(NULL, appPath, NULL, NULL, 1, NORMAL_PRIORITY_CLASS,
						NULL, NULL, &si, &pi);
			}
	}
	break;
	default:
		hr = E_FAIL;
	}
	return hr;
}

void CContextMenuExt_ErrMessage(DWORD dwErrcode)
{
	//MessageBox(NULL, "cERR", NULL, MB_OK + MB_ICONSTOP);
	void* pMsgBuf;
	TCHAR szMessage[MAX_PATH] = {0};

	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
		NULL,
		dwErrcode,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		(LPTSTR)&pMsgBuf,
		0,
		NULL);
	lstrcpy(szMessage, (LPCTSTR)pMsgBuf);

	MessageBox(GetForegroundWindow(), szMessage, "Oops! Error", MB_ICONERROR);
	LocalFree(pMsgBuf);
}

//===============================================
// IShellExtInit interface routines
//===============================================
STDMETHODIMP CShellInitExt_QueryInterface(IShellExtInit * this, REFIID riid, LPVOID *ppv)
{
	//MessageBox(NULL, "cQuery2", NULL, MB_OK + MB_ICONSTOP);
	/*-----------------------------------------------------------------
	IContextMenu Vtbl is the same address as ContextMenuExtStruct.
 	IShellExtInit is sizeof(IContextMenu) further on.
	-----------------------------------------------------------------*/
	ContextMenuExtStruct * pCM = (ContextMenuExtStruct *)(this-1);
			
	if (IsEqualIID (riid, &IID_IUnknown))
	{
		*ppv = (LPUNKNOWN) (IContextMenu *)this;
		pCM->m_ulRef++;
		return NOERROR;
	}
	// Give the IContextMenu interface here
	else if (IsEqualIID (riid, &IID_IContextMenu))
	{
		*ppv = &pCM->cm;
		pCM->m_ulRef++;
		return NOERROR;
	}
	else if (IsEqualIID (riid, &IID_IShellExtInit))
	{
		*ppv = (IContextMenu *)this;
		pCM->m_ulRef++;
		return NOERROR;
	}
	else
	{
		*ppv = NULL;
		return ResultFromScode (E_NOINTERFACE);
	}
}

STDMETHODIMP_(ULONG) CShellInitExt_AddRef(IShellExtInit * this)
{
	//MessageBox(NULL, "cAddref2", NULL, MB_OK + MB_ICONSTOP);
	// Redirect the IShellExtInit's AddRef to the IContextMenu interface
	IContextMenu * pIContextMenu = (IContextMenu *)(this-1);
	return pIContextMenu->lpVtbl->AddRef(pIContextMenu);
}

STDMETHODIMP_(ULONG) CShellInitExt_Release(IShellExtInit * this)
{
	//MessageBox(NULL, "cRelease2", NULL, MB_OK + MB_ICONSTOP);
	// Redirect the IShellExtInit's Release to the IContextMenu interface
	IContextMenu * pIContextMenu = (IContextMenu *)(this-1);
	return pIContextMenu->lpVtbl->Release(pIContextMenu);
}

STDMETHODIMP CShellInitExt_Initialize(IShellExtInit * this, LPCITEMIDLIST pidlFolder, LPDATAOBJECT lpdobj, HKEY hKeyProgID)
{
	//MessageBox(NULL, "cInit", NULL, MB_OK + MB_ICONSTOP);
//	DWORD dwErrcode = 0L;
	FORMATETC   fe;
	STGMEDIUM   stgmed;

	fe.cfFormat   = CF_HDROP;
	fe.ptd        = NULL;
	fe.dwAspect   = DVASPECT_CONTENT;
	fe.lindex     = -1;
	fe.tymed      = TYMED_HGLOBAL;

	ContextMenuExtStruct * pCM = (ContextMenuExtStruct *)(this-1);
	
	// Get the storage medium from the data object.
	HRESULT hr = lpdobj->lpVtbl->GetData(lpdobj, &fe, &stgmed);
	if (SUCCEEDED(hr))
	{
		if(stgmed.hGlobal)
		{
			int iSize = 0;
			LPDROPFILES pDropFiles = (LPDROPFILES)GlobalLock(stgmed.hGlobal);

			LPTSTR pszFiles = NULL, pszTemp = NULL;
			LPWSTR pswFiles = NULL, pswTemp = NULL;

			if (pDropFiles->fWide)
			{
				pswFiles	=	(LPWSTR) ((BYTE*) pDropFiles + pDropFiles->pFiles);
				pswTemp		=	(LPWSTR) ((BYTE*) pDropFiles + pDropFiles->pFiles);
			}
			else
			{
				pszFiles	=	(LPTSTR) pDropFiles + pDropFiles->pFiles;
				pszTemp		=	(LPTSTR) pDropFiles + pDropFiles->pFiles;
			}

			while(pszFiles && *pszFiles || pswFiles && *pswFiles)
			{
				if(pDropFiles->fWide)
				{
					//Get size of first file/folders path
					iSize += WideCharToMultiByte(CP_ACP, 0, pswFiles, -1, NULL, 0, NULL, NULL);

					pswFiles += (wcslen(pswFiles) + 1 );
				}
				else
				{
					//Get size of first file/folders path
					iSize += strlen(pszFiles) + 1;

					pszFiles += (strlen(pszFiles) +1);
				}
			}
			if(iSize)
			{
				iSize += 2;
				pCM->m_pszSource = malloc(iSize);
				memset(pCM->m_pszSource, 0, iSize);
				if(pswTemp)
				{
					WideCharToMultiByte(CP_ACP,	0, pswTemp,	iSize,	pCM->m_pszSource, iSize, NULL, NULL);
				}
				else
				{
					memcpy(pCM->m_pszSource, pszTemp, iSize);
				}
			}
			goto ende; // only allow one file
		}
ende:
		GlobalUnlock(stgmed.hGlobal);
		ReleaseStgMedium(&stgmed);
	}
//	else
//		dwErrcode = GetLastError();

	return NOERROR;
}

//===============================================
// Helper routines
//===============================================
int GetNameFromPath(LPTSTR pPath, char * strBuff, int maxCpy)
{
	char * pos, * del;
	char * p;
	// Find the last back slash
	del = strrchr(pPath, '\\');
	if(del == NULL){
		MessageBox(NULL, "Can't recognize File Name.", NULL, MB_OK + MB_ICONSTOP);
		return FALSE;
	}
	del++;
	pos = pPath + strlen(pPath);
	p = strBuff;
	int size = 0;
	while( (del < pos) && (size < maxCpy) ){
		*p++ = *del++;
		size++;
	}
	*p = 0; // Terminate it

	return 1;
}

#define KEY HKEY_CURRENT_USER
static char szRegSubKey[] = "Software\\";
static char szRegName[]   = "FindFile\\";
static char szCurrentKey[MAX_PATH+100];
// This gets the FindFile.exe path + name from the registry
DWORD ReadRegistryString(char *szKey, char *szSubKey, char *pszDefault, char * szReturnBuffer)
{
	HKEY  hKey;
	DWORD dwType;

	DWORD size = MAX_PATH+100;

	// make whole key string
	strcpy(szCurrentKey, szRegSubKey);
	strcat(szCurrentKey, szRegName);
	strcat(szCurrentKey, szKey);

	if (ERROR_SUCCESS == RegOpenKeyEx(KEY, szCurrentKey, 0, KEY_QUERY_VALUE, &hKey))
	{
		dwType = REG_SZ;

		if (ERROR_SUCCESS == RegQueryValueEx(hKey, szSubKey, 0, &dwType, (PBYTE)szReturnBuffer, &size)){
			RegCloseKey(hKey);
			return (size);
		}else{
			strcpy(szReturnBuffer, pszDefault);
			return 0;
		}
	}else{
		strcpy(szReturnBuffer, pszDefault);
		return 0;
	}
}



