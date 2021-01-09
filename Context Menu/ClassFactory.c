// classfactoy.c

#include <Windows.h>
#include <shlobj.h>
#include "ContextMenu.h"
#include "classfactory.h"

// declared in Main.c
extern UINT	g_uiRefThisDll;		// Reference count for this DLL
extern HINSTANCE g_hInstance;	// Reference count for this DLL

#include <stdio.h>

// The virtual table for the ClassFactory
IClassFactoryVtbl iclassFactoryVtbl = {
	CClassFactory_QueryInterface,
	CClassFactory_AddRef,
	CClassFactory_Release,
	CClassFactory_CreateInstance,
	CClassFactory_LockServer
};

//--------------------------------------------------------------
// ClassFactoryEx constructor
//--------------------------------------------------------------
IClassFactory * CClassFactory_Create(void)
{
	 // Create the ClassFactoryStruct that will contain interfaces and vars
	ClassFactoryStruct * pCF = malloc(sizeof(ClassFactoryStruct));
	if(!pCF)
		return NULL;

	pCF->fc.lpVtbl = &iclassFactoryVtbl;

	// init the vars
	pCF->m_hInstance  = g_hInstance;	// Instance handle for this DLL
	pCF->m_ulRef = 1;					// increment the reference

	g_uiRefThisDll++;

	// Return the IClassFactory virtual table
	return &pCF->fc;
}

STDMETHODIMP CClassFactory_QueryInterface(IClassFactory *this, REFIID riid, LPVOID *ppv)
{
	// The address of the struct is the same as the address
	// of the IClassFactory Virtual table. 
	ClassFactoryStruct * pCF = (ClassFactoryStruct*)this;

	if (IsEqualIID (riid, &IID_IUnknown) || IsEqualIID (riid, &IID_IClassFactory))
	{
		//MessageBox(NULL, "facQuery", NULL, MB_OK + MB_ICONSTOP);
		*ppv = this;
		pCF->m_ulRef++;
		return NOERROR;
	}
	else
	{
		*ppv = NULL;
		return ResultFromScode (E_NOINTERFACE);
	}
}

STDMETHODIMP_(ULONG) CClassFactory_AddRef(IClassFactory *this)
{
	ClassFactoryStruct * pCF = (ClassFactoryStruct*)this;
	return ++pCF->m_ulRef;
}

STDMETHODIMP_(ULONG) CClassFactory_Release(IClassFactory *this)
{
	ClassFactoryStruct * pCF = (ClassFactoryStruct*)this;
	if (--pCF->m_ulRef == 0)
	{
		free(this);
		g_uiRefThisDll--;
		return 0;
	}
	return pCF->m_ulRef;
}

STDMETHODIMP CClassFactory_CreateInstance(IClassFactory *this, LPUNKNOWN pUnkOuter, REFIID riid,  LPVOID *ppv)
{
    *ppv = NULL;
	ClassFactoryStruct * pCF = (ClassFactoryStruct*)this;
	
	if (pUnkOuter)
        return ResultFromScode (CLASS_E_NOAGGREGATION);

	//MessageBox(NULL, "facCreateInstance", NULL, MB_OK + MB_ICONSTOP);
	//if (IsEqualIID (riid, &IID_IShellExtInit))
	//{
		// Creates the IContextMenu incorperating IShellExtInit interfaces
		IContextMenu * pIContextMenu = CContextMenuExt_Create();

		if (NULL == pIContextMenu)
		{
			return E_OUTOFMEMORY;
		}

		// This puts the IContextMenu interface into 'ppv'
		HRESULT hr = pIContextMenu->lpVtbl->QueryInterface(pIContextMenu, riid, ppv);
		pIContextMenu->lpVtbl->Release(pIContextMenu);
		return hr;
	//} else if (IsEqualIID (riid, &IID_IContextMenu)) {
	//	*ppv = (IContextMenu *)this;
	//	pCF->m_ulRef++;
	//	return NOERROR;
	//}
	//return E_NOTIMPL;
}

STDMETHODIMP CClassFactory_LockServer(IClassFactory *this, BOOL fLock)
{
	MessageBox(NULL, "facLock", NULL, MB_OK + MB_ICONSTOP);
	return ResultFromScode(E_NOTIMPL);
}



