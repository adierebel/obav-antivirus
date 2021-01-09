// contextmenu.h

#ifndef _CONTEXTMENU_H_
#define _CONTEXTMENU_H_

// IContextMenu methods
STDMETHODIMP         CContextMenuExt_QueryInterface(IContextMenu *this,REFIID riid, LPVOID *ppvOut);
STDMETHODIMP_(ULONG) CContextMenuExt_AddRef(IContextMenu *this);
STDMETHODIMP_(ULONG) CContextMenuExt_Release(IContextMenu *this);
STDMETHODIMP 		 CContextMenuExt_QueryContextMenu(IContextMenu *,HMENU, UINT, UINT, UINT, UINT);
STDMETHODIMP 		 CContextMenuExt_InvokeCommand(IContextMenu *, LPCMINVOKECOMMANDINFO);
STDMETHODIMP 		 CContextMenuExt_GetCommandString(IContextMenu *,UINT, UINT, UINT *, LPSTR, UINT);

// IContextMenu constructor
IContextMenu * CContextMenuExt_Create(void);

// This struct acts somewhat like a pseudo class in that you have
// variables accociated with an instance of this interface.
typedef struct _ContextMenuExtStruct
{
	// Two interfaces
	IContextMenu   cm;
	IShellExtInit  si;

	// second part of the struct for the variables
	LPTSTR 	m_pszSource;
	ULONG	m_ulRef;
}ContextMenuExtStruct;

// IShellExtInit methods
STDMETHODIMP         CShellInitExt_QueryInterface(IShellExtInit * this, REFIID riid, LPVOID* ppvObject);
STDMETHODIMP_(ULONG) CShellInitExt_AddRef(IShellExtInit *this);
STDMETHODIMP_(ULONG) CShellInitExt_Release(IShellExtInit *this);
STDMETHODIMP         CShellInitExt_Initialize(IShellExtInit * this, LPCITEMIDLIST pidlFolder, LPDATAOBJECT lpdobj, HKEY hKeyProgID);

#endif // _CONTEXTMENU_H_


