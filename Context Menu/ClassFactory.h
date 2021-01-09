// classfactory.h

#ifndef _CLASSFACTORY_H_
#define _CLASSFACTORY_H_

// IClassFactory methods
STDMETHODIMP         CClassFactory_QueryInterface(IClassFactory *this,REFIID riid, LPVOID *ppvOut);
STDMETHODIMP_(ULONG) CClassFactory_AddRef(IClassFactory *this);
STDMETHODIMP_(ULONG) CClassFactory_Release(IClassFactory *this);
STDMETHODIMP         CClassFactory_CreateInstance(IClassFactory *, LPUNKNOWN, REFIID, LPVOID *);
STDMETHODIMP         CClassFactory_LockServer(IClassFactory *this, BOOL);

// IClassFactory constructor
IClassFactory * CClassFactory_Create(void);

// This struct acts somewhat like a pseudo class in that you have
// variables accociated with an instance of this interface.
typedef struct _ClassFactoryStruct
{
	// first part of the struct for the vtable must be fc
	IClassFactory fc;

	// second part of tye struct for the variables
	HINSTANCE	m_hInstance;			// Instance handle for this DLL
	ULONG		m_ulRef;                // Object reference count
}ClassFactoryStruct;

#endif //CLASSFACTORY

