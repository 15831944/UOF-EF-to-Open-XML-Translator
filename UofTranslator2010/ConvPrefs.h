// ConvPrefs.h : Declaration of the CConvPrefs

#pragma once
#include "resource.h"       // main symbols

#include "UofTranslator2010_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CConvPrefs

class ATL_NO_VTABLE CConvPrefs :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConvPrefs, &CLSID_ConvPrefs>,
	public IDispatchImpl<IConvPrefs, &IID_IConvPrefs, &LIBID_UofTranslator2010Lib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IConverterPreferences
{
public:
	CConvPrefs()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CONVPREFS)


BEGIN_COM_MAP(CConvPrefs)
	COM_INTERFACE_ENTRY(IConvPrefs)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IConverterPreferences)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

public:
	STDMETHOD(HrGetMacroEnabled)(int * pfMacroEnabled)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(HrCheckFormat)(int * pFormat)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(HrGetLossySave)(int * pfLossySave)
	{
		return E_NOTIMPL;
	}


};

OBJECT_ENTRY_AUTO(__uuidof(ConvPrefs), CConvPrefs)
