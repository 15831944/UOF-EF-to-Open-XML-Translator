// SpreadsheetConverter2010.h : Declaration of the CSpreadsheetConverter2010

#pragma once
#include "resource.h"       // main symbols

#include "UofTranslator2010_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CSpreadsheetConverter2010

class ATL_NO_VTABLE CSpreadsheetConverter2010 :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSpreadsheetConverter2010, &CLSID_SpreadsheetConverter2010>,
	public IDispatchImpl<ISpreadsheetConverter2010, &IID_ISpreadsheetConverter2010, &LIBID_UofTranslator2010Lib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IConverter
{
public:
	CSpreadsheetConverter2010()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_SPREADSHEETCONVERTER2010)


BEGIN_COM_MAP(CSpreadsheetConverter2010)
	COM_INTERFACE_ENTRY(ISpreadsheetConverter2010)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IConverter)
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
	STDMETHOD(HrInitConverter)(IConverterApplicationPreferences * pcap, IConverterPreferences * * ppcp, IConverterUICallback * pcuic);

	STDMETHOD(HrUninitConverter)(IConverterUICallback * pcuic);

	STDMETHOD(HrImport)(BSTR bstrSourcePath, BSTR bstrDestPath, IConverterApplicationPreferences * pcap, IConverterPreferences * * ppcp, IConverterUICallback * pcuic);

	STDMETHOD(HrExport)(BSTR bstrSourcePath, BSTR bstrDestPath, BSTR bstrClass, IConverterApplicationPreferences * pcap, IConverterPreferences * * ppcp, IConverterUICallback * pcuic);

	STDMETHOD(HrGetFormat)(BSTR bstrPath, BSTR * pbstrClass, IConverterApplicationPreferences * pcap, IConverterPreferences * * ppcp, IConverterUICallback * pcuic);

	STDMETHOD(HrGetErrorString)(long hrErr, BSTR * pbstrErrorMsg, IConverterApplicationPreferences * pcap);

};

OBJECT_ENTRY_AUTO(__uuidof(SpreadsheetConverter2010), CSpreadsheetConverter2010)
