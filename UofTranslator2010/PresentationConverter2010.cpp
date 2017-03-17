// PresentationConverter2010.cpp : Implementation of CPresentationConverter2010

#include "stdafx.h"
#include "PresentationConverter2010.h"


// CPresentationConverter2010

// CPresentationConverter

using namespace System;
using namespace AtlUofTranslatorDll;
using namespace System::Runtime::InteropServices;



// CWordConverter

STDMETHODIMP CPresentationConverter2010::HrInitConverter(IConverterApplicationPreferences *pcap, 
								IConverterPreferences **ppcp, 
								IConverterUICallback *pcuic)
{

	//int cInstances = 0;
	//WCHAR wzTalkingTo[1024];
	//const WCHAR *wzConverterClass = L"TestConv";


	//BSTR bstrApplication = NULL;

	//pcap->HrGetApplication(&bstrApplication);
	//CopyBstrToWz(bstrApplication, wzTalkingTo, cElements(wzTalkingTo));

	//*ppcp = new  CComObject <CConvPrefs>;
	//BSTR bstrApplication = NULL;
	//pcap->HrGetApplication(&bstrApplication);
	//*ppcp = new  CComObject <CWordConvPrefs>;

			
	return S_OK;
}

STDMETHODIMP CPresentationConverter2010::HrUninitConverter(IConverterUICallback *pcuic)
{

	return S_OK;
}




STDMETHODIMP CPresentationConverter2010::HrImport(BSTR bstrSourcePath, 
									BSTR bstrDestPath, 
									IConverterApplicationPreferences *pcap, 
									IConverterPreferences **ppcp, 
									IConverterUICallback *pcuic)
{

	AtlPresentationUofTranslatorOperation ^PresentationUofTranslatorOperation = gcnew AtlPresentationUofTranslatorOperation();
	PresentationUofTranslatorOperation->SimpleAtlPresentationUofTranslatorImport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));
	
	return S_OK;
}

STDMETHODIMP CPresentationConverter2010::HrExport(BSTR bstrSourcePath, 
									BSTR bstrDestPath, 
									BSTR bstrClass,
									IConverterApplicationPreferences *pcap, 
									IConverterPreferences **ppcp, 
									IConverterUICallback *pcuic)
{
	//FILE *stream;
	//if( (stream  = fopen( "C:\\Users\\v-xipia\\Desktop\\temp.txt", "a" )) == NULL )
	//	printf( "The file 'crt_fopen.c' was not opened\n" );
	//else
	//	printf( "The file 'crt_fopen.c' was opened\n" );
	//fprintf(stream,"%s\n","in HrExport.");
	//fclose(stream);
	AtlPresentationUofTranslatorOperation ^PresentationUofTranslatorOperation = gcnew AtlPresentationUofTranslatorOperation();
	PresentationUofTranslatorOperation->SimpleAtlPresentationUofTranslatorExport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));
	
	//return S_OK;
	//NativeOperation  ^NativeOp = gcnew NativeOperation();
	//NativeOp->SimpleExport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));


	return S_OK;
}

STDMETHODIMP CPresentationConverter2010::HrGetFormat(BSTR bstrPath,
								BSTR *pbstrClass,
								IConverterApplicationPreferences *pcap, 
								IConverterPreferences **ppcp, 
								IConverterUICallback *pcuic)
{
	
	return S_OK;
}

STDMETHODIMP CPresentationConverter2010::HrGetErrorString(long hrErr, 
									BSTR *pbstrErrorMsg,
									IConverterApplicationPreferences *pcap)
{
	//FILE *stream;
	//if( (stream  = fopen( "C:\\Users\\v-xipia\\Desktop\\temp.txt", "a" )) == NULL )
	//	printf( "The file 'crt_fopen.c' was not opened\n" );
	//else
	//	printf( "The file 'crt_fopen.c' was opened\n" );
	//fprintf(stream,"%s\n","in HrGetErrorString");
	//fclose(stream);




	return S_OK;
}

