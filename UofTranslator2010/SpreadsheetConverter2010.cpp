// SpreadsheetConverter2010.cpp : Implementation of CSpreadsheetConverter2010

#include "stdafx.h"
#include "SpreadsheetConverter2010.h"

using namespace System;
using namespace AtlUofTranslatorDll;
using namespace System::Runtime::InteropServices;

// CSpreadsheetConverter2010

STDMETHODIMP CSpreadsheetConverter2010::HrInitConverter(IConverterApplicationPreferences *pcap, 
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

STDMETHODIMP CSpreadsheetConverter2010::HrUninitConverter(IConverterUICallback *pcuic)
{

	return S_OK;
}




STDMETHODIMP CSpreadsheetConverter2010::HrImport(BSTR bstrSourcePath, 
									BSTR bstrDestPath, 
									IConverterApplicationPreferences *pcap, 
									IConverterPreferences **ppcp, 
									IConverterUICallback *pcuic)
{

	AtlSpreadsheetUofTranslatorOperation ^SpreadsheetUofTranslatorOperation = gcnew AtlSpreadsheetUofTranslatorOperation();
	SpreadsheetUofTranslatorOperation->SimpleAtlSpreadsheetUofTranslatorImport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));
	
	return S_OK;
}

STDMETHODIMP CSpreadsheetConverter2010::HrExport(BSTR bstrSourcePath, 
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
	AtlSpreadsheetUofTranslatorOperation ^SpreadsheetUofTranslatorOperation = gcnew AtlSpreadsheetUofTranslatorOperation();
	SpreadsheetUofTranslatorOperation->SimpleAtlSpreadsheetUofTranslatorExport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));
	//NativeOperation  ^NativeOp = gcnew NativeOperation();
	//NativeOp->SimpleExport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));


	return S_OK;
}

STDMETHODIMP CSpreadsheetConverter2010::HrGetFormat(BSTR bstrPath,
								BSTR *pbstrClass,
								IConverterApplicationPreferences *pcap, 
								IConverterPreferences **ppcp, 
								IConverterUICallback *pcuic)
{
	//BSTR p = ::SysAllocString(L"UOF Spreadsheet Converter");//BSTR是一个指针类型，SysAllocString 申请一个BSRT指针，并初始化为一个字符串
	//*pbstrClass = p;
	return S_OK;
}

STDMETHODIMP CSpreadsheetConverter2010::HrGetErrorString(long hrErr, 
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