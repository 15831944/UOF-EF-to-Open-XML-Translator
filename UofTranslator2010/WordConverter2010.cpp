// WordConverter2010.cpp : Implementation of CWordConverter2010

#include "stdafx.h"
#include "WordConverter2010.h"
#include "ConvPrefs.h"


using namespace System;
using namespace AtlUofTranslatorDll;
using namespace System::Runtime::InteropServices;
// CWordConverter2010

STDMETHODIMP CWordConverter2010::HrInitConverter(IConverterApplicationPreferences *pcap, 
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
	BSTR bstrApplication = NULL;
	pcap->HrGetApplication(&bstrApplication);
	*ppcp = new  CComObject <CConvPrefs>;

			
	return S_OK;
}

STDMETHODIMP CWordConverter2010::HrUninitConverter(IConverterUICallback *pcuic)
{

	return S_OK;
}




STDMETHODIMP CWordConverter2010::HrImport(BSTR bstrSourcePath, 
									BSTR bstrDestPath, 
									IConverterApplicationPreferences *pcap, 
									IConverterPreferences **ppcp, 
									IConverterUICallback *pcuic)
{

	AtlWordUofTranslatorOperation ^wordUofTranslatorOperation = gcnew AtlWordUofTranslatorOperation();
	wordUofTranslatorOperation->SimpleAtlWordUofTranslatorImport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));
	
	return S_OK;
}

STDMETHODIMP CWordConverter2010::HrExport(BSTR bstrSourcePath, 
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

	AtlWordUofTranslatorOperation ^wordUofTranslatorOperation = gcnew AtlWordUofTranslatorOperation();
	wordUofTranslatorOperation->SimpleAtlWordUofTranslatorExport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));

	//NativeOperation  ^NativeOp = gcnew NativeOperation();
	//NativeOp->SimpleExport(Marshal::PtrToStringBSTR(IntPtr(bstrSourcePath)),Marshal::PtrToStringBSTR(IntPtr(bstrDestPath)));


	return S_OK; //返回一个值，表示函数是否被正常执行了。
}

STDMETHODIMP CWordConverter2010::HrGetFormat(BSTR bstrPath,
								BSTR *pbstrClass,
								IConverterApplicationPreferences *pcap, 
								IConverterPreferences **ppcp, 
								IConverterUICallback *pcuic)
{
	
	return S_OK;
}

STDMETHODIMP CWordConverter2010::HrGetErrorString(long hrErr, 
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