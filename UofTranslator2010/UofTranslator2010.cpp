// UofTranslator2010.cpp : Implementation of WinMain


#include "stdafx.h"
#include "resource.h"
#include "UofTranslator2010_i.h"


class CUofTranslator2010Module : public CAtlExeModuleT< CUofTranslator2010Module >
{
public :
	DECLARE_LIBID(LIBID_UofTranslator2010Lib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_UOFTRANSLATOR2010, "{6A924991-0C4F-43DA-9BD2-944F74B524E3}")
};

CUofTranslator2010Module _AtlModule;



//
extern "C" int WINAPI _tWinMain(HINSTANCE /*hInstance*/, HINSTANCE /*hPrevInstance*/, 
                                LPTSTR /*lpCmdLine*/, int nShowCmd)
{
    return _AtlModule.WinMain(nShowCmd);
}

