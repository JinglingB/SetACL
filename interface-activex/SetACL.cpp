/////////////////////////////////////////////////////////////////////////////
//
//
//	SetACL.cpp
//
//
//	Description:	ActiveX interface for SetACL:
//
//						Implementation of CSetACLApp and DLL registration
//
//	Author:			Helge Klein
//
//	Created with:	MS Visual C++ 8.0
//
// Required:		Headers and libs from the platform SDK
//
//	Tabs set to:	3
//
//
/////////////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
//
// Includes
//
//////////////////////////////////////////////////////////////////////


#include "Stdafx.h"
#include "SetACL.h"


//////////////////////////////////////////////////////////////////////
//
// Defines
//
//////////////////////////////////////////////////////////////////////


#ifdef _DEBUG
	#define new DEBUG_NEW
#undef THIS_FILE
	static char THIS_FILE[] = __FILE__;
#endif


//////////////////////////////////////////////////////////////////////
//
// Global variables
//
//////////////////////////////////////////////////////////////////////


// The application object
CSetACLApp NEAR	theApp;


//////////////////////////////////////////////////////////////////////
//
// Constants
//
//////////////////////////////////////////////////////////////////////


const GUID CDECL BASED_CODE _tlid	=	{0x34052989, 0xca79, 0x44a1, {0x8e, 0x31, 0x31, 0xa6, 0xf1, 0x4b, 0x21, 0xf6}};

const WORD _wVerMajor	=	1;
const WORD _wVerMinor	=	0;


//////////////////////////////////////////////////////////////////////
//
// Functions
//
//////////////////////////////////////////////////////////////////////


//
// InitInstance - DLL initialization
//
BOOL CSetACLApp::InitInstance ()
{
	if (! COleControlModule::InitInstance ())
	{
		return false;
	}


	return true;
}


//
// ExitInstance - DLL end
//
int CSetACLApp::ExitInstance ()
{
	return COleControlModule::ExitInstance ();
}


//
// DllRegisterServer - adds COM data to the registry
//
STDAPI DllRegisterServer (void)
{
	AFX_MANAGE_STATE (_afxModuleAddrThis);

	if (! AfxOleRegisterTypeLib (AfxGetInstanceHandle (), _tlid))
	{
		return ResultFromScode (SELFREG_E_TYPELIB);
	}

	if (! COleObjectFactoryEx::UpdateRegistryAll (true))
	{
		return ResultFromScode (SELFREG_E_CLASS);
	}

	return NOERROR;
}


//
// DllUnregisterServer - removes COM data from the registry
//
STDAPI DllUnregisterServer (void)
{
	AFX_MANAGE_STATE (_afxModuleAddrThis);

	if (! AfxOleUnregisterTypeLib (_tlid, _wVerMajor, _wVerMinor))
	{
		return ResultFromScode (SELFREG_E_TYPELIB);
	}

	if (! COleObjectFactoryEx::UpdateRegistryAll (false))
	{
		return ResultFromScode(SELFREG_E_CLASS);
	}

	return NOERROR;
}
