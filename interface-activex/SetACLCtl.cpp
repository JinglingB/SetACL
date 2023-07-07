/////////////////////////////////////////////////////////////////////////////
//
//
//	SetACL.cpp
//
//
//	Description:	ActiveX interface for SetACL:
//
//						Implementation of the CSetACLCtrl ActiveX control class
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
// Global variables
//
//////////////////////////////////////////////////////////////////////


// The control object - used by global functions to access class members
CSetACLCtrl*		oSetACLCtrl	=	NULL;


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
// Implements
//
//////////////////////////////////////////////////////////////////////


IMPLEMENT_DYNCREATE (CSetACLCtrl, COleControl)


//////////////////////////////////////////////////////////////////////
//
// Message map
//
//////////////////////////////////////////////////////////////////////


BEGIN_MESSAGE_MAP (CSetACLCtrl, COleControl)
//	ON_OLEVERB (AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP ()


//////////////////////////////////////////////////////////////////////
//
// Dispatch table
//
//////////////////////////////////////////////////////////////////////


BEGIN_DISPATCH_MAP (CSetACLCtrl, COleControl)
	DISP_FUNCTION_ID(CSetACLCtrl, "SendMessageEvents", dispidSendMessageEvents, SendMessageEvents, VT_BOOL, VTS_BOOL)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetIgnoreErrors", dispidSetIgnoreErrors, SetIgnoreErrors, VT_BOOL, VTS_BOOL)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetObject", dispidSetObject, SetObject, VT_I4, VTS_BSTR VTS_I4)
	DISP_FUNCTION_ID(CSetACLCtrl, "AddAction", dispidAddAction, AddAction, VT_I4, VTS_I4)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetRecursion", dispidSetRecursion, SetRecursion, VT_I4, VTS_I4)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetObjectFlags", dispidSetObjectFlags, SetObjectFlags, VT_I4, VTS_I4 VTS_I4 VTS_BOOL VTS_BOOL)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetBackupRestoreFile", dispidSetBackupRestoreFile, SetBackupRestoreFile, VT_I4, VTS_BSTR)
	DISP_FUNCTION_ID(CSetACLCtrl, "AddObjectFilter", dispidAddObjectFilter, AddObjectFilter, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(CSetACLCtrl, "AddACE", dispidAddACE, AddACE, VT_I4, VTS_BSTR VTS_BOOL VTS_BSTR VTS_I4 VTS_BOOL VTS_I4 VTS_I4)
	DISP_FUNCTION_ID(CSetACLCtrl, "AddTrustee", dispidAddTrustee, AddTrustee, VT_I4, VTS_BSTR VTS_BSTR VTS_BOOL VTS_BOOL VTS_I4 VTS_BOOL VTS_BOOL)
	DISP_FUNCTION_ID(CSetACLCtrl, "AddDomain", dispidAddDomain, AddDomain, VT_I4, VTS_BSTR VTS_BSTR VTS_I4 VTS_BOOL VTS_BOOL)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetOwner", dispidSetOwner, SetOwner, VT_I4, VTS_BSTR VTS_BOOL)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetPrimaryGroup", dispidSetPrimaryGroup, SetPrimaryGroup, VT_I4, VTS_BSTR VTS_BOOL)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetListOptions", dispidSetListOptions, SetListOptions, VT_I4, VTS_I4 VTS_I4 VTS_BOOL VTS_I4)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetListOptions2", dispidSetListOptions2, SetListOptions2, VT_I4, VTS_I4 VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetAction", dispidSetAction, SetAction, VT_I4, VTS_I4)
	DISP_FUNCTION_ID(CSetACLCtrl, "SetLogFile", dispidSetLogFile, SetLogFile, VT_I4, VTS_BSTR)
	DISP_FUNCTION_ID(CSetACLCtrl, "Run", dispidRun, Run, VT_I4, VTS_NONE)
	DISP_FUNCTION_ID(CSetACLCtrl, "GetResourceString", dispidGetResourceString, GetResourceString, VT_BSTR, VTS_I4)
	DISP_FUNCTION_ID(CSetACLCtrl, "GetLastAPIErrorMessage", dispidGetLastAPIErrorMessage, GetLastAPIErrorMessage, VT_BSTR, VTS_NONE)
	DISP_FUNCTION_ID(CSetACLCtrl, "GetLastAPIError", dispidGetLastAPIError, GetLastAPIError, VT_I4, VTS_NONE)
	DISP_FUNCTION_ID(CSetACLCtrl, "GetLastListOutput", dispidGetLastListOutput, GetLastListOutput, VT_BSTR, VTS_NONE)
	DISP_FUNCTION_ID(CSetACLCtrl, "Reset", dispidReset, Reset, VTS_NONE, VTS_NONE)
END_DISPATCH_MAP ()


//////////////////////////////////////////////////////////////////////
//
// Event map
//
//////////////////////////////////////////////////////////////////////


BEGIN_EVENT_MAP (CSetACLCtrl, COleControl)
	EVENT_CUSTOM_ID("MessageEvent", eventidMessageEvent, FireMessageEvent, VTS_BSTR)
END_EVENT_MAP ()


//////////////////////////////////////////////////////////////////////
//
// Class creation and GUID initialization
//
//////////////////////////////////////////////////////////////////////


IMPLEMENT_OLECREATE_EX (CSetACLCtrl, "SETACL.SetACLCtrl.1", 0x17738736, 0x8831, 0x4197, 0xac, 0xa5, 0xa7, 0x40, 0xa0, 0x50, 0xb4, 0xcf)


//////////////////////////////////////////////////////////////////////
//
// Type library ID and version
//
//////////////////////////////////////////////////////////////////////


IMPLEMENT_OLETYPELIB (CSetACLCtrl, _tlid, _wVerMajor, _wVerMinor)


//////////////////////////////////////////////////////////////////////
//
// Interface IDs
//
//////////////////////////////////////////////////////////////////////


const IID BASED_CODE IID_DSetACL			= {0x85869435, 0xee3, 0x440b, {0xbf, 0x69, 0x6c, 0x52, 0xc6, 0x63, 0x80, 0x73}};
const IID BASED_CODE IID_DSetACLEvents	= {0x89c18363, 0x8f5, 0x4ff7, {0x9c, 0x82, 0xfe, 0x75, 0x79, 0x2f, 0x4c, 0x3f}};


//////////////////////////////////////////////////////////////////////
//
// Control type information
//
//////////////////////////////////////////////////////////////////////


static const DWORD BASED_CODE _dwSetACLOleMisc	=	OLEMISC_INVISIBLEATRUNTIME | 
																	OLEMISC_SETCLIENTSITEFIRST |  
																	OLEMISC_INSIDEOUT | 
																	OLEMISC_CANTLINKINSIDE |
																	OLEMISC_ONLYICONIC |
																	OLEMISC_NOUIACTIVATE;

IMPLEMENT_OLECTLTYPE (CSetACLCtrl, IDS_SETACL, _dwSetACLOleMisc)


//////////////////////////////////////////////////////////////////////
//
// Functions
//
//////////////////////////////////////////////////////////////////////


//
// UpdateRegistry - Adds entries to the registries or removes them (to be or not ...)
//
BOOL CSetACLCtrl::CSetACLCtrlFactory::UpdateRegistry (BOOL bRegister)
{
	if (bRegister)
	{
		return AfxOleRegisterControlClass (AfxGetInstanceHandle (), m_clsid, m_lpszProgID, IDS_SETACL, IDB_SETACL, afxRegApartmentThreading | afxRegFreeThreading, 
			_dwSetACLOleMisc, _tlid, _wVerMajor, _wVerMinor);
	}
	else
	{
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
	}
}


//
// Constructor
//
CSetACLCtrl::CSetACLCtrl ()
{
	InitializeIIDs (&IID_DSetACL, &IID_DSetACLEvents);

	SetInitialSize (ICON_WIDTH, ICON_HEIGHT);

	// Initialize member variables
	m_fSendEvents	=	false;

	// Create our SetACL class object
	m_oSetACL		=	new CSetACL (PrintMsg);

	// Store a pointer to this class in a global variable which can be used by global functions to access our members (mainly PrintMsg ())
	::oSetACLCtrl	=	this;
}


//
// Destructor
//
CSetACLCtrl::~CSetACLCtrl ()
{
	if (m_oSetACL)
	{
		delete m_oSetACL;
      m_oSetACL = NULL;
	}
}


//
// OnDraw - Paint the control
//
void CSetACLCtrl::OnDraw (CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	UNUSED_ALWAYS(rcInvalid);		// suppress warning (unreferenced formal parameter)

	CBitmap			oBitmap;
	BITMAP			bmpIcon;
	CPictureHolder oPictureHolder;
	CRect				oSourceBounds;

	// Load the bitmap
	oBitmap.LoadBitmap (IDB_SETACL);
	oBitmap.GetObject (sizeof (BITMAP), &bmpIcon);
	oSourceBounds.right	= bmpIcon.bmWidth;
	oSourceBounds.bottom	= bmpIcon.bmHeight;

	::DrawEdge (pdc->GetSafeHdc (), CRect (rcBounds), EDGE_RAISED, BF_RECT | BF_ADJUST);

	oPictureHolder.CreateFromBitmap ((HBITMAP) oBitmap.m_hObject, NULL, false);

	// Render the control
	oPictureHolder.Render (pdc, rcBounds, oSourceBounds);
}


//
// DoPropExchange - support for persistent properties
//
void CSetACLCtrl::DoPropExchange (CPropExchange* pPX)
{
	ExchangeVersion (pPX, MAKELONG (_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange (pPX);
}


//
// OnResetState - Resets the control's state
//
void CSetACLCtrl::OnResetState ()
{
	COleControl::OnResetState ();
}


//
// OnSetExtent - Handle resize events
//
BOOL CSetACLCtrl::OnSetExtent(LPSIZEL lpSizeL)
{
	UNUSED_ALWAYS(lpSizeL);		// suppress warning (unreferenced formal parameter)

   return false;
}


//
// IsInvokeAllowed - Return true to allow usage from script languages
//
//	See Q146120 in the MS KB
//
BOOL CSetACLCtrl::IsInvokeAllowed(DISPID dispid)
{
	UNUSED_ALWAYS(dispid);		// suppress warning (unreferenced formal parameter)

   return true;
}


//
// SetIgnoreErrors
//
VARIANT_BOOL CSetACLCtrl::SetIgnoreErrors (VARIANT_BOOL fIgnoreErrors) 
{
	AFX_MANAGE_STATE (AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return false;
	}

	return m_oSetACL->SetIgnoreErrors (fIgnoreErrors != VARIANT_FALSE);
}


//
// SetObject
//
LONG CSetACLCtrl::SetObject (LPCTSTR sObjectPath, LONG nObjectType) 
{
	AFX_MANAGE_STATE (AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetObject (sObjectPath, (SE_OBJECT_TYPE) nObjectType);
}


//
// AddAction
//
LONG CSetACLCtrl::AddAction (LONG nAction)
{
	AFX_MANAGE_STATE (AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->AddAction (nAction);
}


//
// SetRecursion
//
LONG CSetACLCtrl::SetRecursion (LONG nRecursionType)
{
	AFX_MANAGE_STATE (AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetRecursion (nRecursionType);
}


//
// SetObjectFlags
//
LONG CSetACLCtrl::SetObjectFlags (LONG nDACLProtected, LONG nSACLProtected, VARIANT_BOOL fDACLResetChildObjects, VARIANT_BOOL fSACLResetChildObjects)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetObjectFlags (nDACLProtected, nSACLProtected, fDACLResetChildObjects != VARIANT_FALSE, fSACLResetChildObjects != VARIANT_FALSE);
}


//
// SetBackupRestoreFile
//
LONG CSetACLCtrl::SetBackupRestoreFile (LPCTSTR sBackupRestoreFile)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetBackupRestoreFile (sBackupRestoreFile);
}


//
// AddObjectFilter
//
void CSetACLCtrl::AddObjectFilter (LPCTSTR sKeyword)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return;
	}

	m_oSetACL->AddObjectFilter (sKeyword);
}


//
// AddACE
//
LONG CSetACLCtrl::AddACE (LPCTSTR sTrustee, VARIANT_BOOL fTrusteeIsSID, LPCTSTR sPermission, LONG nInheritance, VARIANT_BOOL fInhSpecified, LONG nAccessMode, LONG nACLType)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->AddACE (sTrustee, fTrusteeIsSID != VARIANT_FALSE, sPermission, nInheritance, fInhSpecified != VARIANT_FALSE, nAccessMode, nACLType);
}


//
// AddTrustee
//
LONG CSetACLCtrl::AddTrustee (LPCTSTR sTrustee, LPCTSTR sNewTrustee, VARIANT_BOOL fTrusteeIsSID, VARIANT_BOOL fNewTrusteeIsSID, LONG nAction, VARIANT_BOOL fDACL, VARIANT_BOOL fSACL)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->AddTrustee (sTrustee, sNewTrustee, fTrusteeIsSID != VARIANT_FALSE, fNewTrusteeIsSID != VARIANT_FALSE, nAction,
		fDACL != VARIANT_FALSE, fSACL != VARIANT_FALSE);
}


//
// AddDomain
//
LONG CSetACLCtrl::AddDomain (LPCTSTR sDomain, LPCTSTR sNewDomain, LONG nAction, VARIANT_BOOL fDACL, VARIANT_BOOL fSACL)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->AddDomain (sDomain, sNewDomain, nAction, fDACL != VARIANT_FALSE, fSACL != VARIANT_FALSE);
}


//
// SetOwner
//
LONG CSetACLCtrl::SetOwner (LPCTSTR sTrustee, VARIANT_BOOL fTrusteeIsSID)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetOwner (sTrustee, fTrusteeIsSID != VARIANT_FALSE);
}


//
// SetPrimaryGroup
//
LONG CSetACLCtrl::SetPrimaryGroup (LPCTSTR sTrustee, VARIANT_BOOL fTrusteeIsSID)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetPrimaryGroup (sTrustee, fTrusteeIsSID != VARIANT_FALSE);
}


//
// SetListOptions
//
LONG CSetACLCtrl::SetListOptions (LONG nListFormat, LONG nListWhat, VARIANT_BOOL fListInherited, LONG nListNameSID)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetListOptions (nListFormat, nListWhat, fListInherited != VARIANT_FALSE, nListNameSID, false);
}


//
// SetListOptions2: An extended version
//
LONG CSetACLCtrl::SetListOptions2 (LONG nListFormat, LONG nListWhat, VARIANT_BOOL fListInherited, LONG nListNameSID, VARIANT_BOOL fListCleanOutput)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetListOptions (nListFormat, nListWhat, fListInherited != VARIANT_FALSE, nListNameSID, fListCleanOutput != VARIANT_FALSE);
}


//
// PrintMsg: Called by various CSetACL functions. Sends status messages to caller
//
void PrintMsg (CString sMessage)
{
	if (oSetACLCtrl)
	{
		oSetACLCtrl->FireMessageEvent (sMessage);
	}
}


//
// SetAction: Set the action to be performed. All former values are erased.
//
LONG CSetACLCtrl::SetAction (LONG nAction)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetAction (nAction);
}


//
// SetLogFile: Set the log (backup) file to use.
//
LONG CSetACLCtrl::SetLogFile (LPCTSTR sLogFile)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->SetLogFile (sLogFile);
}


//
// Run: Start the action
//
LONG CSetACLCtrl::Run (void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->Run ();
}


//
// GetResourceString: Retrieve a resource string (used for error messages) identified by ID
//
BSTR CSetACLCtrl::GetResourceString (LONG nID)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString sResult;

	if (! sResult.LoadString (nID))
	{
		sResult.Empty ();
	}

	return sResult.AllocSysString ();
}


//
// GetLastAPIErrorMessage: Retrieve the last API error as a localized string
//
BSTR CSetACLCtrl::GetLastAPIErrorMessage (void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strResult;

	if (m_oSetACL)
	{
		strResult	=	m_oSetACL->GetLastErrorMessage ();
	}

	return strResult.AllocSysString ();
}


//
// GetLastAPIError: Retrieve the last API error as an error number
//
LONG CSetACLCtrl::GetLastAPIError (void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (! m_oSetACL)
	{
		return RTN_ERR_INTERNAL;
	}

	return m_oSetACL->m_nAPIError;
}


//
// Send a message via a COM event
//
void CSetACLCtrl::FireMessageEvent (CString sMessage)
{
	if (m_fSendEvents != true)
		return;

	// Create a BSTR from our CString
	BSTR	bstrMessage	=	sMessage.AllocSysString ();

	// Fire the event, passing the BSTR
	FireEvent (eventidMessageEvent, EVENT_PARAM (VTS_BSTR), bstrMessage);

	// Free memory
	// SysFreeString (bstrMessage);
}


//
// Retrieve the last output generated by the list function
//
BSTR CSetACLCtrl::GetLastListOutput (void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strResult;

	if (m_oSetACL)
	{
		strResult	=	m_oSetACL->GetLastListOutput();
	}

	return strResult.AllocSysString ();
}


//
// Reset the object to its initial state
//
void CSetACLCtrl::Reset (void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (m_oSetACL)
	{
		m_oSetACL->Reset (true);
	}
}


//
// SetIgnoreErrors
//
VARIANT_BOOL CSetACLCtrl::SendMessageEvents (VARIANT_BOOL fSendEvents)
{
	AFX_MANAGE_STATE (AfxGetStaticModuleState());

	if (fSendEvents == VARIANT_TRUE)
		m_fSendEvents = true;
	else
		m_fSendEvents = false;

	return true;
}


