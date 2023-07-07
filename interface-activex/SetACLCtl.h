/////////////////////////////////////////////////////////////////////////////
//
//
//	SetACL.h
//
//
//	Description:	Declaration of the CSetACLCtrl ActiveX control class
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
// Defines
//
//////////////////////////////////////////////////////////////////////


#define ICON_WIDTH	35
#define ICON_HEIGHT	35


//////////////////////////////////////////////////////////////////////
//
// Class CSetACLCtrl
//
//////////////////////////////////////////////////////////////////////


class CSetACLCtrl : public COleControl
{
	DECLARE_DYNCREATE (CSetACLCtrl)

public:
	CSetACLCtrl ();

	virtual	void	OnDraw (CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual	void	DoPropExchange (CPropExchange* pPX);
	virtual	void	OnResetState ();
	virtual	BOOL	OnSetExtent (LPSIZEL lpSizeL);
	virtual	BOOL	IsInvokeAllowed (DISPID dispid);
				void	FireMessageEvent (CString sMessage);

protected:
	~CSetACLCtrl ();

	DECLARE_OLECREATE_EX (CSetACLCtrl)
	DECLARE_OLETYPELIB (CSetACLCtrl)
	DECLARE_OLECTLTYPE (CSetACLCtrl)
	DECLARE_MESSAGE_MAP ()
	DECLARE_DISPATCH_MAP ()
	DECLARE_EVENT_MAP ()

public:
	enum
	{
		dispidSendMessageEvents			= 30L,
		dispidAddObjectFilter			= 29L,
		dispidSetBackupRestoreFile		= 28L,
		dispidSetObjectFlags				= 27L,
		dispidSetRecursion				= 26L,
		dispidAddAction					= 25L,
		dispidSetObject					= 24L,
		dispidSetIgnoreErrors			= 23L,
		dispidSetListOptions2			= 22L,
		dispidReset							= 21L,
		dispidGetLastListOutput			= 20L,
		dispidGetLastAPIError			= 19L,
		dispidGetLastAPIErrorMessage	= 18L,
		dispidGetResourceString			= 17L,
		dispidRun							= 16L,
		dispidSetLogFile					= 15L,
		dispidSetAction					= 14L,
		dispidSetListOptions				= 13L,
		dispidSetPrimaryGroup			= 12L,
		dispidSetOwner						= 11L,
		dispidAddDomain					= 10L,
		dispidAddTrustee					= 9L,
		dispidAddACE						= 8L,

		eventidMessageEvent				= 1L,
	};

private:
	CSetACL*		m_oSetACL;
	bool			m_fSendEvents;

protected:
	VARIANT_BOOL	SetIgnoreErrors (VARIANT_BOOL fIgnoreErrors);
	VARIANT_BOOL	SendMessageEvents (VARIANT_BOOL fSendEvents);
	LONG				SetObject (LPCTSTR sObjectPath, LONG nObjectType);
	LONG				AddAction (LONG nAction);
	LONG				SetRecursion (LONG nRecursionType);
	LONG				SetObjectFlags (LONG nDACLProtected, LONG nSACLProtected, VARIANT_BOOL fDACLResetChildObjects, VARIANT_BOOL fSACLResetChildObjects);
	LONG				SetBackupRestoreFile (LPCTSTR sBackupRestoreFile);
	void				AddObjectFilter (LPCTSTR sKeyword);
	LONG				AddACE (LPCTSTR sTrustee, VARIANT_BOOL fTrusteeIsSID, LPCTSTR sPermission, LONG nInheritance, VARIANT_BOOL fInhSpecified, LONG nAccessMode, LONG nACLType);
	LONG				AddTrustee (LPCTSTR sTrustee, LPCTSTR sNewTrustee, VARIANT_BOOL fTrusteeIsSID, VARIANT_BOOL fNewTrusteeIsSID, LONG nAction, VARIANT_BOOL fDACL, VARIANT_BOOL fSACL);
	LONG				AddDomain (LPCTSTR sDomain, LPCTSTR sNewDomain, LONG nAction, VARIANT_BOOL fDACL, VARIANT_BOOL fSACL);
	LONG				SetOwner (LPCTSTR sTrustee, VARIANT_BOOL fTrusteeIsSID);
	LONG				SetPrimaryGroup (LPCTSTR sTrustee, VARIANT_BOOL fTrusteeIsSID);
	LONG				SetListOptions (LONG nListFormat, LONG nListWhat, VARIANT_BOOL fListInherited, LONG nListNameSID);
	LONG				SetListOptions2 (LONG nListFormat, LONG nListWhat, VARIANT_BOOL fListInherited, LONG nListNameSID, VARIANT_BOOL fListCleanOutput);
	LONG				SetAction (LONG nAction);
	LONG				SetLogFile (LPCTSTR sLogFile);
	LONG				Run (void);
	BSTR				GetResourceString (LONG nID);
	BSTR				GetLastAPIErrorMessage (void);
	LONG				GetLastAPIError (void);
	BSTR				GetLastListOutput (void);
	void				Reset (void);
};


//////////////////////////////////////////////////////////////////////
//
// Function definitions
//
//////////////////////////////////////////////////////////////////////


void		PrintMsg (CString sMessage);		// Print a line to log file

