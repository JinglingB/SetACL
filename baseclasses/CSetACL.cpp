/////////////////////////////////////////////////////////////////////////////
//
//
//	CSetACL.cpp
//
//
//	Description:	Main (worker) classes for SetACL
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
#include "CSetACL.h"
#include <comdef.h>

#pragma comment(lib, "netapi32.lib")
#pragma comment(lib, "mpr.lib")

//////////////////////////////////////////////////////////////////////
//
// Class CSetACL
//
//////////////////////////////////////////////////////////////////////


//
// Constructor: initialize all member variables
//
CSetACL::CSetACL (void (* funcNotify) (CString))
{
	// Notification function
	m_funcNotify					=	funcNotify;

	// Set initial values
	Init ();

	// Resolve well-known SIDs required somewhere else (for performance reasons we do that only once)
	m_WellKnownSIDCreatorOwner.m_sTrustee			= L"S-1-3-0";
	m_WellKnownSIDCreatorOwner.m_fTrusteeIsSID	= true;
	m_WellKnownSIDCreatorOwner.LookupSID ();
}


//
// Destructor: clean up
//
CSetACL::~CSetACL ()
{
	Reset (false);
}


//
// Reset the object to its initial state
//
void CSetACL::Reset (bool bReInit)
{
	// Free memory
   delete m_pOwner;
   m_pOwner = NULL;

   delete m_pPrimaryGroup;
   m_pPrimaryGroup = NULL;

	//
	// Delete all entries from m_lstACEs
	//
	POSITION	pos	=	m_lstACEs.GetHeadPosition ();

	while (pos)
	{
		delete m_lstACEs.GetNext (pos);
	}

	// Remove all list entries. Objects pointed to by the list members are NOT destroyed!
	m_lstACEs.RemoveAll ();

	//
	// Delete all entries from m_lstTrustees
	//
	pos	=	m_lstTrustees.GetHeadPosition ();

	while (pos)
	{
		CTrustee*	oTrustee	=	m_lstTrustees.GetNext (pos);

		if (oTrustee->m_oNewTrustee) delete oTrustee->m_oNewTrustee;
		delete oTrustee;
	}

	// Remove all list entries. Objects pointed to by the list members are NOT destroyed!
	m_lstTrustees.RemoveAll ();

	//
	// Delete all entries from m_lstDomains
	//
	pos	=	m_lstDomains.GetHeadPosition ();

	while (pos)
	{
		CDomain*	oDomain	=	m_lstDomains.GetNext (pos);

		if (oDomain->m_oNewDomain) delete oDomain->m_oNewDomain;
		delete oDomain;
	}

	// Remove all list entries. Objects pointed to by the list members are NOT destroyed!
	m_lstDomains.RemoveAll ();

	// Remove all entries from the filter list
	m_lstObjectFilter.RemoveAll ();

	// Close the output file, if necessary
	if (m_fhBackupRestoreFile)
	{
		fclose (m_fhBackupRestoreFile);
		m_fhBackupRestoreFile	=	NULL;
	}
   m_sBackupRestoreFile.Empty();

	// Close the log file if it is open
	if (m_fhLog)
	{
		fclose (m_fhLog);

		m_fhLog				=	NULL;
	}

   if (bReInit)
		Init ();
}


//
// Initialize the object
//
void CSetACL::Init ()
{
	m_nObjectType					=	SE_UNKNOWN_OBJECT_TYPE;
	m_nAction						=	0;
	m_nDACLProtected				=	INHPARNOCHANGE;
	m_fDACLResetChildObjects	=	false;
	m_nSACLProtected				=	INHPARNOCHANGE;
	m_fSACLResetChildObjects	=	false;
	m_nRecursionType				=	RECURSE_NO;
	m_nAPIError						=	0;
	m_paclExistingDACL			=	NULL;
	m_paclExistingSACL			=	NULL;
	m_psidExistingOwner			=	NULL;
	m_psidExistingGroup			=	NULL;
	m_pOwner							=	new CTrustee;
	m_pPrimaryGroup				=	new CTrustee;
	m_nDACLEntries					=	0;
	m_nSACLEntries					=	0;
	m_nListFormat					=	LIST_CSV;
	m_nListWhat						=	ACL_DACL;
	m_nListNameSID					=	LIST_NAME;
	m_fListInherited				=	false;
	m_fProcessSubObjectsOnly	=	false;
	m_fhBackupRestoreFile		=	NULL;
	m_fIgnoreErrors				=	false;
	m_fhLog							=	NULL;
	m_fTrusteesProcessDACL		=	false;
	m_fTrusteesProcessSACL		=	false;
	m_fDomainsProcessDACL		=	false;
	m_fDomainsProcessSACL		=	false;
	m_fRawMode						=	false;
   m_sBackupRestoreFile.Empty();
}

//
// SetObject: Sets the path and type of the object on which all actions are performed
//
DWORD CSetACL::SetObject (CString sObjectPath, SE_OBJECT_TYPE nObjectType)
{
	DWORD		nLocalPart;
	HANDLE	hNetEnum;

	if (sObjectPath.GetLength () >= 1 && 
		(nObjectType == SE_FILE_OBJECT || nObjectType == SE_SERVICE || nObjectType == SE_PRINTER || nObjectType == SE_REGISTRY_KEY || nObjectType == SE_LMSHARE || nObjectType == SE_WMIGUID_OBJECT))
	{
		m_sObjectPath	=	sObjectPath;
		m_nObjectType	=	nObjectType;

		//
		// Try to determine the name of the computer the path points to
		//
		m_sTargetSystemName	= TEXT ("");

		// If the path starts with a drive letter, try to find the server it is connected to
		if (sObjectPath.GetLength () >= 2 && nObjectType == SE_FILE_OBJECT && sObjectPath[1] == TCHAR (':'))
		{
			// Enumerate network connections
			if (WNetOpenEnum (RESOURCE_CONNECTED, RESOURCETYPE_DISK, 0, NULL, &hNetEnum) == NO_ERROR)
			{
				// Initialize a buffer for WNetEnumResource
				DWORD	nListEntries	= 1;
				DWORD	nBufferSize		= 16384;
				BYTE*	pBuffer			= new BYTE[16384];
				SecureZeroMemory (pBuffer, nBufferSize);

				while (WNetEnumResource (hNetEnum, &nListEntries, pBuffer, &nBufferSize) == NO_ERROR)
				{
					NETRESOURCE*	nrConnection	=	(NETRESOURCE*) pBuffer;
					CString			sLocalName		=	nrConnection->lpLocalName;

					if (sObjectPath.Left (2).CompareNoCase (sLocalName) == 0)
					{
						sObjectPath						=	nrConnection->lpRemoteName;
					}

					nListEntries	=	1;
					nBufferSize		=	16384;
					SecureZeroMemory (pBuffer, nBufferSize);
				}

				delete [] pBuffer;
            pBuffer = NULL;
				WNetCloseEnum (hNetEnum);
			}
		}

   	// Find the computer name from UNC paths
		if (sObjectPath.Left (2) == L"\\\\")
		{
         // Treat the path as a DFS link and try to find the target
         DFS_INFO_3* pDFSInfo = NULL;
         if (NetDfsGetClientInfo(sObjectPath.GetBuffer(), NULL, NULL, 3, (BYTE**) &pDFSInfo) == NERR_Success)
         {
            if (pDFSInfo && pDFSInfo->NumberOfStorages > 0 && pDFSInfo->Storage != NULL)
            {
               DFS_STORAGE_INFO* pStorages = pDFSInfo->Storage;
               for (DWORD i = 0; i < pDFSInfo->NumberOfStorages; i++)
               {
                  DFS_STORAGE_INFO* pStorage = pStorages + i;
                  if ((pStorage->State & DFS_STORAGE_STATE_ACTIVE) == DFS_STORAGE_STATE_ACTIVE)
                  {
                     m_sTargetSystemName = pStorage->ServerName;
                  }
               }
            }

            // Clean up
            NetApiBufferFree(pDFSInfo);
            pDFSInfo = NULL;
         }

         if (m_sTargetSystemName.IsEmpty())
         {
            // Extract the computer name from the UNC path
			   nLocalPart = sObjectPath.Find (L"\\", 2);

			   if (nLocalPart != -1)
			   {
				   // Extract the computer name
				   m_sTargetSystemName	=	sObjectPath.Mid (2, nLocalPart - 2);
			   }
         }
		}

		return RTN_OK;
	}
	else
	{
		m_sObjectPath	=	TEXT ("");
		m_nObjectType	=	SE_UNKNOWN_OBJECT_TYPE;

		return RTN_ERR_PARAMS;
	}
}


//
// SetBackupRestoreFile: Specify a (unicode) file to be used for backup/restore operations
//
DWORD CSetACL::SetBackupRestoreFile (CString sBackupRestoreFile)
{
	m_sBackupRestoreFile	=	sBackupRestoreFile;

	return RTN_OK;
}


//
// LogMessage: Log a message string
//
DWORD CSetACL::LogMessage (LogSeverities eSeverity, CString sMessage)
{
	int nBytesWritten	=	0;

	switch (eSeverity)
	{
	case Information:
		sMessage = L"INFORMATION: " + sMessage;
		break;
	case Warning:
		sMessage = L"WARNING: " + sMessage;
		break;
	case Error:
		sMessage = L"ERROR: " + sMessage;
		break;
	}

	// Pass the message on to our caller
	if (m_funcNotify)
	{
		(*m_funcNotify) (sMessage);
	}

	// Write the message to the log file, if there is one
	if (m_fhLog)
	{
		// Output file was opened in binary mode (due to unicode). Correct newlines.
		sMessage.Replace (L"\n", L"\r\n");

		// Print to log file
		nBytesWritten	=	_ftprintf (m_fhLog, L"%s\r\n", sMessage.GetString ());

		// Flush the buffer to disk directly
		fflush (m_fhLog);
	}

	if (nBytesWritten < 0)
	{
		return RTN_ERR_WRITE_LOGFILE;
	}
	else
	{
		return RTN_OK;
	}
}


//
// SetLogFile: Specify a (unicode) file to be used for logging
//
DWORD CSetACL::SetLogFile (CString sLogFile)
{
	m_sLogFile				=	sLogFile;

	// Close the log file if it is already open
	if (m_fhLog)
	{
		fclose (m_fhLog);

		m_fhLog				=	NULL;
	}

	// Check if the log file already exists
	errno_t	nErr			=	0;
	FILE*		fhTest		=	NULL;

	nErr	=	_tfopen_s (&fhTest, sLogFile, L"r");
	bool	fExists			=	(fhTest != NULL);

	if (fhTest)	fclose (fhTest);

	// Open the log file for appending. If it does not exist, it is created
	nErr	=	_tfopen_s (&m_fhLog, sLogFile, L"ab");

	if (m_fhLog == NULL || nErr != 0)
	{
		LogMessage (Error, L"Opening log file: <" + sLogFile + L"> failed!");

		return RTN_ERR_OPEN_LOGFILE;
	}

	// Write the BOM if the file did not exist already
	if (! fExists)
	{
		WCHAR cBOM	=	0xFEFF;
		fwrite (&cBOM, sizeof (WCHAR), 1, m_fhLog);
	}

	return RTN_OK;
}


//
// SetAction: Set the action to be performed. All former values are erased.
//
DWORD CSetACL::SetAction (DWORD nAction)
{
	if (CheckAction (nAction))
	{
		m_nAction	=	nAction;

		return RTN_OK;
	}
	else
	{
		m_nAction	=	0;

		return RTN_ERR_PARAMS;
	}
}


//
// AddAction: Add an action to be performed. All former values are preserved.
//
DWORD CSetACL::AddAction (DWORD nAction)
{
	if (CheckAction (nAction))
	{
		m_nAction	|=	nAction;

		return RTN_OK;
	}
	else
	{
		return RTN_ERR_PARAMS;
	}
}


//
// AddObjectFilter: Add a keyword to be filtered out - objects containing this keyword are not processed
//
void CSetACL::AddObjectFilter (CString sKeyword)
{
	m_lstObjectFilter.AddTail (sKeyword);
}


//
// AddACE: Add an ACE to be processed.
//
DWORD CSetACL::AddACE (CString sTrustee, bool fTrusteeIsSID, CString sPermission, DWORD nInheritance, bool fInhSpecified, DWORD nAccessMode, DWORD nACLType)
{
	if (sTrustee.IsEmpty ())
	{
		LogMessage (Error, L"AddACE: No trustee specified.");

		return RTN_ERR_PARAMS;
	}

	if (fInhSpecified && (! CheckInheritance (nInheritance)))
	{
		LogMessage (Error, L"AddACE: Invalid inheritance specified.");

		return RTN_ERR_PARAMS;
	}

	if (! CheckACEAccessMode (nAccessMode, nACLType))
	{
		LogMessage (Error, L"AddACE: Invalid access mode for this ACL type specified (e.g. you cannot add audit ACEs to the DACL, only to the SACL).");

		return RTN_ERR_PARAMS;
	}

	// Prevent audit ACEs to be added to shares
	if (m_nObjectType == SE_LMSHARE && (nAccessMode == SET_AUDIT_SUCCESS || nAccessMode == SET_AUDIT_FAILURE))
	{
		LogMessage (Error, L"AddACE: Audit ACEs cannot be set on shares.");

		return RTN_ERR_PARAMS;
	}

	CTrustee*	pTrustee	=	new CTrustee  (sTrustee, fTrusteeIsSID, ACTN_ADDACE, false, false);
	CACE*			pACE		=	new CACE (pTrustee, sPermission, nInheritance, fInhSpecified, (ACCESS_MODE) nAccessMode, nACLType);

	// Modify the count of DACL/SACL entries in the list
	if (nACLType == ACL_DACL)
	{
		m_nDACLEntries++;
	}
	else if (nACLType == ACL_SACL)
	{
		m_nSACLEntries++;
	}

	m_lstACEs.AddTail (pACE);

	return RTN_OK;
}


//
// AddTrustee: Add a trustee to be processed.
//
DWORD CSetACL::AddTrustee (CString sTrustee, CString sNewTrustee, bool fTrusteeIsSID, bool fNewTrusteeIsSID, DWORD nAction, bool fDACL, bool fSACL)
{
	if (! sTrustee.IsEmpty ())
	{
		CTrustee*	pTrustee		=	NULL;
		CTrustee*	pNewTrustee	=	NULL;

		try
		{
			pTrustee		=	new CTrustee  (sTrustee, fTrusteeIsSID, nAction, fDACL, fSACL);
			pNewTrustee	=	new CTrustee  (sNewTrustee, fNewTrusteeIsSID, nAction, false, false);

			pTrustee->m_oNewTrustee	=	pNewTrustee;

			m_lstTrustees.AddTail (pTrustee);

			// Remember whether DACL and/or SACL have to be processed
			if (fDACL)
			{
				m_fTrusteesProcessDACL	=	true;
			}
			if (fSACL)
			{
				m_fTrusteesProcessSACL	=	true;
			}

			return RTN_OK;

		}
		catch (...)
		{
			delete pTrustee;
			pTrustee = NULL;
			delete pNewTrustee;
			pNewTrustee = NULL;

			return RTN_ERR_OUT_OF_MEMORY;
		}
	}
	else
	{
		LogMessage (Error, L"AddTrustee: No trustee specified.");

		return RTN_ERR_PARAMS;
	}
}


//
// AddDomain: Add a domain to be processed.
//
DWORD CSetACL::AddDomain (CString sDomain, CString sNewDomain, DWORD nAction, bool fDACL, bool fSACL)
{
	DWORD	nError					=	RTN_OK;

	CDomain* pDomain		= NULL;
	CDomain*	pNewDomain	= NULL;

	try
	{
		if (! sDomain.IsEmpty ())
		{
			pDomain						=	new CDomain  ();

			nError						=	pDomain->SetDomain  (sDomain, nAction, fDACL, fSACL);
			if (nError != RTN_OK)
			{
				LogMessage (Error, L"AddDomain: Domain name <" + sDomain +  L"> is probably incorrect.");

				delete pDomain;
				pDomain = NULL;

				return nError;
			}

			if (! sNewDomain.IsEmpty ())
			{
				pNewDomain				=	new CDomain  ();

				nError					=	pNewDomain->SetDomain  (sNewDomain, nAction, fDACL, fSACL);
				if (nError != RTN_OK)
				{
					LogMessage (Error, L"AddDomain: Domain name <" + sNewDomain +  L"> is probably incorrect.");

					delete pNewDomain;
					pNewDomain = NULL;

					return nError;
				}

				pDomain->m_oNewDomain	=	pNewDomain;
			}

			m_lstDomains.AddTail (pDomain);

			// Remember whether DACL and/or SACL have to be processed
			if (fDACL)
			{
				m_fDomainsProcessDACL	=	true;
			}
			if (fSACL)
			{
				m_fDomainsProcessSACL	=	true;
			}

			return RTN_OK;
		}
		else
		{
			LogMessage (Error, L"AddDomain: No domain specified.");

			return RTN_ERR_PARAMS;
		}
	}
	catch (...)
	{
		delete pDomain;
		pDomain = NULL;
		delete pNewDomain;
		pNewDomain = NULL;

		return RTN_ERR_OUT_OF_MEMORY;
	}
}


//
// SetIgnoreErrors: Ignore errors, do NOT stop execution (unknown consequences!)
//
bool CSetACL::SetIgnoreErrors (bool fIgnoreErrors)
{
	m_fIgnoreErrors	=	fIgnoreErrors;

	return RTN_OK;
}


//
// SetOwner: Set an owner to be set by Run ()
//
DWORD CSetACL::SetOwner (CString sTrustee, bool fTrusteeIsSID)
{
	delete m_pOwner;				
	m_pOwner	= NULL;

	if (! sTrustee.IsEmpty ())
	{
		m_pOwner	=	new CTrustee (sTrustee, fTrusteeIsSID, ACTN_SETOWNER, false, false);

		return RTN_OK;
	}
	else
	{
		return RTN_ERR_PARAMS;
	}
}


//
// SetPrimaryGroup: Set the primary group to be set by Run ()
//
DWORD CSetACL::SetPrimaryGroup (CString sTrustee, bool fTrusteeIsSID)
{
	delete m_pPrimaryGroup;
	m_pPrimaryGroup = NULL;

	if (! sTrustee.IsEmpty ())
	{
		m_pPrimaryGroup	=	new CTrustee (sTrustee, fTrusteeIsSID, ACTN_SETGROUP, false, false);

		return RTN_OK;
	}
	else
	{
		return RTN_ERR_PARAMS;
	}
}


//
// SetRecursion: Set the recursion
//
DWORD CSetACL::SetRecursion (DWORD nRecursionType)
{
	if (! m_nObjectType)
	{
		m_nRecursionType	=	0;

		return RTN_ERR_OBJECT_NOT_SET;
	}

	if (m_nObjectType == SE_FILE_OBJECT)
	{
		if (nRecursionType == RECURSE_NO || nRecursionType == RECURSE_CONT || nRecursionType == RECURSE_OBJ || nRecursionType == RECURSE_CONT_OBJ)
		{
			m_nRecursionType	=	nRecursionType;

			return RTN_OK;
		}
	}
	else if (m_nObjectType == SE_REGISTRY_KEY)
	{
		if (nRecursionType == RECURSE_CONT_OBJ)
		{
			nRecursionType	=	RECURSE_CONT;
		}

		m_nRecursionType	=	nRecursionType;

		return RTN_OK;
	}

	return RTN_OK;
}


//
// SetObjectFlags: Set flags specific to the object
//
DWORD CSetACL::SetObjectFlags (DWORD nDACLProtected, DWORD nSACLProtected, bool fDACLResetChildObjects, bool fSACLResetChildObjects)
{
	if (CheckInhFromParent (nDACLProtected) && CheckInhFromParent (nSACLProtected))
	{
		m_nDACLProtected				=	nDACLProtected;
		m_fDACLResetChildObjects	=	fDACLResetChildObjects;

		m_nSACLProtected				=	nSACLProtected;
		m_fSACLResetChildObjects	=	fSACLResetChildObjects;

		return RTN_OK;
	}
	else
	{
		m_nDACLProtected				=	INHPARNOCHANGE;
		m_fDACLResetChildObjects	=	false;

		m_nSACLProtected				=	INHPARNOCHANGE;
		m_fSACLResetChildObjects	=	false;

		return RTN_ERR_PARAMS;
	}
}


//
// SetListOptions: Set the options for ACL listing
//
DWORD CSetACL::SetListOptions (DWORD nListFormat, DWORD nListWhat, bool fListInherited, DWORD nListNameSID, bool fListCleanOutput)
{
	if (nListWhat > 0 && nListWhat <= (ACL_DACL + ACL_SACL + SD_OWNER + SD_GROUP) && (nListFormat == LIST_SDDL || nListFormat == LIST_CSV || nListFormat == LIST_TAB) && nListNameSID >= LIST_NAME && nListNameSID <= LIST_NAME_SID)
	{
		m_nListFormat			= nListFormat;
		m_nListWhat				= nListWhat;
		m_nListNameSID			= nListNameSID;
		m_fListInherited		= fListInherited;
		m_fListCleanOutput	= fListCleanOutput;
		return RTN_OK;
	}

	m_nListFormat			= LIST_CSV;
	m_nListWhat				= ACL_DACL;
	m_nListNameSID			= LIST_NAME;
	m_fListInherited		= false;
	m_fListCleanOutput	= false;
	return RTN_ERR_LIST_OPTIONS;
}


//
// GetLastErrorMessage: Return the error string in the system's language of either the error code passed in or from m_nAPIError
//
CString CSetACL::GetLastErrorMessage (DWORD nError)
{
	CString	sMessage;
	HMODULE	hModule			= NULL;
	DWORD		nFormatFlags	= FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_FROM_SYSTEM;

	if (! nError)
	{
		if (! m_nAPIError)
		{
			nError	=	GetLastError ();
		}
		else
		{
			nError	=	m_nAPIError;
		}
	}

	//
	// If nError is in the network range, load the message source
	//
	if (nError >= NERR_BASE && nError <= MAX_NERR)
	{
		hModule	=	LoadLibraryEx (L"netmsg.dll", NULL, LOAD_LIBRARY_AS_DATAFILE);

		if (hModule)
		{
			nFormatFlags	|=	FORMAT_MESSAGE_FROM_HMODULE;
		}
	}

	//
	// Call FormatMessage () to get the message text from the system or the supplied module handle
	//
	DWORD nChars = FormatMessage (nFormatFlags, hModule, nError, MAKELANGID (LANG_NEUTRAL, SUBLANG_DEFAULT), sMessage.GetBuffer (1024), sizeof (TCHAR) * 1024, NULL);
	sMessage.ReleaseBuffer ();

	if (nChars == 0)
	{
		// FormatMessage could not look up the error, try interpreting it as a HRESULT (COM error)
		sMessage = _com_error(nError).ErrorMessage();
	}

	// Replace newlines added by FormatMessage
	sMessage.TrimRight (L"\r\n");

	// If we loaded a message source, unload it
	if (hModule)
	{
		FreeLibrary (hModule);
	}

	return sMessage;
}


//
// Run: Do it: apply all settings.
//
DWORD CSetACL::Run ()
{
	DWORD		nError				=	RTN_OK;
	CString	sInternalError;
	CString	sAPIError;
	CString	sErrorMessage;

	try
	{
		// Do things like looking up binary SIDs for every trustee
		nError	=	Prepare ();

		if (nError != RTN_OK)
		{
			throw nError;
		}

		//
		// Perform requested action(s)
		//
		// But first check if there IS an action
		if (m_nAction == 0)
		{
			throw (DWORD) RTN_ERR_NO_ACTN_SPECIFIED;
		}

		if (m_nAction & ACTN_RESTORE)
		{
			// Do the work
			nError	=	DoActionRestore ();

			if (nError != RTN_OK)
			{
				throw nError;
			}
		}

		if (m_nAction & ACTN_ADDACE || m_nAction & ACTN_SETOWNER || m_nAction & ACTN_SETGROUP || m_nAction & ACTN_SETINHFROMPAR || m_nAction & ACTN_CLEARDACL || m_nAction & ACTN_CLEARSACL || m_nAction & ACTN_TRUSTEE || m_nAction & ACTN_DOMAIN && nError == RTN_OK)
		{
			// Do the work
			nError	=	DoActionWrite ();

			if (nError != RTN_OK)
			{
				throw nError;
			}
		}

		if (m_nAction & ACTN_RESETCHILDPERMS && nError == RTN_OK)
		{
			// Set sub-object-only processing to on.
			m_fProcessSubObjectsOnly	=	true;

			//
			// Set some variables to the correct values for this operation, but first back them up so we can restore the original values later.
			//
			// 1. Backup
			DWORD	nAction					=	m_nAction;
			DWORD	nRecursionType			=	m_nRecursionType;
			DWORD nDACLProtected			=	m_nDACLProtected;
			DWORD nSACLProtected			=	m_nSACLProtected;
			DWORD	nDACLEntries			=	m_nDACLEntries;
			DWORD	nSACLEntries			=	m_nSACLEntries;

			// 2. Set new values: do nothing unless explicitly specified (by parameter "-rst")
			m_nAction						=	0;
			m_nRecursionType				=	RECURSE_CONT_OBJ;
			m_nDACLProtected				=	INHPARNOCHANGE;
			m_nSACLProtected				=	INHPARNOCHANGE;
			m_nDACLEntries					=	0;
			m_nSACLEntries					=	0;

			if (m_fDACLResetChildObjects)
			{
				m_nAction					|= ACTN_SETINHFROMPAR | ACTN_CLEARDACL;
				m_nDACLProtected			=	INHPARYES;
			}
			if (m_fSACLResetChildObjects)
			{
				m_nAction					|= ACTN_SETINHFROMPAR | ACTN_CLEARSACL;
				m_nSACLProtected			=	INHPARYES;
			}

			// 3. Do the work
			if (m_nAction)
			{
				nError						=	DoActionWrite ();

				if (nError != RTN_OK)
				{
					throw nError;
				}
			}
			else
			{
				LogMessage (Warning, L"Action 'reset children' was used without specifying whether to reset the DACL, SACL, or both. Nothing was reset.");
			}

			// 4. Restore the variables we backed up.
			m_nAction						=	nAction;
			m_nRecursionType				=	nRecursionType;
			m_nDACLProtected				=	nDACLProtected;
			m_nSACLProtected				=	nSACLProtected;
			m_nDACLEntries					=	nDACLEntries;
			m_nSACLEntries					=	nSACLEntries;

			// Set sub-object-only processing back to default state.
			m_fProcessSubObjectsOnly	=	false;
		}

		if (m_nAction & ACTN_LIST && nError == RTN_OK)
		{
			// Do the work
			nError	=	DoActionList ();

			if (nError != RTN_OK)
			{
				throw nError;
			}
		}

		// No Errors occured
		sErrorMessage		=	L"\nSetACL finished successfully.";
	}
	catch (DWORD nRunError)
	{
		if (nRunError || m_nAPIError)
		{
			sErrorMessage	=	L"\nSetACL finished with error(s): ";

			if (nRunError)
			{
				// Get the internal error message from the resource table
				CString sText = L"\nSetACL error message: [could not retrieve]";
				if (sInternalError.LoadString (nRunError))
				{
					sText	=	L"\nSetACL error message: " + sInternalError;
				}
				sErrorMessage += sText;
			}

			if (m_nAPIError)
			{
				// Get the API error from the OS
				sAPIError	=	GetLastErrorMessage ();
				sErrorMessage	+=	L"\nOperating system error message: " + sAPIError;
			}
		}
	}

	// Notify caller of result
	LogMessage (NoSeverity, sErrorMessage);

	return nError;
}


//
// Prepare: Prepare for execution.
//
DWORD CSetACL::Prepare ()
{
	DWORD	nError;

	// 
	// Sanity checks
	//
	if (m_sObjectPath.IsEmpty ())
	{
		LogMessage (Error, L"The object path was not specified.");

		return RTN_ERR_OBJECT_NOT_SET;
	}
	if (m_nObjectType == SE_UNKNOWN_OBJECT_TYPE)
	{
		LogMessage (Error, L"The object type was not specified.");

		return RTN_ERR_OBJECT_NOT_SET;
	}

	// Check the type and version of the OS we are running on
	OSVERSIONINFO osviVersion;
	SecureZeroMemory (&osviVersion, sizeof (OSVERSIONINFO));
	osviVersion.dwOSVersionInfoSize = sizeof (OSVERSIONINFO);
	if (GetVersionEx (&osviVersion))
	{
		bool fVersionOK = (osviVersion.dwPlatformId == VER_PLATFORM_WIN32_NT) && (osviVersion.dwMajorVersion > 4);
		if (! fVersionOK)
		{
			LogMessage (Error, L"SetACL works only on NT based operating systems newer than NT4.");

			return RTN_ERR_OS_NOT_SUPPORTED;
		}
	}
	else
	{
		LogMessage (Information, L"The version of your operating system could not be determined. SetACL may not work correctly.");
	}

	// Try to enable the backup privilege. This allows us to read file system objects that we do not have permission for.
	if (SetPrivilege (SE_BACKUP_NAME, true, false) != RTN_OK)
	{
		LogMessage (Warning, L"Privilege 'Back up files and directories' could not be enabled. SetACL's powers are restricted. Better run SetACL with admin rights.");
	}

	// Also enable the restore privilege which lets us set the owner to any (!) user/group, not only administrators
	if (SetPrivilege (SE_RESTORE_NAME, true, false) != RTN_OK)
	{
		LogMessage (Warning, L"Privilege 'Restore files and directories' could not be enabled. SetACL's powers are restricted. Better run SetACL with admin rights.");
	}

	// Enable a privilege which lets us set the owner regardless of current permissions
	if (SetPrivilege (SE_TAKE_OWNERSHIP_NAME, true, false) != ERROR_SUCCESS)
	{
		LogMessage (Warning, L"Privilege 'Take ownership of files or other objects' could not be enabled. SetACL's powers are restricted. Better run SetACL with admin rights.");
	}

	//
	// Look up the SID of every trustee set
	//
	// First iterate through the entries in m_lstACEs
	POSITION	pos	=	m_lstACEs.GetHeadPosition ();

	while (pos)
	{
		CACE*	pACE	=	m_lstACEs.GetNext (pos);

		if (pACE && pACE->m_pTrustee->LookupSID () != RTN_OK)
		{
			return RTN_ERR_LOOKUP_SID;
		}
	}

	// Next, iterate through the entries in m_lstTrustees
	pos	=	m_lstTrustees.GetHeadPosition ();

	while (pos)
	{
		CTrustee*	pTrustee	=	m_lstTrustees.GetNext (pos);

		if (pTrustee && pTrustee->LookupSID () != RTN_OK)
		{
			return RTN_ERR_LOOKUP_SID;
		}
		if (pTrustee && pTrustee->m_oNewTrustee->LookupSID () != RTN_OK)
		{
			return RTN_ERR_LOOKUP_SID;
		}
	}

	// Finally process the owner and primary group
	if (m_pOwner && m_pOwner->LookupSID () != RTN_OK)
	{
		return RTN_ERR_LOOKUP_SID;
	}

	if (m_pPrimaryGroup && m_pPrimaryGroup->LookupSID () != RTN_OK)
	{
		return RTN_ERR_LOOKUP_SID;
	}

	//
	// Set the correct permission and inheritance values for each ACE entry
	//
	nError = DetermineACEAccessMasks ();
	if (nError != ERROR_SUCCESS)
	{
		return nError;
	}

	// If the registry is to be processed: correct the registry path (ie. "hklm\software" -> "machine\software")
	if (m_nObjectType == SE_REGISTRY_KEY)
	{
		HKEY hKey	=	NULL;
		nError		=	RegKeyFixPathAndOpen (m_sObjectPath, hKey, true, 0);

		if (nError != RTN_OK)
		{
			return nError;
		}
	}

	//
	// Correct file system paths. We use Unicode and have to use another syntax to overcome the MAX_PATH restriction (260 chars):
	//
	// c:						local file system root, does not persist when the system is restarted	->	\\?\c:
	// c:\data				"normal" paths																				->	\\?\c:\
	// \\server\share		UNC paths																					->	\\?\UNC\server\share
	//
	if (m_nObjectType	== SE_FILE_OBJECT)
	{
		// Build the "long" version of the path
		BuildLongUnicodePath (m_sObjectPath);
	}

	return RTN_OK;
}


//
// CheckAction: Check if an action is valid
//
bool CSetACL::CheckAction (DWORD nAction)
{
	switch (nAction)
	{
	case ACTN_ADDACE:
	case ACTN_LIST:
	case ACTN_SETOWNER:
	case ACTN_SETGROUP:
	case ACTN_CLEARDACL:
	case ACTN_CLEARSACL:
	case ACTN_TRUSTEE:
	case ACTN_DOMAIN:
	case ACTN_SETINHFROMPAR:
	case ACTN_RESETCHILDPERMS:
	case ACTN_RESTORE:
		return true;
	default:
		return false;
	}
}


//
// CheckInhFromParent: Check if inheritance from parent flags are valid
//
bool CSetACL::CheckInhFromParent (DWORD nInheritance)
{
	switch (nInheritance)
	{
	case INHPARNOCHANGE:
	case INHPARYES:
	case INHPARCOPY:
	case INHPARNOCOPY:
		return true;
	default:
		return false;
	}
}


//
// CheckInheritance: Check if inheritance flags are valid
//
bool CSetACL::CheckInheritance (DWORD nInheritance)
{
	DWORD	nAllFlags	=	SUB_OBJECTS_ONLY_INHERIT | SUB_CONTAINERS_ONLY_INHERIT | INHERIT_NO_PROPAGATE | INHERIT_ONLY;

	if (nInheritance >= NO_INHERITANCE && nInheritance <= nAllFlags)
	{
		return true;
	}
	else
	{
		return false;
	}
}


//
// CheckACEType: Check if an ACE access mode (deny, set...) is valid
//
bool CSetACL::CheckACEAccessMode (DWORD nAccessMode, DWORD nACLType)
{
	if (nACLType == ACL_DACL)
	{
		switch (nAccessMode)
		{
		case DENY_ACCESS:
		case REVOKE_ACCESS:
		case SET_ACCESS:
		case GRANT_ACCESS:
			return true;
		default:
			return false;
		}
	}
	else if (nACLType == ACL_SACL)
	{
		switch (nAccessMode)
		{
		case REVOKE_ACCESS:
		case SET_AUDIT_SUCCESS:
		case SET_AUDIT_FAILURE:
		case SET_AUDIT_FAILURE + SET_AUDIT_SUCCESS:
			return true;
		default:
			return false;
		}
	}

	return false;
}


//
// CheckFilterList: Check whether a certain path needs to be filtered out
//
bool CSetACL::CheckFilterList (CString sObjectPath)
{
	// Ignore case when searching for keywords
	sObjectPath.MakeLower ();

	POSITION pos	=	m_lstObjectFilter.GetHeadPosition ();

	while (pos)
	{
		CString	sKeyword	=	m_lstObjectFilter.GetNext (pos);

		sKeyword.MakeLower ();

		if (sObjectPath.Find (sKeyword) != -1)
		{
			// The path is on the filter list
			return true;
		}
	}

	return false;
}


//
// DetermineACEAccessMask: Set the access masks and inheritance values for all ACEs
//
DWORD CSetACL::DetermineACEAccessMasks ()
{
	CTypedPtrList<CPtrList, CACE*>	lstACEs2Add;

	//
	// Iterate through the entries in m_lstACEs
	//
	POSITION	pos	=	m_lstACEs.GetHeadPosition ();

	while (pos)
	{
		// Get the current list element
		CACE*	pACE	=	m_lstACEs.GetNext (pos);

		//
		// In ONE case a combination of access modes is allowed (audit success + audit fail). We have to copy
		// the ACE and set both modes separately because SetEntriesInAcl is rather picky about this.
		//
		if (pACE->m_nAccessMode == SET_AUDIT_FAILURE + SET_AUDIT_SUCCESS)
		{
			pACE->m_nAccessMode	=	SET_AUDIT_FAILURE;

			CACE*	pACE2				=	CopyACE (pACE);
			pACE2->m_nAccessMode	=	SET_AUDIT_SUCCESS;

			POSITION posNew		=	m_lstACEs.AddTail (pACE2);

			// Make sure the new element is processed even if this is the last element
			if (! pos)
			{
				pos					=	posNew;
			}
		}

		// Reset the access mask
		pACE->m_nAccessMask	=	0;

		//
		// Multiple permissions can be specified. Process them all
		//
		CStringArray	asPermissions;

		// Split the permission list
		if (! Split (L",", pACE->m_sPermission, asPermissions))
		{
			return RTN_ERR_PARAMS;
		}

		//
		// Loop through the permission list
		//
		for (int k = 0; k < asPermissions.GetSize (); k++)
		{
			//
			// FILE / DIRECTORY
			//
			if (m_nObjectType == SE_FILE_OBJECT)
			{
				if (asPermissions[k].CompareNoCase (L"read") == 0)
				{
					pACE->m_nAccessMask			|=	MY_DIR_READ_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"write") == 0)
				{
					pACE->m_nAccessMask			|= MY_DIR_WRITE_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"list_folder") == 0)
				{
					pACE->m_nAccessMask			|= MY_DIR_LIST_FOLDER_ACCESS;

					if (! pACE->m_fInhSpecified)
					{
						pACE->m_fInhSpecified	=	true;
						pACE->m_nInheritance		=	CONTAINER_INHERIT_ACE;
					}
				}
				else if (asPermissions[k].CompareNoCase (L"read_ex") == 0 || asPermissions[k].CompareNoCase (L"read_execute") == 0)
				{
					pACE->m_nAccessMask			|= MY_DIR_READ_EXECUTE_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"change") == 0)
				{
					pACE->m_nAccessMask			|= MY_DIR_CHANGE_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"profile") == 0)
				{
					pACE->m_nAccessMask			|= MY_DIR_CHANGE_ACCESS | WRITE_DAC;
				}
				else if (asPermissions[k].CompareNoCase (L"full") == 0)
				{
					pACE->m_nAccessMask			|= MY_DIR_FULL_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"traverse") == 0 || asPermissions[k].CompareNoCase (L"FILE_TRAVERSE") == 0)
				{
					pACE->m_nAccessMask			|= FILE_TRAVERSE;
				}
				else if (asPermissions[k].CompareNoCase (L"list_dir") == 0 || asPermissions[k].CompareNoCase (L"FILE_LIST_DIRECTORY") == 0)
				{
					pACE->m_nAccessMask			|= FILE_LIST_DIRECTORY;
				}
				else if (asPermissions[k].CompareNoCase (L"read_attr") == 0 || asPermissions[k].CompareNoCase (L"FILE_READ_ATTRIBUTES") == 0)
				{
					pACE->m_nAccessMask			|= FILE_READ_ATTRIBUTES;
				}
				else if (asPermissions[k].CompareNoCase (L"read_ea") == 0 || asPermissions[k].CompareNoCase (L"FILE_READ_EA") == 0)
				{
					pACE->m_nAccessMask			|= FILE_READ_EA;
				}
				else if (asPermissions[k].CompareNoCase (L"add_file") == 0 || asPermissions[k].CompareNoCase (L"FILE_ADD_FILE") == 0)
				{
					pACE->m_nAccessMask			|= FILE_ADD_FILE;
				}
				else if (asPermissions[k].CompareNoCase (L"add_subdir") == 0 || asPermissions[k].CompareNoCase (L"FILE_ADD_SUBDIRECTORY") == 0)
				{
					pACE->m_nAccessMask			|= FILE_ADD_SUBDIRECTORY;
				}
				else if (asPermissions[k].CompareNoCase (L"write_attr") == 0 || asPermissions[k].CompareNoCase (L"FILE_WRITE_ATTRIBUTES") == 0)
				{
					pACE->m_nAccessMask			|= FILE_WRITE_ATTRIBUTES;
				}
				else if (asPermissions[k].CompareNoCase (L"write_ea") == 0 || asPermissions[k].CompareNoCase (L"FILE_WRITE_EA") == 0)
				{
					pACE->m_nAccessMask			|= FILE_WRITE_EA;
				}
				else if (asPermissions[k].CompareNoCase (L"del_child") == 0 || asPermissions[k].CompareNoCase (L"FILE_DELETE_CHILD") == 0)
				{
					pACE->m_nAccessMask			|= FILE_DELETE_CHILD;
				}
				else if (asPermissions[k].CompareNoCase (L"delete") == 0)
				{
					pACE->m_nAccessMask			|= DELETE;
				}
				else if (asPermissions[k].CompareNoCase (L"read_dacl") == 0 || asPermissions[k].CompareNoCase (L"READ_CONTROL") == 0)
				{
					pACE->m_nAccessMask			|= READ_CONTROL;
				}
				else if (asPermissions[k].CompareNoCase (L"write_dacl") == 0 || asPermissions[k].CompareNoCase (L"WRITE_DAC") == 0)
				{
					pACE->m_nAccessMask			|= WRITE_DAC;
				}
				else if (asPermissions[k].CompareNoCase (L"write_owner") == 0)
				{
					pACE->m_nAccessMask			|= WRITE_OWNER;
				}
				else
				{
					return RTN_ERR_INV_DIR_PERMS;
				}

				//
				// When setting permissions (excluding DENY), always set the SYNCHRONIZE flag
				//
				if (pACE->m_nAccessMode != DENY_ACCESS)
				{
					pACE->m_nAccessMask			|=	SYNCHRONIZE;
				}

				//
				// Set default inheritance flags if none were specified
				//
				if (! pACE->m_fInhSpecified)
				{
					pACE->m_nInheritance			=	SUB_CONTAINERS_AND_OBJECTS_INHERIT;
				}
			}

			//
			// PRINTER
			//
			else if (m_nObjectType == SE_PRINTER)
			{
				if (asPermissions[k].CompareNoCase (L"print") == 0)
				{
					pACE->m_nAccessMask			|= MY_PRINTER_PRINT_ACCESS;

					if (! pACE->m_fInhSpecified)
					{
						pACE->m_nInheritance		= 0;
					}
				}
				else if (asPermissions[k].CompareNoCase (L"man_printer") == 0 || asPermissions[k].CompareNoCase (L"manage_printer") == 0)
				{
					pACE->m_nAccessMask			|= MY_PRINTER_MAN_PRINTER_ACCESS;

					if (! pACE->m_fInhSpecified)
					{
						pACE->m_nInheritance		= 0;
					}
				}
				else if (asPermissions[k].CompareNoCase (L"man_docs") == 0 || asPermissions[k].CompareNoCase (L"manage_documents") == 0)
				{
					// We have to set two different ACEs in order to obtain standard manage documents permissions;

					if (pACE->m_fInhSpecified)
					{
						// The user specified inheritance flags: warn him that is contraproductive with man_docs
						LogMessage (Warning, L"You specified inheritance flags, which is incompatible with man_docs. Your flags are being ignored in order to be able to set standard manage documents permissions.");
					}

					CACE*			pACE2			=	CopyACE (pACE);

					lstACEs2Add.AddTail (pACE2);

					pACE->m_nAccessMask		=	READ_CONTROL;
					pACE->m_nInheritance		=	CONTAINER_INHERIT_ACE | INHERIT_ONLY_ACE;

					pACE2->m_nAccessMask		=	MY_PRINTER_MAN_DOCS_ACCESS;
					pACE2->m_nInheritance	=	OBJECT_INHERIT_ACE | INHERIT_ONLY_ACE;
				}
				else if (asPermissions[k].CompareNoCase (L"full") == 0)
				{
					// We have to set two different ACEs, in order to obtain standard manage documents permissions;
					// Only do it if inheritance was NOT specified by the user. If it was, assume he knows what he is doing

					if (! pACE->m_fInhSpecified)
					{
						CACE*			pACE2			=	CopyACE (pACE);

						lstACEs2Add.AddTail (pACE2);

						pACE->m_nAccessMask		=	MY_PRINTER_MAN_PRINTER_ACCESS;
						pACE->m_nInheritance		=	0;

						pACE2->m_nAccessMask		=	MY_PRINTER_MAN_DOCS_ACCESS;
						pACE2->m_nInheritance	=	OBJECT_INHERIT_ACE | INHERIT_ONLY_ACE;
					}
					else
					{
						pACE->m_nAccessMask		|=	MY_PRINTER_MAN_PRINTER_ACCESS;
					}
				}
				else
				{
					return RTN_ERR_INV_PRN_PERMS;
				}
			}

			//
			// REGISTRY
			//
			else if (m_nObjectType == SE_REGISTRY_KEY)
			{
				if (asPermissions[k].CompareNoCase (L"read") == 0)
				{
					pACE->m_nAccessMask			|= MY_REG_READ_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"full") == 0)
				{
					pACE->m_nAccessMask			|= MY_REG_FULL_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"query_val") == 0 || asPermissions[k].CompareNoCase (L"KEY_QUERY_VALUE") == 0)
				{
					pACE->m_nAccessMask			|= KEY_QUERY_VALUE;
				}
				else if (asPermissions[k].CompareNoCase (L"set_val") == 0 || asPermissions[k].CompareNoCase (L"KEY_SET_VALUE") == 0)
				{
					pACE->m_nAccessMask			|= KEY_SET_VALUE;
				}
				else if (asPermissions[k].CompareNoCase (L"create_subkey") == 0 || asPermissions[k].CompareNoCase (L"KEY_CREATE_SUB_KEY") == 0)
				{
					pACE->m_nAccessMask			|= KEY_CREATE_SUB_KEY;
				}
				else if (asPermissions[k].CompareNoCase (L"enum_subkeys") == 0 || asPermissions[k].CompareNoCase (L"KEY_ENUMERATE_SUB_KEYS") == 0)
				{
					pACE->m_nAccessMask			|= KEY_ENUMERATE_SUB_KEYS;
				}
				else if (asPermissions[k].CompareNoCase (L"notify") == 0 || asPermissions[k].CompareNoCase (L"KEY_NOTIFY") == 0)
				{
					pACE->m_nAccessMask			|= KEY_NOTIFY;
				}
				else if (asPermissions[k].CompareNoCase (L"create_link") == 0 || asPermissions[k].CompareNoCase (L"KEY_CREATE_LINK") == 0)
				{
					pACE->m_nAccessMask			|= KEY_CREATE_LINK;
				}
				else if (asPermissions[k].CompareNoCase (L"delete") == 0)
				{
					pACE->m_nAccessMask			|= DELETE;
				}
				else if (asPermissions[k].CompareNoCase (L"write_dacl") == 0 || asPermissions[k].CompareNoCase (L"WRITE_DAC") == 0)
				{
					pACE->m_nAccessMask			|= WRITE_DAC;
				}
				else if (asPermissions[k].CompareNoCase (L"write_owner") == 0)
				{
					pACE->m_nAccessMask			|= WRITE_OWNER;
				}
				else if (asPermissions[k].CompareNoCase (L"read_access") == 0 || asPermissions[k].CompareNoCase (L"READ_CONTROL") == 0)
				{
					pACE->m_nAccessMask			|= READ_CONTROL;
				}
				else
				{
					return RTN_ERR_INV_REG_PERMS;
				}

				//
				// Set default inheritance flags if none were specified
				//
				if (! pACE->m_fInhSpecified)
				{
					pACE->m_nInheritance			=	CONTAINER_INHERIT_ACE;
				}
			}

			//
			// SERVICE
			//
			else if (m_nObjectType == SE_SERVICE)
			{
				if (asPermissions[k].CompareNoCase (L"read") == 0)
				{
					pACE->m_nAccessMask			|= MY_SVC_READ_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"start_stop") == 0)
				{
					pACE->m_nAccessMask			|= MY_SVC_STARTSTOP_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"full") == 0)
				{
					pACE->m_nAccessMask			|= MY_SVC_FULL_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_CHANGE_CONFIG") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_CHANGE_CONFIG;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_ENUMERATE_DEPENDENTS") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_ENUMERATE_DEPENDENTS;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_INTERROGATE") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_INTERROGATE;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_PAUSE_CONTINUE") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_PAUSE_CONTINUE;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_QUERY_CONFIG") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_QUERY_CONFIG;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_QUERY_STATUS") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_QUERY_STATUS;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_START") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_START;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_STOP") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_STOP;
				}
				else if (asPermissions[k].CompareNoCase (L"SERVICE_USER_DEFINED_CONTROL") == 0)
				{
					pACE->m_nAccessMask			|= SERVICE_USER_DEFINED_CONTROL;
				}
				else
				{
					return RTN_ERR_INV_SVC_PERMS;
				}

				//
				// Set default inheritance flags if none were specified
				//
				if (! pACE->m_fInhSpecified)
				{
					pACE->m_nInheritance			=	0;
				}
			}

			//
			// SHARE
			//
			else if (m_nObjectType == SE_LMSHARE)
			{
				if (asPermissions[k].CompareNoCase (L"read") == 0)
				{
					pACE->m_nAccessMask			|= MY_SHARE_READ_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"change") == 0)
				{
					pACE->m_nAccessMask			|= MY_SHARE_CHANGE_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"full") == 0)
				{
					pACE->m_nAccessMask			|= MY_SHARE_FULL_ACCESS;
				}
				else
				{
					return RTN_ERR_INV_SHR_PERMS;
				}

				//
				// Set default inheritance flags if none were specified
				//
				if (! pACE->m_fInhSpecified)
				{
					pACE->m_nInheritance			=	0;
				}
			}
			//
			// WMI NAMESPACE
			//
			else if (m_nObjectType == SE_WMIGUID_OBJECT)
			{
				if (asPermissions[k].CompareNoCase (L"enable_account") == 0 || asPermissions[k].CompareNoCase(L"WBEM_ENABLE") == 0)
				{
					pACE->m_nAccessMask			|= MY_WMI_ENABLE_ACCOUNT;
				}
				else if (asPermissions[k].CompareNoCase (L"execute") == 0)
				{
					pACE->m_nAccessMask			|= MY_WMI_EXECUTE_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"remote_access") == 0)
				{
					pACE->m_nAccessMask			|= MY_WMI_REMOTE_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"full") == 0)
				{
					pACE->m_nAccessMask			|= MY_WMI_FULL_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"WBEM_METHOD_EXECUTE") == 0)
				{
					pACE->m_nAccessMask			|= WBEM_METHOD_EXECUTE;
				}
				else if (asPermissions[k].CompareNoCase (L"WBEM_FULL_WRITE_REP") == 0)
				{
					pACE->m_nAccessMask			|= WBEM_FULL_WRITE_REP;
				}
				else if (asPermissions[k].CompareNoCase (L"WBEM_PARTIAL_WRITE_REP") == 0)
				{
					pACE->m_nAccessMask			|= WBEM_PARTIAL_WRITE_REP;
				}
				else if (asPermissions[k].CompareNoCase (L"WBEM_WRITE_PROVIDER") == 0)
				{
					pACE->m_nAccessMask			|= WBEM_WRITE_PROVIDER;
				}
				else if (asPermissions[k].CompareNoCase (L"WBEM_REMOTE_ACCESS") == 0)
				{
					pACE->m_nAccessMask			|= WBEM_REMOTE_ACCESS;
				}
				else if (asPermissions[k].CompareNoCase (L"WBEM_RIGHT_SUBSCRIBE") == 0)
				{
					pACE->m_nAccessMask			|= WBEM_RIGHT_SUBSCRIBE;
				}
				else if (asPermissions[k].CompareNoCase (L"WBEM_RIGHT_PUBLISH") == 0)
				{
					pACE->m_nAccessMask			|= WBEM_RIGHT_PUBLISH;
				}
				else
				{
					return RTN_ERR_INV_WMI_PERMS;
				}

				//
				// Set default inheritance flags if none were specified
				//
				if (! pACE->m_fInhSpecified)
				{
					pACE->m_nInheritance			=	CONTAINER_INHERIT_ACE;
				}
			}

			//
			// Never set the SYNCHRONIZE flag on audit ACEs or ANY type of access will be audited, not only the one specified.
			//
			if (pACE->m_nAccessMode == SET_AUDIT_FAILURE || pACE->m_nAccessMode == SET_AUDIT_SUCCESS)
			{
				pACE->m_nAccessMask				&=	~SYNCHRONIZE;
			}
		}
	}

	// Maybe additional ACEs were generated. Append them to the main ACE list.
	m_lstACEs.AddTail (&lstACEs2Add);

	// Remove all temporary list entries. Objects pointed to by the list members are NOT destroyed!
	lstACEs2Add.RemoveAll ();

	return RTN_OK;
}


//
// CopyACE: Copy an existing ACE
//
CACE* CSetACL::CopyACE (CACE* pACE)
{
	if (! pACE)
	{
		return NULL;
	}

	CTrustee*	pTrustee		=	new CTrustee  (pACE->m_pTrustee->m_sTrustee, pACE->m_pTrustee->m_fTrusteeIsSID, pACE->m_pTrustee->m_nAction, pACE->m_pTrustee->m_fDACL, pACE->m_pTrustee->m_fSACL);

	if (! pTrustee) return NULL;

	pTrustee->m_psidTrustee	=	CopySID (pACE->m_pTrustee->m_psidTrustee);

	if (! pTrustee->m_psidTrustee) return NULL;

	CACE*			pACENew		=	new CACE (pTrustee, pACE->m_sPermission, pACE->m_nInheritance, pACE->m_fInhSpecified, pACE->m_nAccessMode, pACE->m_nACLType);

	if (! pACENew) return NULL;

	pACENew->m_nAccessMask	=	pACE->m_nAccessMask;

	// Adjust the internal ACE count
	if (pACE->m_nACLType == ACL_DACL)
	{
		m_nDACLEntries++;
	}
	else if (pACE->m_nACLType == ACL_SACL)
	{
		m_nSACLEntries++;
	}

	return pACENew;
}


//
// DoActionList: Create a permission listing
//
DWORD CSetACL::DoActionList ()
{
	DWORD	nError;

	// Close the output file should it be open
	if (m_fhBackupRestoreFile)
	{
		fclose (m_fhBackupRestoreFile);
		m_fhBackupRestoreFile	=	NULL;
	}

	// Was an output file specified?
	if (! m_sBackupRestoreFile.IsEmpty ())
	{
		// Open the output file
		errno_t	nErr	= 0;
		nErr	=	_tfopen_s (&m_fhBackupRestoreFile, m_sBackupRestoreFile, L"wb");

		if (m_fhBackupRestoreFile != NULL && nErr == 0)
		{
			// Write the unicode BOM (byte order mark)
			WCHAR cBOM	=	0xFEFF;
			fwrite (&cBOM, sizeof (WCHAR), 1, m_fhBackupRestoreFile);
		}
		else
		{
			return RTN_ERR_OPEN_LOGFILE;
		}
	}

	if (m_nObjectType == SE_FILE_OBJECT)
	{
		nError =	RecurseDirs (m_sObjectPath, &CSetACL::ListSD);
	}
	else if (m_nObjectType == SE_REGISTRY_KEY)
	{
		nError =	RecurseRegistry (m_sObjectPath, &CSetACL::ListSD);
	}
	else
	{
		nError =	ListSD (m_sObjectPath, false);
	}

	// Close the output file
	if (m_fhBackupRestoreFile)
	{
		fclose (m_fhBackupRestoreFile);
		m_fhBackupRestoreFile	=	NULL;
	}

	return nError;
}


//
// DoActionRestore: Restore permissions from a file
//
DWORD CSetACL::DoActionRestore ()
{
	DWORD						nError			=	RTN_OK;
	PACL						paclDACL			=	NULL;
	PACL						paclSACL			=	NULL;
	PSID						psidOwner		=	NULL;
	PSID						psidGroup		=	NULL;
	PSECURITY_DESCRIPTOR	pSDSelfRel		=	NULL;
	PSECURITY_DESCRIPTOR	pSDAbsolute		=	NULL;
	CSD						csdSD (this);
	CString					sLineIn;
	CStdioFile				oFile;


	// Was an input file specified?
	if (m_sBackupRestoreFile.IsEmpty ())
	{
		return RTN_ERR_NO_LOGFILE;
	}

	// Close input file should it be open
	if (m_fhBackupRestoreFile)
	{
		fclose (m_fhBackupRestoreFile);
		m_fhBackupRestoreFile	=	NULL;
	}

	// Open via stream to enable reading UNICODE (http://blog.kalmbachnet.de/?postid=105)
	errno_t nErr	= _tfopen_s (&m_fhBackupRestoreFile, m_sBackupRestoreFile, L"rt, ccs=UNICODE");
	if (nErr != 0)
		return RTN_ERR_OPEN_LOGFILE;

	// Now assign the stream to the CStdioFile object
	oFile.m_pStream	= m_fhBackupRestoreFile;

	LogMessage (Information, L"Input file for restore operation opened: '" + m_sBackupRestoreFile + L"'");

	// Read the entire file...
	CStringArray asLines;
	while (oFile.ReadString(sLineIn))
	{
		asLines.Add(sLineIn);
	}

	//
	// Process the input file from end to beginning
	//
	for (INT_PTR i = asLines.GetCount() - 1; i >= 0; i--)
	{
		// Free buffers that may have been allocated in the previous iteration
		if (pSDSelfRel)	{LocalFree (pSDSelfRel);	pSDSelfRel	=	NULL;}
		if (pSDAbsolute)	{LocalFree (pSDAbsolute);	pSDAbsolute	=	NULL;}
		if (paclDACL)		{LocalFree (paclDACL);		paclDACL		=	NULL;}
		if (paclSACL)		{LocalFree (paclSACL);		paclSACL		=	NULL;}
		if (psidOwner)		{LocalFree (psidOwner);		psidOwner	=	NULL;}
		if (psidGroup)		{LocalFree (psidGroup);		psidGroup	=	NULL;}

		DWORD						nBufSD			=	0;
		DWORD						nBufDACL			=	0;
		DWORD						nBufSACL			=	0;
		DWORD						nBufOwner		=	0;
		DWORD						nBufGroup		=	0;
		CString					sObjectPath;
		CString					sSDDL;
		SE_OBJECT_TYPE			nObjectType		=	SE_UNKNOWN_OBJECT_TYPE;
		SECURITY_INFORMATION	siSecInfo		=	0;

		// The current line
		sLineIn = asLines[i];

		// Remove unnecessary characters from the beginning and the end
		sLineIn.TrimRight (TEXT ("\r\n\"\t "));
		sLineIn.TrimLeft (TEXT ("\"\t "));

		// Ignore empty lines
		if (sLineIn.IsEmpty ())
		{
			continue;
		}

		// Split the line into object path, object type and SDDL string
		DWORD nFind1	=	sLineIn.Find (L"\",");
		sObjectPath		=	sLineIn.Left (nFind1);
		DWORD nFind2	=	sLineIn.Find (L",\"", nFind1 + 2);
		if (nFind2 == -1)
		{
			nFind2		=	sLineIn.Find (L",", nFind1 + 2);
			nObjectType	=	(SE_OBJECT_TYPE) _ttoi (sLineIn.Mid (nFind1 + 2, nFind2 - nFind1));
			sSDDL			=	sLineIn.Right (sLineIn.GetLength () - nFind2 - 1);
		}
		else
		{
			nObjectType	=	(SE_OBJECT_TYPE) _ttoi (sLineIn.Mid (nFind1 + 2, nFind2 - nFind1 - 1));
			sSDDL			=	sLineIn.Right (sLineIn.GetLength () - nFind2 - 2);
		}
		sSDDL.MakeUpper ();

		// Check if the SDDL string is empty
		if (sSDDL.IsEmpty ())
		{
			// Nothing to do
			LogMessage (Information, L"Omitting SD of: <" + sObjectPath + L"> because neither owner, group, DACL nor SACL were backed up.");

			continue;
		}

		// Check if the current path is on the filter list. If yes -> next line
		if (CheckFilterList (sObjectPath))
		{
			// Notify caller of omission
			LogMessage (Information, L"Omitting SD of: <" + sObjectPath + L"> because a filter keyword matched.");

			continue;
		}

		// Notify caller of progress
		LogMessage (Information, L"Restoring SD of: <" + sObjectPath + L">");

		//
		// Which parts of the SD do we set? Also determine DACL and SACL flags.
		//
		CString	sFlags;

		DWORD		nFind		=	sSDDL.Find (L"D:");
		if (nFind != -1)
		{
			// Process the DACL
			siSecInfo		|=	DACL_SECURITY_INFORMATION;

			// Find the flags
			DWORD	nFlags	=	UNPROTECTED_DACL_SECURITY_INFORMATION;
			DWORD	nFind3	=	sSDDL.Find (L"(", nFind + 2);

			if (nFind3 != -1)
			{
				sFlags		=	sSDDL.Mid (nFind + 2, nFind3 - nFind - 2);
			}
			else
			{
				// This is an empty DACL
				sFlags		=	sSDDL.Mid (nFind + 2);
			}

			// Check for the protection flag
			if (sFlags.Find (L"P") != -1)
			{
				nFlags		=	PROTECTED_DACL_SECURITY_INFORMATION;
			}

			// Set the appropriate flag(s)
			siSecInfo		|=	nFlags;
		}

		nFind					=	sSDDL.Find (L"S:");
		if (nFind != -1)
		{
			// Process the SACL
			siSecInfo		|=	SACL_SECURITY_INFORMATION;

			// Find the flags
			DWORD	nFlags	=	UNPROTECTED_SACL_SECURITY_INFORMATION;
			DWORD	nFind3	=	sSDDL.Find (L"(", nFind + 2);

			if (nFind3 != -1)
			{
				sFlags		=	sSDDL.Mid (nFind + 2, nFind3 - nFind - 2);
			}
			else
			{
				// This is an empty SACL
				sFlags		=	sSDDL.Mid (nFind + 2);
			}

			// Check for the protection flag
			if (sFlags.Find (L"P") != -1)
			{
				nFlags		=	PROTECTED_SACL_SECURITY_INFORMATION;
			}

			// Set the appropriate flag(s)
			siSecInfo		|=	nFlags;
		}
		if (sSDDL.Find (L"O:") != -1)
		{
			// Process the owner
			siSecInfo		|=	OWNER_SECURITY_INFORMATION;
		}
		if (sSDDL.Find (L"G:") != -1)
		{
			// Process the primary group
			siSecInfo		|=	GROUP_SECURITY_INFORMATION;
		}

		// Convert the SDDL string into an SD
		if (! ConvertStringSecurityDescriptorToSecurityDescriptor (sSDDL, SDDL_REVISION_1, &pSDSelfRel, NULL))
		{
			if (m_fIgnoreErrors)
			{
				LogMessage (Error, L"Restoring SD of <" + sObjectPath + ">: " + GetLastErrorMessage (GetLastError ()));

				continue;
			}
			else
			{
				nError		=	RTN_ERR_CONVERT_SD;
				m_nAPIError	=	GetLastError ();

				goto CleanUp;
			}
		}

		// Convert the self-relative SD into an absolute SD:
		// 1. Determine the size of the buffers needed
		MakeAbsoluteSD (pSDSelfRel, pSDAbsolute, &nBufSD, paclDACL, &nBufDACL, paclSACL, &nBufSACL, psidOwner, &nBufOwner, psidGroup, &nBufGroup);
		if (GetLastError () != ERROR_INSUFFICIENT_BUFFER)
		{
			if (m_fIgnoreErrors)
			{
				LogMessage (Error, L"Restoring SD of <" + sObjectPath + L">: " + GetLastErrorMessage (GetLastError ()));

				continue;
			}
			else
			{
				nError		=	RTN_ERR_CONVERT_SD;
				m_nAPIError	=	GetLastError ();

				goto CleanUp;
			}
		}

		// 2. Allocate the buffers
		pSDAbsolute		=	(PSECURITY_DESCRIPTOR)	LocalAlloc (LPTR, nBufSD);
		paclDACL			=	(PACL)						LocalAlloc (LPTR, nBufDACL);
		paclSACL			=	(PACL)						LocalAlloc (LPTR, nBufSACL);
		psidOwner		=	(PSID)						LocalAlloc (LPTR, nBufOwner);
		psidGroup		=	(PSID)						LocalAlloc (LPTR, nBufGroup);

		// 3. Do the conversion
		if (! MakeAbsoluteSD (pSDSelfRel, pSDAbsolute, &nBufSD, paclDACL, &nBufDACL, paclSACL, &nBufSACL, psidOwner, &nBufOwner, psidGroup, &nBufGroup))
		{
			if (m_fIgnoreErrors)
			{
				LogMessage (Error, L"Restoring SD of <" + sObjectPath + L">: " + GetLastErrorMessage (GetLastError ()));

				continue;
			}
			else
			{
				nError		=	RTN_ERR_CONVERT_SD;
				m_nAPIError	=	GetLastError ();

				goto CleanUp;
			}
		}

		// Determine if this object is a container
		bool fIsContainer	 = false;
		if (nObjectType == SE_FILE_OBJECT)
		{
			fIsContainer	= IsDirectory (sObjectPath);
		}
		else if (nObjectType == SE_REGISTRY_KEY)
		{
			fIsContainer	= true;
		}

		// Set the SD on the object.
		nError			=	csdSD.SetSD (sObjectPath, nObjectType, siSecInfo, paclDACL, paclSACL, psidOwner, psidGroup, false, fIsContainer);
		m_nAPIError		=	csdSD.m_nAPIError;

		if (nError != RTN_OK)
		{
			LogMessage (Error, L"Writing SD to <" + sObjectPath + L"> failed with: " + GetLastErrorMessage (m_nAPIError));

			if (m_nAPIError == ERROR_ACCESS_DENIED || m_nAPIError == ERROR_FILE_NOT_FOUND || m_fIgnoreErrors)
			{
				// The error can be ignored -> suppress and next line
				m_nAPIError	=	ERROR_SUCCESS;
				nError		=	RTN_OK;

				continue;
			}

			goto CleanUp;
		}
	}

CleanUp:

	// If errors are to be ignored (param "-ignoreerr" set), then do NOT return an error
	if (nError == RTN_ERR_IGNORED)
	{
		m_nAPIError	=	ERROR_SUCCESS;
		nError		=	RTN_OK;
	}

	// Close the input file
	if (m_fhBackupRestoreFile)
	{
		fclose (m_fhBackupRestoreFile);
		m_fhBackupRestoreFile	=	NULL;
	}

	// Free Memory
	if (pSDSelfRel)	{LocalFree (pSDSelfRel);	pSDSelfRel	=	NULL;}
	if (pSDAbsolute)	{LocalFree (pSDAbsolute);	pSDAbsolute	=	NULL;}
	if (paclDACL)		{LocalFree (paclDACL);		paclDACL		=	NULL;}
	if (paclSACL)		{LocalFree (paclSACL);		paclSACL		=	NULL;}
	if (psidOwner)		{LocalFree (psidOwner);		psidOwner	=	NULL;}
	if (psidGroup)		{LocalFree (psidGroup);		psidGroup	=	NULL;}

	return nError;
}


//
// DoActionWrite: Process actions that write to the SD
//
DWORD CSetACL::DoActionWrite ()
{
	if (m_nObjectType == SE_FILE_OBJECT)
	{
		return RecurseDirs (m_sObjectPath, &CSetACL::Write2SD);
	}
	else if (m_nObjectType == SE_REGISTRY_KEY)
	{
		return RecurseRegistry (m_sObjectPath, &CSetACL::Write2SD);
	}
	else
	{
		return Write2SD (m_sObjectPath, false);
	}
}


//
// RecurseDirs: Recurse a directory structure and call the function for every file / dir
//
DWORD CSetACL::RecurseDirs (CString sObjectPath, DWORD (CSetACL::*funcProcess) (CString sObjectPath, bool fIsContainer))
{
	WIN32_FIND_DATA		FindFileData;
	HANDLE					hFind;
	DWORD						nError				=	RTN_OK;
	CString					sSearchPattern;

	// Determine the container status of this object
	bool	fIsContainer	= IsDirectory (sObjectPath);

	// Process the object itself if specified
	if (m_nRecursionType & RECURSE_NO || 
		m_nRecursionType & RECURSE_CONT && fIsContainer ||
		m_nRecursionType & RECURSE_OBJ && (! fIsContainer))
	{
		nError	=	(this->*funcProcess) (sObjectPath, fIsContainer);
		if (nError != RTN_OK)
		{
			return nError;
		}
	}

	// Check for file system root and warn the user in case
	if (sObjectPath.GetLength() >= 2 && sObjectPath.Right(1) == ':')
   {
      if (sObjectPath.Mid(sObjectPath.GetLength() - 3, 1) == '\\' || sObjectPath.GetLength() == 2)
	   {
		   // The user specified a local file system root (e.g. "c:" or "\\?\c:"). Permissions set on that do not persist a reboot.
		   // Also, we cannot recurse in such a case.
		   if (m_nRecursionType & RECURSE_NO)
		   {
			   LogMessage (Warning, L"You just accessed permissions on a file system root, not on the root of the drive. These very special permissions do not persist a reboot and cannot be displayed in Explorer. You probably want to add a backslash to the path, e.g.: C:\\.");
		   }
		   else
		   {
			   LogMessage (Warning, L"You just accessed permissions on a file system root, not on the root of the drive. These very special permissions do not persist a reboot and cannot be displayed in Explorer. You probably want to add a backslash to the path, e.g.: C:\\. Please note that file system roots cannot be recursed.");
			   return RTN_WRN_RECURSION_IMPOSSIBLE;
		   }
	   }
   }

	// If no recursion is desired, we are done here
	if (m_nRecursionType & RECURSE_NO)
	{
		return nError;
	}

	//
	// Start recursively processing the path
	//
	// Build a search pattern for FindFirstFile
	sObjectPath.TrimRight ('\\');
	sSearchPattern	= sObjectPath + L"\\*.*";

	// Initiate the find operation.
	hFind = FindFirstFile (sSearchPattern, &FindFileData);
	if (hFind == INVALID_HANDLE_VALUE)
	{
		if (m_fIgnoreErrors)
		{
			m_nAPIError	= ERROR_SUCCESS;
			return RTN_OK;
		}
		else
		{
			// FindFirstFile reported an error: probably an invalid path
			m_nAPIError	=	GetLastError ();
			return RTN_ERR_FINDFILE;
		}
	}

	// Now loop through the results, including the first entry
	do
	{
		// The directories '.' and '..' are returned, too -> ignore
		if (_tcscmp (FindFileData.cFileName, L".") == 0 || _tcscmp (FindFileData.cFileName, L"..") == 0)
		{
			continue;
		}

		if (FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
		{
			// Check if this is a directory junction or a symbolic link
			if (FindFileData.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT)
			{
				if (FindFileData.dwReserved0 == IO_REPARSE_TAG_MOUNT_POINT || FindFileData.dwReserved0 == IO_REPARSE_TAG_SYMLINK)
				{
					// Do not follow junctions or links when recursing!
					continue;
				}
			}

			// Now recurse down the tree
			nError	=	RecurseDirs (sObjectPath + L"\\" + FindFileData.cFileName, funcProcess);
			if (nError != RTN_OK)
			{
				return nError;
			}
		}
		else
		{
			// This is a file: process it if requested
			if (m_nRecursionType & RECURSE_OBJ)
			{
				nError	=	(this->*funcProcess) (sObjectPath + L"\\" + FindFileData.cFileName, false);
				if (nError != RTN_OK)
				{
					return nError;
				}
			}
		}
	}
	while (FindNextFile (hFind, &FindFileData));

	FindClose (hFind);

	return nError;
}


//
// RegKeyFixPathAndOpen: Make sure we have a valid reg path and optionally open the key
//
DWORD CSetACL::RegKeyFixPathAndOpen (CString& sObjectPath, HKEY& hSubKey, bool fFixPathOnly, REGSAM samDesired)
{
	int								nLocalPart		=	0;
	int								nMainKey			=	0;
	HKEY								hMainKey			=	NULL;
	HKEY								hRemoteKey		=	NULL;
	HKEY								hOpenStd			=	NULL;
	HKEY								hOpenBckp		=	NULL;
	CString							sSubkeyPath;
	CString							sMainKey;
	CString							sMachinePath;
	CString							sLocalPath;

	// No backslash at the end of the path
	sObjectPath.TrimRight (L"\\");

	// Talk to the registry on the local or on a remote computer?
	if (sObjectPath.Left (2) == L"\\\\")
	{
		// Find end of computer name
		nLocalPart = sObjectPath.Find (L"\\", 2);

		// Nothing found? -> exit with error
		if (nLocalPart == -1)
		{
			return RTN_ERR_REG_PATH;
		}

		// Split into computer name and rest of path
		sMachinePath	=	sObjectPath.Left (nLocalPart);
		sLocalPath		=	sObjectPath.Right (sObjectPath.GetLength () - nLocalPart - 1);
	}
	else
	{
		// sMachinePath stays empty -> talk to the local machine
		sLocalPath		=	sObjectPath;
	}

	// Find the registry root
	nMainKey				=	sLocalPath.Find (L"\\");

	if (nMainKey == -1)
	{
		// One of the root keys was specified
		sMainKey			=	sLocalPath;
	}
	else
	{
		// A subkey was specified
		sMainKey			=	sLocalPath.Left (nMainKey);
		sSubkeyPath		=	sLocalPath.Right (sLocalPath.GetLength () - nMainKey - 1);
	}

	//
	// Make registry paths easier to use: accept formats: machine, hklm, hkey_local_machine
	//
	if (sMainKey.CompareNoCase (L"hklm") == 0 ||
		sMainKey.CompareNoCase (L"hkey_local_machine") == 0 ||
		sMainKey.CompareNoCase (L"machine") == 0)
	{
		sMainKey			=	L"machine";
		hMainKey			=	HKEY_LOCAL_MACHINE;
	}
	else if (sMainKey.CompareNoCase (L"hku") == 0 ||
				sMainKey.CompareNoCase (L"hkey_users") == 0 ||
				sMainKey.CompareNoCase (L"users") == 0)
	{
		sMainKey			=	L"users";
		hMainKey			=	HKEY_USERS;
	}
	else if (sMainKey.CompareNoCase (L"hkcr") == 0 ||
				sMainKey.CompareNoCase (L"hkey_classes_root") == 0 ||
				sMainKey.CompareNoCase (L"classes_root") == 0)
	{
		sMainKey			=	L"classes_root";
		hMainKey			=	HKEY_CLASSES_ROOT;
	}
	else if (sMainKey.CompareNoCase (L"hkcu") == 0 ||
				sMainKey.CompareNoCase (L"hkey_current_user") == 0 ||
				sMainKey.CompareNoCase (L"current_user") == 0)
	{
		sMainKey			=	L"current_user";
		hMainKey			=	HKEY_CURRENT_USER;
	}
	else
	{
		return RTN_ERR_REG_PATH;
	}

	// Build a corrected object path
	if (! sMachinePath.IsEmpty ())
	{
		sObjectPath		=	sMachinePath + L"\\";
	}
	else
	{
		sObjectPath.Empty ();
	}

	sObjectPath			+=	sMainKey;

	if (! sSubkeyPath.IsEmpty ())
	{
		sObjectPath		+=	L"\\" + sSubkeyPath;
	}

	// HKEY_CLASSES_ROOT and HKEY_CURRENT_USER cannot be used with RegConnectRegistry
	if (! sMachinePath.IsEmpty ())
	{
		if (hMainKey != HKEY_LOCAL_MACHINE && hMainKey != HKEY_USERS)
		{
			return RTN_ERR_REG_PATH;
		}
	}

	// The path is now "clean". Are we done?
	if (fFixPathOnly)
	{
		return RTN_OK;
	}

	// Now we can connect to the remote registry
	if (! sMachinePath.IsEmpty ())
	{
		m_nAPIError		=	RegConnectRegistry (sMachinePath, hMainKey, &hRemoteKey);
		if (m_nAPIError != ERROR_SUCCESS)
		{
			return RTN_ERR_REG_CONNECT;
		}

		// From now on, we only use the main key, but we'll keep the remote variable so we can correctly close the key later
		hMainKey			=	hRemoteKey;
	}

	//
	// Open the key specified either on the remote or the local machine
	//

	// Open the key using regular methods
	m_nAPIError			=	RegOpenKeyEx (hMainKey,  sSubkeyPath, 0, samDesired, &hOpenStd);

	if ((hOpenStd && m_nAPIError == ERROR_SUCCESS) || m_nAPIError == ERROR_ACCESS_DENIED)
	{
		// We now either know or can safely guess (access denied) that the key exists. Let's try some black magic and open it like a backup program

		DWORD	nNewCreated	=	0;
		DWORD nErrTmp		=	0;

		nErrTmp				=	RegCreateKeyEx (hMainKey,  sSubkeyPath, 0, NULL, REG_OPTION_BACKUP_RESTORE, samDesired,
														NULL, &hOpenBckp, &nNewCreated);

		// Assert we did not unintentionally create a new key
		if (nNewCreated == REG_CREATED_NEW_KEY)
		{
			LogMessage (Error, L"Unintentionally the following registry key was created: <" + sObjectPath + L">.");

			return RTN_ERR_INTERNAL;
		}

		// Check which opened key to use (standard or with privileges)
		if (hOpenBckp && nErrTmp == ERROR_SUCCESS)
		{
			hSubKey		=	hOpenBckp;

			// The standard key is not needed
			RegCloseKey (hOpenStd);

			// Reset the API error
			m_nAPIError	=	ERROR_SUCCESS;
		}
		else
		{
			hSubKey		=	hOpenStd;
		}
	}

	// The remote key is no longer needed - we have a handle to the subkey
	if (hRemoteKey)	{RegCloseKey (hRemoteKey);	hRemoteKey	= NULL;}

	if (m_nAPIError != ERROR_SUCCESS)
	{
		return RTN_ERR_REG_OPEN;
	}

	return RTN_OK;
}


//
// RecurseRegistry: Recurse the registry and call the function for every key
//
DWORD CSetACL::RecurseRegistry (CString  sObjectPath, DWORD (CSetACL::*funcProcess) (CString sObjectPath, bool fIsContainer))
{
	CString							sFindKey;
	CArray<CString, CString>	asSubKeys;
	DWORD								nError			=	RTN_OK;
	DWORD								nBuffer			=	512;
	HKEY								hSubKey			=	NULL;
	PFILETIME						pFileTime		=	NULL;

	// Process the current key
	nError	=	(this->*funcProcess) (sObjectPath, true);

	// Stop here on error or if no recursion is desired
	if (nError != RTN_OK || m_nRecursionType & RECURSE_NO)
	{
		return nError;
	}

	// Open the key
	nError	=	RegKeyFixPathAndOpen (sObjectPath, hSubKey, false, KEY_READ);
	if (nError != RTN_OK)
	{
		return nError;
	}

	// We have to save the subkeys since they change when setting permissions!
	m_nAPIError			=	RegEnumKeyEx (hSubKey, 0, sFindKey.GetBuffer (512), &nBuffer, NULL, NULL, NULL, pFileTime);
	sFindKey.ReleaseBuffer ();
	for (int i = 1; m_nAPIError == ERROR_SUCCESS; i++)
	{
		asSubKeys.Add (sFindKey);

		nBuffer			=	512;
		m_nAPIError		=	RegEnumKeyEx (hSubKey, i, sFindKey.GetBuffer (512), &nBuffer, NULL, NULL, NULL, pFileTime);
		sFindKey.ReleaseBuffer ();
	}

	if (m_nAPIError != ERROR_NO_MORE_ITEMS)
	{
		if (hSubKey) RegCloseKey (hSubKey);

		return RTN_ERR_REG_ENUM;
	}
	else
	{
		m_nAPIError		=	ERROR_SUCCESS;
	}

	if (hSubKey) RegCloseKey (hSubKey);

	// Process each subkey (start recursion)
	for (int i = 0; i < asSubKeys.GetSize (); i++)
	{
		RecurseRegistry (sObjectPath + L"\\" + asSubKeys.GetAt (i), funcProcess);
	}

	return nError;
}


//
// Write2SD: Set/Add the ACEs, owner and primary group specified to the ACLs
//
DWORD CSetACL::Write2SD (CString sObjectPath, bool fIsContainer)
{
	EXPLICIT_ACCESS					*eaDACL					=	NULL;
	EXPLICIT_ACCESS					*eaSACL					=	NULL;
	PACL									paclDACLNew				=	NULL;
	PACL									paclSACLNew				=	NULL;
	DWORD									nError					=	RTN_OK;
	POSITION								pos;
	DWORD									nDACLACEs				=	0;
	DWORD									nSACLACEs				=	0;
	SECURITY_INFORMATION				siSecInfo				=	0;
	SECURITY_INFORMATION				siSecInfoParent		=	0;
	bool									fDelInhACEsFromDACL	=	false;
	bool									fDelInhACEsFromSACL	=	false;
	CSD									csdSD(this);
	CSD									csdSDParent(this);


	//
	// If sub-object-only processing is enabled: check if this is a sub-object or not
	//
	if (m_fProcessSubObjectsOnly)
	{
		// Remove a trailing backslash to get a consistent state
		CString	sPath1	=	sObjectPath;
		CString	sPath2	=	m_sObjectPath;

		sPath1.TrimRight (L"\\");
		sPath2.TrimRight (L"\\");

		if (sPath1.CompareNoCase (sPath2) == 0)
		{
			// This is the object itself -> do nothing
			return RTN_OK;
		}
	}

	// Check if the current path is on the filter list. If yes -> do nothing
	if (CheckFilterList (sObjectPath))
	{
		// Notify caller of omission
		LogMessage (Information, L"Omitting ACL of: <" + sObjectPath + L"> because a filter keyword matched.");

		return RTN_OK;
	}

	// Notify caller of progress
	LogMessage (Information, L"Processing ACL of: <" + sObjectPath + L">");

	//
	// Check which elements of the SD we need to process, depending on the action(s) desired
	//
	if (m_nDACLEntries && (m_nAction & ACTN_ADDACE))
	{
		siSecInfo		|=	DACL_SECURITY_INFORMATION;
	}
	if (m_nSACLEntries && (m_nAction & ACTN_ADDACE))
	{
		siSecInfo		|=	SACL_SECURITY_INFORMATION;
	}
	if (m_pOwner->m_psidTrustee && (m_nAction & ACTN_SETOWNER))
	{
		siSecInfo		|=	OWNER_SECURITY_INFORMATION;
	}
	if (m_pPrimaryGroup->m_psidTrustee && (m_nAction & ACTN_SETGROUP))
	{
		siSecInfo		|=	GROUP_SECURITY_INFORMATION;
	}

	if (m_nAction & ACTN_SETINHFROMPAR)
	{
		if (m_nDACLProtected	== INHPARYES)
		{
			siSecInfo	|=	DACL_SECURITY_INFORMATION | UNPROTECTED_DACL_SECURITY_INFORMATION;
		}
		else if (m_nDACLProtected	== INHPARCOPY)
		{
			siSecInfo	|=	DACL_SECURITY_INFORMATION | PROTECTED_DACL_SECURITY_INFORMATION;
		}
		else if (m_nDACLProtected	== INHPARNOCOPY)
		{
			// Remove inherited ACEs
			fDelInhACEsFromDACL	=	true;

			siSecInfo			|=	DACL_SECURITY_INFORMATION | PROTECTED_DACL_SECURITY_INFORMATION;
			siSecInfoParent	|= DACL_SECURITY_INFORMATION;
		}

		if (m_nSACLProtected	== INHPARYES)
		{
			siSecInfo	|=	SACL_SECURITY_INFORMATION | UNPROTECTED_SACL_SECURITY_INFORMATION;
		}
		else if (m_nSACLProtected	== INHPARCOPY)
		{
			siSecInfo	|=	SACL_SECURITY_INFORMATION | PROTECTED_SACL_SECURITY_INFORMATION;
		}
		else if (m_nSACLProtected	== INHPARNOCOPY)
		{
			// Remove inherited ACEs
			fDelInhACEsFromSACL	=	true;

			siSecInfo			|=	SACL_SECURITY_INFORMATION | PROTECTED_SACL_SECURITY_INFORMATION;
			siSecInfoParent	|= SACL_SECURITY_INFORMATION;
		}
	}

	if (m_nAction & ACTN_CLEARDACL)
	{
		siSecInfo	|=	DACL_SECURITY_INFORMATION;
	}
	if (m_nAction & ACTN_CLEARSACL)
	{
		siSecInfo	|=	SACL_SECURITY_INFORMATION;
	}

	if ((m_nAction & ACTN_TRUSTEE) && m_fTrusteesProcessDACL)
	{
		siSecInfo	|=	DACL_SECURITY_INFORMATION;
	}
	if ((m_nAction & ACTN_TRUSTEE) && m_fTrusteesProcessSACL)
	{
		siSecInfo	|=	SACL_SECURITY_INFORMATION;
	}

	if ((m_nAction & ACTN_DOMAIN) && m_fDomainsProcessDACL)
	{
		siSecInfo	|=	DACL_SECURITY_INFORMATION;
	}
	if ((m_nAction & ACTN_DOMAIN) && m_fDomainsProcessSACL)
	{
		siSecInfo	|=	SACL_SECURITY_INFORMATION;
	}


	//
	// Get the SD
	//
	nError				=	csdSD.GetSD (sObjectPath, m_nObjectType, siSecInfo);
	m_nAPIError			=	csdSD.m_nAPIError;

	if (nError != RTN_OK)
	{
		if (m_fIgnoreErrors)
		{
			LogMessage (Error, L"Reading the SD from <" + sObjectPath + L"> failed with: " + GetLastErrorMessage (m_nAPIError));
			nError		=	RTN_ERR_IGNORED;
		}
		else
		{
			LogMessage (Error, L"Reading the SD from <" + sObjectPath + L"> failed with: " + GetLastErrorMessage (m_nAPIError));
		}

		goto CleanUp;
	}

	//
	// Get the parent object's SD so we can determine pseudo-inheritance
	//
	if (siSecInfoParent)
	{
		// Get the parent object
		CString sParentPath;
		if (! GetParentObject (sObjectPath, m_nObjectType, sParentPath))
		{
			goto DoneGettingParent;
		}

		// Get the parent's SD
		nError			=	csdSDParent.GetSD (sParentPath, m_nObjectType, siSecInfoParent);
		m_nAPIError		=	csdSDParent.m_nAPIError;

		if (nError != RTN_OK || csdSDParent.m_psdSD == NULL)
		{
			// Clear the SD object
			csdSDParent.Reset ();
			// Reset error codes
			nError		=	RTN_OK;
			m_nAPIError	=	NO_ERROR;
			goto DoneGettingParent;
		}
	}

DoneGettingParent:

	//
	// Remove inherited ACEs if necessary
	//
	if (fDelInhACEsFromDACL)
	{
		// Regular inherited ACEs
		nError	=	csdSD.DeleteACEsByHeaderFlags (ACL_DACL, INHERITED_ACE, true);
		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
		// Pseudo-inherited ACEs
		DWORD nRemainingACEsInDACL = 0;
		nError	= csdSD.DeletePseudoInheritedACEs (ACL_DACL, fIsContainer, csdSDParent.m_paclDACL, nRemainingACEsInDACL);
		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
	}
	if (fDelInhACEsFromSACL)
	{
		// Regular inherited ACEs
		nError	=	csdSD.DeleteACEsByHeaderFlags (ACL_SACL, INHERITED_ACE, true);
		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
		// Pseudo-inherited ACEs
		DWORD nRemainingACEsInSACL = 0;
		nError	= csdSD.DeletePseudoInheritedACEs (ACL_SACL, fIsContainer, csdSDParent.m_paclSACL, nRemainingACEsInSACL);
		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
	}

	//
	// Remove noninherited ACEs if necessary
	//
	if (m_nAction & ACTN_CLEARDACL)
	{
		nError	=	csdSD.DeleteACEsByHeaderFlags (ACL_DACL, INHERITED_ACE, false);
		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
	}
	if (m_nAction & ACTN_CLEARSACL)
	{
		nError	=	csdSD.DeleteACEsByHeaderFlags (ACL_SACL, INHERITED_ACE, false);
		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
	}

	//
	// Process specified trustees if necessary
	//
	if ((m_nAction & ACTN_TRUSTEE) && m_fTrusteesProcessDACL)
	{
		nError	=	csdSD.ProcessACEsOfGivenTrustees (ACL_DACL);

		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
	}

	if ((m_nAction & ACTN_TRUSTEE) && m_fTrusteesProcessSACL)
	{
		nError	=	csdSD.ProcessACEsOfGivenTrustees (ACL_SACL);

		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
	}

	//
	// Process specified domains if necessary
	//
	if ((m_nAction & ACTN_DOMAIN) && m_fDomainsProcessDACL)
	{
		nError	=	csdSD.ProcessACEsOfGivenDomains (ACL_DACL);

		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
	}

	if ((m_nAction & ACTN_DOMAIN) && m_fDomainsProcessSACL)
	{
		nError	=	csdSD.ProcessACEsOfGivenDomains (ACL_SACL);

		if (nError != RTN_OK)
		{
			goto CleanUp;
		}
	}

	//
	//	Process action AddACE
	//
	if (m_nAction & ACTN_ADDACE)
	{
		// Allocate memory for two lists of EXPLICIT_ACCESS structures
		if (m_nDACLEntries)
		{
			eaDACL			=	new EXPLICIT_ACCESS[m_nDACLEntries];
			
			if (eaDACL == NULL)
			{
				nError	= RTN_ERR_OUT_OF_MEMORY;
				goto CleanUp;
			}
		}
		if (m_nSACLEntries)
		{
			eaSACL			=	new EXPLICIT_ACCESS[m_nSACLEntries];
			
			if (eaSACL == NULL)
			{
				nError	= RTN_ERR_OUT_OF_MEMORY;
				goto CleanUp;
			}
		}

		//
		// We have got pointers to the DACL and SACL. Now loop through the ACEs specified and fill EXPLCIT_ACCESS structures describing them.
		//
		pos	=	m_lstACEs.GetHeadPosition ();

		while (pos)
		{
			CACE*	pACE	=	m_lstACEs.GetNext (pos);

			// Fill an EXPLICIT_ACCESS structure describing this ACE
			if (pACE->m_nACLType == ACL_DACL && m_nDACLEntries)
			{
				eaDACL[nDACLACEs].grfAccessMode			=	pACE->m_nAccessMode;
				eaDACL[nDACLACEs].grfAccessPermissions	=	pACE->m_nAccessMask;
				eaDACL[nDACLACEs].grfInheritance			=	pACE->m_nInheritance;
				eaDACL[nDACLACEs].Trustee.TrusteeForm	=	TRUSTEE_IS_SID;
				eaDACL[nDACLACEs].Trustee.ptstrName		=	(TCHAR *) pACE->m_pTrustee->m_psidTrustee;

				nDACLACEs++;
			}
			if (pACE->m_nACLType == ACL_SACL && m_nSACLEntries)
			{
				eaSACL[nSACLACEs].grfAccessMode			=	pACE->m_nAccessMode;
				eaSACL[nSACLACEs].grfAccessPermissions	=	pACE->m_nAccessMask;
				eaSACL[nSACLACEs].grfInheritance			=	pACE->m_nInheritance;
				eaSACL[nSACLACEs].Trustee.TrusteeForm	=	TRUSTEE_IS_SID;
				eaSACL[nSACLACEs].Trustee.ptstrName		=	(TCHAR *) pACE->m_pTrustee->m_psidTrustee;

				nSACLACEs++;
			}
		}	// if (m_nAction == ACTN_ADDACE)

		// ASSERT count is correct
		if (nDACLACEs != m_nDACLEntries || nSACLACEs != m_nSACLEntries)
		{
			nError	=	RTN_ERR_INTERNAL;

			goto CleanUp;
		}
	}

	//
	// Merge the existing (maybe modified) and (maybe empty) new ACL
	//
	m_nAPIError		=	SetEntriesInAcl (m_nDACLEntries, eaDACL, csdSD.m_paclDACL, &paclDACLNew);
	if (m_nAPIError != ERROR_SUCCESS)
	{
		if (m_fIgnoreErrors)
		{
			LogMessage (Error, L"SetEntriesInAcl for DACL of <" + sObjectPath + L"> failed with: " + GetLastErrorMessage (m_nAPIError));
			nError		=	RTN_ERR_IGNORED;
		}
		else
		{
			nError		=	RTN_ERR_SETENTRIESINACL;
		}
		goto CleanUp;
	}

	m_nAPIError		=	SetEntriesInAcl (m_nSACLEntries, eaSACL, csdSD.m_paclSACL, &paclSACLNew);
	if (m_nAPIError != ERROR_SUCCESS)
	{
		if (m_fIgnoreErrors)
		{
			LogMessage (Error, L"SetEntriesInAcl for SACL of <" + sObjectPath + L"> failed with: " + GetLastErrorMessage (m_nAPIError));
			nError		=	RTN_ERR_IGNORED;
		}
		else
		{
			nError		=	RTN_ERR_SETENTRIESINACL;
		}
		goto CleanUp;
	}

	//
	// Set the new ACLs, owner and primary group on the object
	//
	// Is there anything to do?
	if(m_nObjectType == SE_WMIGUID_OBJECT) {
		nError		=	csdSD.SetWmiSD (sObjectPath, m_nObjectType, paclDACLNew, paclSACLNew, m_pOwner->m_psidTrustee, m_pPrimaryGroup->m_psidTrustee);
		m_nAPIError	=	csdSD.m_nAPIError;
	} 
	else if (siSecInfo)
	{
		nError		=	csdSD.SetSD (sObjectPath, m_nObjectType, siSecInfo, paclDACLNew, paclSACLNew, m_pOwner->m_psidTrustee, m_pPrimaryGroup->m_psidTrustee, false, fIsContainer);
		m_nAPIError	=	csdSD.m_nAPIError;
	}

	if (nError != RTN_OK)
	{
		if (m_fIgnoreErrors)
		{
			LogMessage (Error, L"Writing SD to <" + sObjectPath + L"> failed with: " + GetLastErrorMessage (m_nAPIError));
			nError	=	RTN_ERR_IGNORED;
		}

		goto CleanUp;
	}

CleanUp:

	// If errors are to be ignored (param "-ignoreerr" set), then do NOT return an error
	if (nError == RTN_ERR_IGNORED)
	{
		m_nAPIError	=	ERROR_SUCCESS;
		nError		=	RTN_OK;
	}

	delete [] eaDACL; eaDACL = NULL;
	delete [] eaSACL; eaSACL = NULL;
	LocalFree (paclDACLNew); paclDACLNew = NULL;
	LocalFree (paclSACLNew); paclSACLNew = NULL;

	return nError;
}


//
// ListSD: List the contents of a SD in text format
//
DWORD CSetACL::ListSD (CString sObjectPath, bool fIsContainer)
{
	DWORD									nError					=	RTN_OK;
	SECURITY_INFORMATION				siSecInfo				=	0;
	SECURITY_INFORMATION				siSecInfoParent		=	0;
	LPTSTR								lptstrSDDL				=	NULL;
	CString								sSDDL;
	CString								sLineOut;
	CString								sObjectType;
	CString								sSD;
	CString								sDACLControl;
	CString								sSACLControl;
	CString								sParentPath;
	CSD									csdSD (this);
	CSD									csdSDParent (this);
	DWORD									nACEs						=	0;
	DWORD									nRemainingACEsInDACL	=	0;
	DWORD									nRemainingACEsInSACL	=	0;
	DWORD									nSDRevision				=	0;
	SECURITY_DESCRIPTOR_CONTROL	sdControl;
	bool									fDACLPseudoProtected	=	false;
	bool									fSACLPseudoProtected	=	false;

	// Reset the last result buffer
	m_sListOutput.Empty();

	// Convert the object type to a string for storage
	sObjectType.Format (L"%d", m_nObjectType);

	// Were the options set?
	if (! m_nListWhat)
	{
		return RTN_ERR_LIST_OPTIONS;
	}

	// Check if the current path is on the filter list. If yes -> do nothing
	if (CheckFilterList (sObjectPath))
	{
		// Notify caller of the omission
		LogMessage (Information, L"Omitting ACL of: <" + sObjectPath + L"> because a filter keyword matched.");

		return RTN_OK;
	}

	// Check which elements of the SD we need to process
	// We always query the owner in order to be able to determine if CREATOR_OWNER was inherited correctly (pseudo-protection check)
	if (m_nListWhat & ACL_DACL)
	{
		siSecInfo			|=	DACL_SECURITY_INFORMATION;
		siSecInfoParent	|=	DACL_SECURITY_INFORMATION;
	}
	if (m_nListWhat & ACL_SACL)
	{
		siSecInfo			|=	SACL_SECURITY_INFORMATION;
		siSecInfoParent	|=	SACL_SECURITY_INFORMATION;
	}
	if (m_nListWhat & SD_GROUP)
	{
		siSecInfo	|=	GROUP_SECURITY_INFORMATION;
	}
	if (m_nListWhat & SD_OWNER)
	{
		siSecInfo	|=	OWNER_SECURITY_INFORMATION;
	}

	//
	// Get the SD and the SD's control information
	//
	// Get the SD
	nError			=	csdSD.GetSD (sObjectPath, m_nObjectType, siSecInfo | OWNER_SECURITY_INFORMATION);
	m_nAPIError		=	csdSD.m_nAPIError;

	if (nError != RTN_OK)
	{
		goto CleanUp;
	}
	else if (csdSD.m_psdSD == NULL)
	{
		LogMessage (Information, L"The object <" + m_sObjectPath + L"> has a NULL security descriptor (granting full control to everyone) and is being ignored.");
		nError	=	RTN_ERR_IGNORED;

		goto CleanUp;
	}

	// Get the SD's control information, which contains the protection flags
	if (! GetSecurityDescriptorControl (csdSD.m_psdSD, &sdControl, &nSDRevision))
	{
		nError			=	RTN_ERR_GET_SD_CONTROL;
		m_nAPIError		=	GetLastError ();

		goto CleanUp;
	}

	// If DACL or SACL are protected, they need not be fetched from the parent for pseudo-inheritance checking
	if (sdControl & SE_DACL_PROTECTED)
	{
		siSecInfoParent	&= ~DACL_SECURITY_INFORMATION;
	}
	// ... and SACL
	if (sdControl & SE_SACL_PROTECTED)
	{
		siSecInfoParent	&= ~SACL_SECURITY_INFORMATION;
	}

	//
	// Get the parent object's SD so we can determine pseudo-inheritance
	//
	if (siSecInfoParent)
	{
		// Get the parent object
		if (! GetParentObject (sObjectPath, m_nObjectType, sParentPath))
		{
			goto DoneGettingParent;
		}

		// Get the parent's SD
		nError			=	csdSDParent.GetSD (sParentPath, m_nObjectType, siSecInfoParent);
		m_nAPIError		=	csdSDParent.m_nAPIError;

		if (nError != RTN_OK || csdSDParent.m_psdSD == NULL)
		{
			// Clear the SD object
			csdSDParent.Reset ();
			// Reset error codes
			nError		=	RTN_OK;
			m_nAPIError	=	NO_ERROR;
			goto DoneGettingParent;
		}
	}

DoneGettingParent:

	// Build the SD control strings: DACL
	if (sdControl & SE_DACL_PRESENT)
	{
		if (sdControl & SE_DACL_PROTECTED)
		{
			sDACLControl	=	L"(protected";
		}
		else
		{
			// Check if the DACL is "pseudo-protected"
			if (IsACLPseudoProtected (csdSD.m_paclDACL, fIsContainer, csdSD.m_psidOwner, csdSDParent.m_paclDACL))
			{
				sDACLControl			= L"(pseudo_protected";
				fDACLPseudoProtected	= true;

				// Delete the parents DACL so everyone below thinks the DACL is protected
				csdSDParent.DeleteBufDACL();
				csdSDParent.m_paclDACL	= NULL;
			}
			else
			{
				sDACLControl	=	L"(not_protected";
			}
		}
		if (sdControl & SE_DACL_AUTO_INHERITED)
		{
			sDACLControl	+=	L"+auto_inherited)";
		}
		else
		{
			sDACLControl	+=	L")";
		}
	}

	// Build the SD control strings: SACL
	if (sdControl & SE_SACL_PRESENT)
	{
		if (sdControl & SE_SACL_PROTECTED)
		{
			sSACLControl	=	L"(protected";
		}
		else
		{
			if (IsACLPseudoProtected (csdSD.m_paclSACL, fIsContainer, csdSD.m_psidOwner, csdSDParent.m_paclSACL))
			{
				sSACLControl			= L"(pseudo_protected";
				fSACLPseudoProtected	= true;

				// Delete the parents SACL so everyone below thinks the SACL is protected
				csdSDParent.DeleteBufSACL();
				csdSDParent.m_paclSACL	= NULL;
			}
			else
			{
				sSACLControl	=	L"(not_protected";
			}
		}
		if (sdControl & SE_SACL_AUTO_INHERITED)
		{
			sSACLControl	+=	L"+auto_inherited)";
		}
		else
		{
			sSACLControl	+=	L")";
		}
	}

	//
	// SDDL listing
	//
	if (m_nListFormat == LIST_SDDL)
	{
		if (! m_fListInherited)
		{
			// Remove any (pseudo-) inherited ACEs

			if (siSecInfo & DACL_SECURITY_INFORMATION)
			{
				nError	= csdSD.DeleteACEsByHeaderFlags (ACL_DACL, INHERITED_ACE, true);
				if (nError != RTN_OK)
				{
					goto CleanUp;
				}
				nError	= csdSD.DeletePseudoInheritedACEs (ACL_DACL, fIsContainer, csdSDParent.m_paclDACL, nRemainingACEsInDACL);
				if (nError != RTN_OK)
				{
					goto CleanUp;
				}
			}
			if (siSecInfo & SACL_SECURITY_INFORMATION)
			{
				nError	= csdSD.DeleteACEsByHeaderFlags (ACL_SACL, INHERITED_ACE, true);
				if (nError != RTN_OK)
				{
					goto CleanUp;
				}
				nError	 = csdSD.DeletePseudoInheritedACEs (ACL_SACL, fIsContainer, csdSDParent.m_paclSACL, nRemainingACEsInSACL);
				if (nError != RTN_OK)
				{
					goto CleanUp;
				}
			}
		}

		// Convert the SD into the SDDL format
		if (! ConvertSecurityDescriptorToStringSecurityDescriptor (csdSD.m_psdSD, SDDL_REVISION_1, siSecInfo, &lptstrSDDL, NULL))
		{
			nError		=	RTN_ERR_CONVERT_SD;
			m_nAPIError	=	GetLastError ();

			goto CleanUp;
		}

		//
		// Format and log the result
		//
		sSDDL	= lptstrSDDL;
		if (! sSDDL.IsEmpty ())
		{
			sSDDL	=	L"\"" + sSDDL + L"\"";
		}
		if (sObjectPath.Left (1) != TEXT ("\""))
		{
			sLineOut		=	L"\"" + sObjectPath + L"\"," + sObjectType + L"," + sSDDL;
		}
		else
		{
			sLineOut		=	sObjectPath + L"," + sObjectType + L"," + sSDDL;
		}
	}
	else if (m_nListFormat == LIST_CSV)
	{
		//
		// CSV listing
		//
		if (sObjectPath.Left (1) != TEXT ("\""))
		{
			sLineOut		=	TEXT ("\"") + sObjectPath + TEXT ("\",") + sObjectType + L",";
		}
		else
		{
			sLineOut		=	sObjectPath + L"," + sObjectType + L",";
		}

		// Process DACL if necessary
		if (m_nListWhat & ACL_DACL)
		{
			if (csdSD.m_paclDACL)
			{
				CString sDACL			=	ListACL (csdSD.m_paclDACL, fIsContainer, csdSDParent.m_paclDACL, nACEs);
				if (m_nAPIError)
				{
					sSD			=	L"DACL" + sDACLControl + L":[error:" + GetLastErrorMessage (m_nAPIError) + L"]";
					if (m_fIgnoreErrors)
					{
						m_nAPIError	=	ERROR_SUCCESS;
						nError		=	RTN_OK;
					}
					else
					{
						nError		=	RTN_ERR_LIST_ACL;
						goto CleanUp;
					}
				}
				else
				{
					if (nACEs == 0)
					{
						sSD			=	L"DACL" + sDACLControl + L":[empty]";
					}
					else if (! sDACL.IsEmpty())
					{
						sSD			=	L"DACL" + sDACLControl + L":" + sDACL;
					}
				}
			}
			else
			{
				sSD					=	L"DACL" + sDACLControl + L":[NULL]";
			}
		}

		// Process SACL if necessary
		if (m_nListWhat & ACL_SACL)
		{
			if (! sSD.IsEmpty())
				sSD	+=	L";";

			if (csdSD.m_paclSACL)
			{
				CString sSACL			=	ListACL (csdSD.m_paclSACL, fIsContainer, csdSDParent.m_paclSACL, nACEs);
				if (m_nAPIError)
				{
					sSD			+=	L"SACL" + sSACLControl + L":[error:" + GetLastErrorMessage (m_nAPIError) + L"]";
					if (m_fIgnoreErrors)
					{
						m_nAPIError	=	ERROR_SUCCESS;
						nError		=	RTN_OK;
					}
					else
					{
						nError		=	RTN_ERR_LIST_ACL;
						goto CleanUp;
					}
				}
				else
				{
					if (nACEs == 0)
					{
						sSD			+=	L"SACL" + sSACLControl + L":[empty]";
					}
					else if (! sSACL.IsEmpty())
					{
						sSD			+=	L"SACL" + sSACLControl + L":" + sSACL;
					}
				}
			}
			else
			{
				sSD					+=	L"SACL" + sSACLControl + L":[NULL]";
			}
		}

		// Process owner if necessary
		if (m_nListWhat & SD_OWNER)
		{
			if (! sSD.IsEmpty())
				sSD	+=	L";";

			if (csdSD.m_psidOwner)
			{
				CString sOwner	=	GetTrusteeFromSID (csdSD.m_psidOwner, NULL);
				if (m_nAPIError)
				{
					sSD			+=	L"Owner:[error:" + GetLastErrorMessage (m_nAPIError) + L"]";
					if (m_fIgnoreErrors)
					{
						m_nAPIError	=	ERROR_SUCCESS;
						nError		=	RTN_OK;
					}
					else
					{
						nError		=	RTN_ERR_LIST_ACL;
						goto CleanUp;
					}
				}
				else
				{
					if (sOwner.IsEmpty ())
					{
						sSD			+=	L"Owner:[empty]";
					}
					else
					{
						sSD			+=	L"Owner:" + sOwner;
					}
				}
			}
			else
			{
				sSD					+=	L"Owner:[NULL]";
			}
		}

		// Process primary group if necessary
		if (m_nListWhat & SD_GROUP)
		{
			if (! sSD.IsEmpty())
				sSD	+=	L";";

			if (csdSD.m_psidGroup)
			{
				CString sGroup	=	GetTrusteeFromSID (csdSD.m_psidGroup, NULL);
				if (m_nAPIError)
				{
					sSD			+=	L"Group:[error:" + GetLastErrorMessage (m_nAPIError) + L"]";
					if (m_fIgnoreErrors)
					{
						m_nAPIError	=	ERROR_SUCCESS;
						nError		=	RTN_OK;
					}
					else
					{
						nError		=	RTN_ERR_LIST_ACL;
						goto CleanUp;
					}
				}
				else
				{
					if (sGroup.IsEmpty ())
					{
						sSD			+=	L"Group:[empty]";
					}
					else
					{
						sSD			+=	L"Group:" + sGroup;
					}
				}
			}
			else
			{
				sSD					+=	L"Group:[NULL]";
			}
		}

		// Format the result
		if (! sSD.IsEmpty ())
		{
			sLineOut		+=	TEXT ("\"") + sSD + TEXT ("\"");
		}
		else
		{
			sLineOut		=	TEXT ("");
		}
	}
	else if (m_nListFormat == LIST_TAB)
	{
		//
		// Tab format listing
		//
		sLineOut		=	sObjectPath + L"\n";
		sSD			=	TEXT ("");

		// Process owner if necessary
		if (m_nListWhat & SD_OWNER)
		{
			if (csdSD.m_psidOwner)
			{
				CString sOwner	=	GetTrusteeFromSID (csdSD.m_psidOwner, NULL);
				if (m_nAPIError)
				{
					sSD			+=	L"\n   Owner: [error:" + GetLastErrorMessage (m_nAPIError) + L"]\n";
					if (m_fIgnoreErrors)
					{
						m_nAPIError	=	ERROR_SUCCESS;
						nError		=	RTN_OK;
					}
					else
					{
						nError		=	RTN_ERR_LIST_ACL;
						goto CleanUp;
					}
				}
				else
				{
					if (sOwner.IsEmpty ())
					{
						sSD			+=	L"\n   Owner: [empty]\n";
					}
					else
					{
						sSD			+=	L"\n   Owner: " + sOwner + L"\n";
					}
				}
			}
			else
			{
				sSD					+=	L"\n   Owner: [NULL]\n";
			}
		}

		// Process primary group if necessary
		if (m_nListWhat & SD_GROUP)
		{
			if (csdSD.m_psidGroup)
			{
				CString sGroup	=	GetTrusteeFromSID (csdSD.m_psidGroup, NULL);
				if (m_nAPIError)
				{
					sSD			+=	L"\n   Group: [error:" + GetLastErrorMessage (m_nAPIError) + L"]\n";
					if (m_fIgnoreErrors)
					{
						m_nAPIError	=	ERROR_SUCCESS;
						nError		=	RTN_OK;
					}
					else
					{
						nError		=	RTN_ERR_LIST_ACL;
						goto CleanUp;
					}
				}
				else
				{
					if (sGroup.IsEmpty ())
					{
						sSD			+=	L"\n   Group: [empty]\n";
					}
					else
					{
						sSD			+=	L"\n   Group: " + sGroup + "\n";
					}
				}
			}
			else
			{
				sSD					+=	L"\n   Group: [NULL]\n";
			}
		}

		// Process DACL if necessary
		if (m_nListWhat & ACL_DACL)
		{
			if (csdSD.m_paclDACL)
			{
				CString sDACL	=	ListACL (csdSD.m_paclDACL, fIsContainer, csdSDParent.m_paclDACL, nACEs);

				sDACL.Replace (L":", L"\n   ");
				sDACL.Replace (L",", L"   ");

				if (m_nAPIError)
				{
					sSD			+=	L"\n   DACL: [error:" + GetLastErrorMessage (m_nAPIError) + L"]\n";
					if (m_fIgnoreErrors)
					{
						m_nAPIError	=	ERROR_SUCCESS;
						nError		=	RTN_OK;
					}
					else
					{
						nError		=	RTN_ERR_LIST_ACL;
						goto CleanUp;
					}
				}
				else
				{
					if (nACEs == 0)
					{
						sSD			+=	L"\n   DACL" + sDACLControl + L":\n   [empty]\n";
					}
					else if (! sDACL.IsEmpty())
					{
						sSD			+=	L"\n   DACL" + sDACLControl + L":\n   " + sDACL + "\n";
					}
					else
					{
						sSD			+=	L"\n   DACL" + sDACLControl + L":\n   [no implicit permissions]\n";
					}
				}
			}
			else
			{
				sSD					+=	L"\n   DACL" + sDACLControl + L":\n   [NULL]\n";
			}
		}

		// Process SACL if necessary
		if (m_nListWhat & ACL_SACL)
		{
			if (csdSD.m_paclSACL)
			{
				CString sSACL	=	ListACL (csdSD.m_paclSACL, fIsContainer, csdSDParent.m_paclSACL, nACEs);

				sSACL.Replace (L":", L"\n   ");
				sSACL.Replace (L",", L"   ");

				if (m_nAPIError)
				{
					sSD			+=	L"\n   SACL: [error:" + GetLastErrorMessage (m_nAPIError) + L"]\n";
					if (m_fIgnoreErrors)
					{
						m_nAPIError	=	ERROR_SUCCESS;
						nError		=	RTN_OK;
					}
					else
					{
						nError		=	RTN_ERR_LIST_ACL;
						goto CleanUp;
					}
				}
				else
				{
					if (nACEs == 0)
					{
						sSD			+=	L"\n   SACL" + sSACLControl + L":\n   [empty]\n";
					}
					else if (! sSACL.IsEmpty())
					{
						sSD			+=	L"\n   SACL" + sSACLControl + L":\n   " + sSACL + "\n";
					}
					else
					{
						sSD			+=	L"\n   SACL" + sSACLControl + L":\n   [no implicit permissions]\n";
					}
				}
			}
			else
			{
				sSD					+=	L"\n   SACL" + sSACLControl + L":\n   [NULL]\n";
			}
		}

		// Format the result
		if (! sSD.IsEmpty ())
		{
			sLineOut		+=	sSD;
		}
		else
		{
			sLineOut		=	TEXT ("");
		}
	}

	// Store the result in the last result buffer
	m_sListOutput = sLineOut;

	// Save the result
	if (! sLineOut.IsEmpty ())
	{
		if (m_fhBackupRestoreFile)
		{
			if (_ftprintf (m_fhBackupRestoreFile, L"%s\r\n", sLineOut.GetString ()) < 0)
			{
				return RTN_ERR_WRITE_LOGFILE;
			}
		}

		// Print the result
		LogMessage (NoSeverity, sLineOut);
	}

CleanUp:

	// Log errors
	if (m_nAPIError)
	{
		LogMessage (Error, L"Parsing the SD of <" + sObjectPath + L"> failed with: " + GetLastErrorMessage (m_nAPIError));
	}

	// If errors are to be ignored (param "-ignoreerr" set), then do NOT return an error
	if (m_fIgnoreErrors || nError == RTN_ERR_IGNORED)
	{
		m_nAPIError	=	ERROR_SUCCESS;
		nError		=	RTN_OK;
	}

	if (lptstrSDDL)			{LocalFree (lptstrSDDL);	lptstrSDDL			= NULL;}

	return nError;
}


//
// GetTrusteeFromSID: Convert a binary SID into a trustee name
//
CString CSetACL::GetTrusteeFromSID (PSID psidSID, SID_NAME_USE* snuSidType)
{
	DWORD				nError			=	RTN_OK;
	DWORD				nAccountName	=	0;
	DWORD				nDomainName		=	0;
	LPTSTR			pcSID				=	NULL;
	CString			sSID;
	CString			sDomainName;
	CString			sAccountName;
	CString			sTrustee;
	CString			sTrusteeName;
	SID_NAME_USE	snuSidTypeTemp	=	SidTypeUnknown;

	// Convert binary to string SID
	if (! ConvertSidToStringSid (psidSID, &pcSID))
	{
		m_nAPIError	=	GetLastError ();
		return sTrustee;
	}
	sSID	=	pcSID;

	if (m_nListNameSID & LIST_NAME)
	{
		// Try to find the SID in the lookup cache
		if (!m_SIDLookupCache.Lookup(sSID, sTrusteeName))
		{
			// The name is not yet in the cache --> look it up

			// First get buffer sizes
			LookupAccountSid (m_sTargetSystemName.IsEmpty () ? NULL : (LPCTSTR) m_sTargetSystemName, psidSID, NULL, &nAccountName, NULL, &nDomainName, &snuSidTypeTemp);

			// Second, look up account and domain names
			BOOL	fOK	=	LookupAccountSid (m_sTargetSystemName.IsEmpty () ? NULL : (LPCTSTR) m_sTargetSystemName, psidSID, 
									sAccountName.GetBuffer (nAccountName), &nAccountName, sDomainName.GetBuffer (nDomainName), &nDomainName, &snuSidTypeTemp);
			sAccountName.ReleaseBuffer ();
			sDomainName.ReleaseBuffer ();

			if (! fOK)
			{
				nError	=	GetLastError ();
				if (nError != ERROR_NONE_MAPPED && nError != ERROR_TRUSTED_RELATIONSHIP_FAILURE)
				{
					// Error -> return
					m_nAPIError	=	nError;
					return sTrustee;
				}

				// The account name was not found: display the SID instead
			}

			// Build the trustee name string
			if (! sDomainName.IsEmpty ())
			{
				sTrusteeName	=	sDomainName + L"\\" + sAccountName;
			}
			else
			{
				sTrusteeName	=	sAccountName;
			}

			// Remove non-domain qualifiers (e.g. NT-AUTHORITY)?
			if (m_fListCleanOutput)
			{
				if (snuSidTypeTemp != SidTypeGroup && snuSidTypeTemp != SidTypeUser)
				{
					int nBackslash = sTrusteeName.Find('\\');
					if (nBackslash >= 0)
					{
						sTrusteeName = sTrusteeName.Right(sTrusteeName.GetLength() - nBackslash - 1);
					}
				}
			}

			// Add the name to the cache
			m_SIDLookupCache.SetAt(sSID, sTrusteeName);
		}
	}

	// Build the trustee string
	if (sTrusteeName.IsEmpty ())
	{
		// We could not or were not asked to find the name -> display the SID only
		sTrustee	=	sSID;
	}
	else if (m_nListNameSID == LIST_NAME_SID)
	{
		// Display the name and the SID
		sTrustee	=	sTrusteeName + L"=" + sSID;
	}
	else if (m_nListNameSID & LIST_NAME)
	{
		// Display the name only
		sTrustee	=	sTrusteeName;
	}

	if (pcSID)
	{
		LocalFree (pcSID);
	}

	if (snuSidType != NULL && m_nListNameSID & LIST_NAME)
		*snuSidType = snuSidTypeTemp;

	return sTrustee;
}


//
// ListACL: Return the contents of an ACL as a string
//
CString CSetACL::ListACL (PACL paclObject, bool fIsContainer, PACL paclParent, DWORD& nACEsObject)
{
	ACL_SIZE_INFORMATION				asiACLSize;
	ACCESS_ALLOWED_ACE*				paceACE						=	NULL;
	CString								sOut;
	CString								sTrustee;
	CString								sPermissions;
	CString								sFlags;
	CString								sACEType;
	bool									fACEIsPseudoInherited	= false;

	// If this is a NULL ACL, do nothing
	if (! paclObject)
	{
		return sOut;
	}

	// Get the number of entries in the object's ACL
	if (! GetAclInformation (paclObject, &asiACLSize, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		m_nAPIError	=	GetLastError ();
		return sOut;
	}

	// Store the number of ACEs in the object's ACL for access by the caller
	nACEsObject	=	asiACLSize.AceCount;

	// Loop through the ACEs in the object's ACL
	for (WORD i = 0; i < asiACLSize.AceCount; i++)
	{
		sTrustee.Empty ();
		sPermissions.Empty ();
		sFlags.Empty ();
		sACEType.Empty ();
		fACEIsPseudoInherited	= false;

		// Get the current ACE
		if (! GetAce (paclObject, i, (LPVOID*) &paceACE))
		{
			m_nAPIError	=	GetLastError ();
			return sOut;
		}

		if (! (paceACE->Header.AceFlags & INHERITED_ACE))
		{
			// If the ACE is not inherited, it might still be pseudo-inherited. Check that.
			fACEIsPseudoInherited	= IsACEPseudoInherited (paceACE, fIsContainer, paclParent);
		}

		// Omit inherited ACEs?
		if (! m_fListInherited)
		{
			if (paceACE->Header.AceFlags & INHERITED_ACE)
			{
				// ACE is marked as inherited
				continue;
			}

			if (fACEIsPseudoInherited)
			{
				// ACE is pseudo-inherited from parent without being marked as such (sad, but very common)
				continue;
			}
		}

		// Find the name corresponding to the SID in the ACE
		SID_NAME_USE snuSIDType;
		sTrustee				=	GetTrusteeFromSID ((PSID) &(paceACE->SidStart), &snuSIDType);
		if (m_nAPIError)
			return sOut;

		// Get the permissions
		sPermissions		=	GetPermissions (paceACE->Mask);

		// Get the ACE type
		sACEType				=	GetACEType (paceACE->Header.AceType);

		// Get the ACE flags
		sFlags				=	GetACEFlags (paceACE->Header.AceFlags, fACEIsPseudoInherited);

		sOut					+=	sTrustee + L"," + sPermissions + L"," + sACEType + L"," + sFlags + L":";
	}

	sOut.TrimRight (L":");

	return sOut;
}


//
// ProcessACEsOfGivenTrustees: Process (delete, replace, copy) all ACEs belonging to trustees specified
//
DWORD CSD::ProcessACEsOfGivenTrustees (DWORD nWhere)
{
	DWORD									nError			=	RTN_OK;
	ACL_SIZE_INFORMATION				asiACLSize;
	ACCESS_ALLOWED_ACE*				paceACE			=	NULL;
	DWORD									nACECount		=	0;
	PACL									paclACL			=	NULL;
	bool									fIsSACL			=	false;

	if (nWhere == ACL_DACL)
	{
		paclACL	=	m_paclDACL;
		fIsSACL	=	false;
	}
	else if (nWhere == ACL_SACL)
	{
		paclACL	=	m_paclSACL;
		fIsSACL	=	true;
	}
	else
	{
		return RTN_ERR_PARAMS;
	}

	// If this is a NULL ACL, do nothing
	if (! paclACL)
	{
		return nError;
	}

	// Get the number of entries in the ACL
	if (! GetAclInformation (paclACL, &asiACLSize, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		m_nAPIError	=	GetLastError ();
		return RTN_ERR_LOOP_ACL;
	}

	nACECount	=	asiACLSize.AceCount;

	// Loop through the ACEs
	for (DWORD i = 0; i < nACECount; i++)
	{
		// Get the current ACE
		if (! GetAce (paclACL, i, (LPVOID*) &paceACE))
		{
			m_nAPIError	=	GetLastError ();
			return RTN_ERR_LOOP_ACL;
		}

		// Do NOT change inherited ACEs
		if (paceACE->Header.AceFlags & INHERITED_ACE)
		{
			continue;
		}

		// Iterate through the Trustees specified
		POSITION	pos	=	m_setaclMain->m_lstTrustees.GetHeadPosition ();

		while (pos)
		{
			CTrustee*	pTrustee	=	m_setaclMain->m_lstTrustees.GetNext (pos);

			// Process this ACE only if the specified options match
			if (fIsSACL && !pTrustee->m_fSACL)
			{
				continue;
			}
			if (!fIsSACL && !pTrustee->m_fDACL)
			{
				continue;
			}

			//
			// Process the ACE if it belongs to a SID on the list.
			//
			if (EqualSid ((PSID) &(paceACE->SidStart), pTrustee->m_psidTrustee))
			{
				PACL	paclACLNew	=	NULL;

				if (pTrustee->m_nAction & ACTN_REMOVETRUSTEE)
				{
					if (! DeleteAce (paclACL, i))
					{
						m_nAPIError	=	GetLastError ();
						return RTN_ERR_DEL_ACE;
					}

					// The ACECount is now reduced by one!
					nACECount--;
					i--;
				}
				else if (pTrustee->m_nAction & ACTN_REPLACETRUSTEE)
				{
					// Replace the ACE into a new ACL
					paclACLNew		=	ACLReplaceACE (paclACL, i, pTrustee->m_oNewTrustee->m_psidTrustee);

					if (! paclACLNew)
					{
						m_nAPIError	=	GetLastError ();
						return RTN_ERR_COPY_ACL;
					}
				}
				else if (pTrustee->m_nAction & ACTN_COPYTRUSTEE)
				{
					// Copy the ACE into a new ACL
					paclACLNew		=	ACLCopyACE (paclACL, i, pTrustee->m_oNewTrustee->m_psidTrustee);

					if (! paclACLNew)
					{
						m_nAPIError	=	GetLastError ();
						return RTN_ERR_COPY_ACL;
					}

					// The ACE count has increased
					nACECount++;
				}

				if (paclACLNew)
				{
					// Free the old ACL
					if (nWhere == ACL_DACL)
					{
						DeleteBufDACL ();
					}
					else if (nWhere == ACL_SACL)
					{
						DeleteBufSACL ();
					}

					// Continue with the new ACL
					paclACL	=	paclACLNew;

					// Save the new ACL
					if (nWhere == ACL_DACL)
					{
						m_paclDACL			=	paclACL;
						m_fBufDACLAlloc	=	true;
					}
					else if (nWhere == ACL_SACL)
					{
						m_paclSACL			=	paclACL;
						m_fBufSACLAlloc	=	true;
					}
				}

				// Break out of the while loop and continue with next ACE
				break;
			}	// if (EqualSid)
		}		// while (pos)
	}			//	for (int i = 0; i < nACECount; i++)

	return nError;
}


//
// ProcessACEsOfGivenDomains: Process (delete, replace, copy) all ACEs belonging to domains specified
//
DWORD CSD::ProcessACEsOfGivenDomains (DWORD nWhere)
{
	DWORD									nError			=	RTN_OK;
	ACL_SIZE_INFORMATION				asiACLSize;
	ACCESS_ALLOWED_ACE*				paceACE			=	NULL;
	DWORD									nACECount		=	0;
	PACL									paclACL			=	NULL;
	bool									fIsSACL			=	false;

	if (nWhere == ACL_DACL)
	{
		paclACL	=	m_paclDACL;
		fIsSACL	=	false;
	}
	else if (nWhere == ACL_SACL)
	{
		paclACL	=	m_paclSACL;
		fIsSACL	=	true;
	}
	else
	{
		return RTN_ERR_PARAMS;
	}

	// If this is a NULL ACL, do nothing
	if (! paclACL)
	{
		return nError;
	}

	// Get the number of entries in the ACL
	if (! GetAclInformation (paclACL, &asiACLSize, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		m_nAPIError	=	GetLastError ();
		return RTN_ERR_LOOP_ACL;
	}

	nACECount	=	asiACLSize.AceCount;

	// Loop through the ACEs
	for (DWORD i = 0; i < nACECount; i++)
	{
		// Get the current ACE
		if (! GetAce (paclACL, i, (LPVOID*) &paceACE))
		{
			m_nAPIError	=	GetLastError ();
			return RTN_ERR_LOOP_ACL;
		}

		// Check for validity of SID
		if (! IsValidSid ((PSID) &(paceACE->SidStart)))
		{
			return RTN_ERR_LOOP_ACL;
		}

		// Do NOT change inherited ACEs
		if (paceACE->Header.AceFlags & INHERITED_ACE)
		{
			continue;
		}

		//
		// Get domain and trustee of the SID in the ACE
		//
		DWORD				nAccountName	=	0;
		DWORD				nDomainName		=	0;
		CString			sDomainName;
		CString			sAccountName;
		SID_NAME_USE	snuSidType;
		LPCTSTR			lpSystemName	=	m_setaclMain->m_sTargetSystemName.IsEmpty () ? NULL : (LPCTSTR) m_setaclMain->m_sTargetSystemName;

		// Try to look up the account name

		// First get buffer sizes
		LookupAccountSid (lpSystemName, (PSID) &(paceACE->SidStart), NULL, &nAccountName, NULL, &nDomainName, &snuSidType);

		// Second, look up account and domain names
      // TODO: Use the cache
		BOOL	fOK	=	LookupAccountSid (lpSystemName, (PSID) &(paceACE->SidStart), sAccountName.GetBuffer (nAccountName), &nAccountName, 
													sDomainName.GetBuffer (nDomainName), &nDomainName, &snuSidType);
		sAccountName.ReleaseBuffer ();
		sDomainName.ReleaseBuffer ();

		if (! fOK)
		{
			// Ignore SIDs that cannot be looked up
			continue;
		}

		if (snuSidType == SidTypeDeletedAccount || snuSidType == SidTypeInvalid || snuSidType == SidTypeUnknown)
		{
			// Ignore invalid or deleted SIDs
			continue;
		}

		if (sAccountName.IsEmpty () || sDomainName.IsEmpty ())
		{
			// Ignore dubious errors
			continue;
		}

		//
		// Iterate through the domains specified
		//
		POSITION	pos	=	m_setaclMain->m_lstDomains.GetHeadPosition ();

		while (pos)
		{
			CDomain*	pDomain	=	m_setaclMain->m_lstDomains.GetNext (pos);

			// Process this ACE only if the specified options match
			if (fIsSACL && !pDomain->m_fSACL)
			{
				continue;
			}
			if (!fIsSACL && !pDomain->m_fDACL)
			{
				continue;
			}

			//
			// Process the ACE if it belongs to a domain on the list.
			//
			if (sDomainName.CompareNoCase (pDomain->m_sDomain) == 0)
			{
				PACL	paclACLNew	=	NULL;

				if (pDomain->m_nAction & ACTN_REMOVETRUSTEE)
				{
					if (! DeleteAce (paclACL, i))
					{
						m_nAPIError	=	GetLastError ();
						return RTN_ERR_DEL_ACE;
					}

					// The ACECount is now reduced by one!
					nACECount--;
					i--;
				}
				else
				{
					// Make sure a new domain is set
					if (! pDomain->m_oNewDomain || pDomain->m_oNewDomain->m_sDomain.IsEmpty	())
					{
						return RTN_ERR_PARAMS;
					}

					// Build a Trustee object with the new domain, old trustee name, and look up the SID
					CTrustee*	pTrusteeNew	=	new CTrustee  (pDomain->m_oNewDomain->m_sDomain + L"\\" + sAccountName, false, 0, false, false);

					if (pTrusteeNew->LookupSID () != RTN_OK)
					{
						m_setaclMain->LogMessage (NoSeverity, L"   Account <" + sAccountName + L"> was not found in domain <" + pDomain->m_oNewDomain->m_sDomain + ">.");

						continue;
					}

					if (pDomain->m_nAction & ACTN_REPLACETRUSTEE)
					{
						// Replace the ACE into a new ACL
						paclACLNew		=	ACLReplaceACE (paclACL, i, pTrusteeNew->m_psidTrustee);

						if (! paclACLNew)
						{
							m_nAPIError	=	GetLastError ();
							delete pTrusteeNew;
                     pTrusteeNew = NULL;
							return RTN_ERR_COPY_ACL;
						}
					}
					else if (pDomain->m_nAction & ACTN_COPYTRUSTEE)
					{
						// Copy the ACE into a new ACL
						paclACLNew		=	ACLCopyACE (paclACL, i, pTrusteeNew->m_psidTrustee);

						if (! paclACLNew)
						{
							m_nAPIError	=	GetLastError ();
							delete pTrusteeNew;
                     pTrusteeNew = NULL;
							return RTN_ERR_COPY_ACL;
						}

						// The ACE count has increased
						nACECount++;
					}

					delete pTrusteeNew;
               pTrusteeNew = NULL;
				}

				if (paclACLNew)
				{
					// Free the old ACL
					if (nWhere == ACL_DACL)
					{
						DeleteBufDACL ();
					}
					else if (nWhere == ACL_SACL)
					{
						DeleteBufSACL ();
					}

					// Continue with the new ACL
					paclACL	=	paclACLNew;

					// Save the new ACL
					if (nWhere == ACL_DACL)
					{
						m_paclDACL			=	paclACL;
						m_fBufDACLAlloc	=	true;
					}
					else if (nWhere == ACL_SACL)
					{
						m_paclSACL			=	paclACL;
						m_fBufSACLAlloc	=	true;
					}
				}

				// Break out of the while loop and continue with next ACE
				break;
			}	// if (EqualSid)
		}		// while (pos)
	}			//	for (int i = 0; i < nACECount; i++)

	return nError;
}


//
// ACLReplaceACE: Replace an ACE in an ACL
//
PACL CSD::ACLReplaceACE (PACL paclACL, DWORD nACE, PSID psidNewTrustee)
{
	ACL_SIZE_INFORMATION				asiACLSize;
	ACCESS_ALLOWED_ACE*				paceACE			=	NULL;
	ACCESS_ALLOWED_ACE*				paceACE2			=	NULL;

	// Do a few checks
	if (! paclACL)
	{
		return NULL;
	}
	if (! IsValidAcl (paclACL))
	{
		return NULL;
	}

	// Get the ACE
	if (! GetAce (paclACL, nACE, (LPVOID*) &paceACE))
	{
		m_nAPIError				=	GetLastError ();
		return NULL;
	}

	// Store the ACE flags for later use
	BYTE			nACEType		=	((PACE_HEADER) paceACE)->AceType;
	BYTE			nACEFlags	=	((PACE_HEADER) paceACE)->AceFlags;
	ACCESS_MASK	nACEMask		=	((ACCESS_ALLOWED_ACE*) paceACE)->Mask;

	// Now delete the ACE
	if (! DeleteAce (paclACL, nACE))
	{
		m_nAPIError				=	GetLastError ();
		return NULL;
	}

	//
	// The new ACE may (!) be of different length (due to the variable length SID) than the deleted one ->
	// We have to initialize a new ACL of correct size and copy all ACEs over; then we can add the new ACE.
	//
	// Get the current size of the ACL excluding the deleted ACE
	if (! GetAclInformation (paclACL, &asiACLSize, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		m_nAPIError				=	GetLastError ();
		return NULL;
	}

	// Allocate memory for the new ACL
	DWORD	nACLNewSize			=	asiACLSize.AclBytesInUse + sizeof (ACCESS_ALLOWED_ACE) + GetLengthSid (psidNewTrustee) - sizeof (DWORD);
	PACL	paclACLNew			=	(PACL) LocalAlloc (LPTR, nACLNewSize);

	if (! paclACLNew)
	{
		m_nAPIError				=	GetLastError ();
		return NULL;
	}

	// Initialize the new ACL
	if(! InitializeAcl (paclACLNew, nACLNewSize, ACL_REVISION))
	{
		m_nAPIError				=	GetLastError ();
		LocalFree (paclACLNew);
		return NULL;
	}

	//
	// Copy all ACEs from the old to the new ACL; insert our new ACE at the correct position
	//
	// The new ACE might belong at the end of the ACL.
	bool	fNewACEInserted	=	false;
	WORD	j;

	for (j = 0; j < asiACLSize.AceCount; j++)
	{
		if (j == nACE)
		{
			// Insert the new ACE
			if (! AddAccessAllowedAce (paclACLNew, ACL_REVISION, nACEMask, psidNewTrustee))
			{
				m_nAPIError	=	GetLastError ();
				LocalFree (paclACLNew);
				return NULL;
			}

			// Get the ACE we just added in order to change the ACE flags
			if (! GetAce (paclACLNew, j, (LPVOID*) &paceACE2))
			{
				m_nAPIError		=	GetLastError ();
				LocalFree (paclACLNew);
				return NULL;
			}

			// Set the original ACE flags on the new ACE
			((PACE_HEADER) paceACE2)->AceType	=	nACEType;
			((PACE_HEADER) paceACE2)->AceFlags	=	nACEFlags;

			// The new ACE does not belong at the end.
			fNewACEInserted	=	true;
		}

		// Get the current ACE from the old ACL
		if (! GetAce (paclACL, j, (LPVOID*) &paceACE2))
		{
			m_nAPIError			=	GetLastError ();
			LocalFree (paclACLNew);
			return NULL;
		}

		// Copy the current ACE from the old to the new ACL
		if(! AddAce (paclACLNew, ACL_REVISION, MAXDWORD, paceACE2, ((PACE_HEADER) paceACE2)->AceSize))
		{
			m_nAPIError			=	GetLastError ();
			LocalFree (paclACLNew);
			return NULL;
		}
	}

	// If the new ACE was not yet been inserted, it belongs at the end of the ACL. Append it now.
	if (! fNewACEInserted)
	{
		// Insert the new ACE
		if (! AddAccessAllowedAce (paclACLNew, ACL_REVISION, nACEMask, psidNewTrustee))
		{
			m_nAPIError			=	GetLastError ();
			LocalFree (paclACLNew);
			return NULL;
		}

		// Get the ACE we just added in order to change the ACE flags
		if (! GetAce (paclACLNew, j, (LPVOID*) &paceACE2))
		{
			m_nAPIError			=	GetLastError ();
			LocalFree (paclACLNew);
			return NULL;
		}

		// Set the original ACE flags on the new ACE
		((PACE_HEADER) paceACE2)->AceType	=	nACEType;
		((PACE_HEADER) paceACE2)->AceFlags	=	nACEFlags;
	}

	// Check for validity
	if (! IsValidAcl (paclACLNew))
	{
		return NULL;
	}

	// return the new ACL
	return paclACLNew;
}


//
// ACLCopyACE: Copy an ACE in an ACL
//
PACL CSD::ACLCopyACE (PACL paclACL, DWORD nACE, PSID psidNewTrustee)
{
	ACL_SIZE_INFORMATION				asiACLSize;
	ACCESS_ALLOWED_ACE*				paceACE			=	NULL;
	ACCESS_ALLOWED_ACE*				paceACE2			=	NULL;

	// Do a few checks
	if (! paclACL)
	{
		return NULL;
	}
	if (! IsValidAcl (paclACL))
	{
		return NULL;
	}

	//
	// We have to initialize a new ACL of correct size and copy all ACEs over; then we can add the new ACE.
	//
	// Get the current size of the ACL
	if (! GetAclInformation (paclACL, &asiACLSize, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		m_nAPIError		=	GetLastError ();
		return NULL;
	}

	// Allocate memory for the new ACL
	DWORD	nACLNewSize	=	asiACLSize.AclBytesInUse + sizeof (ACCESS_ALLOWED_ACE) + GetLengthSid (psidNewTrustee) - sizeof (DWORD);
	PACL	paclACLNew	=	(PACL) LocalAlloc (LPTR, nACLNewSize);

	if (! paclACLNew)
	{
		m_nAPIError		=	GetLastError ();
		return NULL;
	}

	// Initialize the new ACL
	if(! InitializeAcl (paclACLNew, nACLNewSize, ACL_REVISION))
	{
		m_nAPIError	=	GetLastError ();
		LocalFree (paclACLNew);
		return NULL;
	}

	//
	// Copy all ACEs from the old to the new ACL; insert our new ACE at the correct position
	//
	for (WORD j = 0; j < asiACLSize.AceCount; j++)
	{
		// Get the current ACE from the old ACL
		if (! GetAce (paclACL, j, (LPVOID*) &paceACE))
		{
			m_nAPIError	=	GetLastError ();
			LocalFree (paclACLNew);
			return NULL;
		}

		// Copy the current ACE from the old to the new ACL
		if(! AddAce (paclACLNew, ACL_REVISION, MAXDWORD, paceACE, ((PACE_HEADER) paceACE)->AceSize))
		{
			m_nAPIError	=	GetLastError ();
			LocalFree (paclACLNew);
			return NULL;
		}

		if (j == nACE)
		{
			// Insert the new ACE
			if (! AddAccessAllowedAce (paclACLNew, ACL_REVISION, ((ACCESS_ALLOWED_ACE*) paceACE)->Mask, psidNewTrustee))
			{
				m_nAPIError	=	GetLastError ();
				LocalFree (paclACLNew);
				return NULL;
			}

			// Get the ACE we just added in order to change the ACE flags
			if (! GetAce (paclACLNew, j + 1, (LPVOID*) &paceACE2))
			{
				m_nAPIError	=	GetLastError ();
				LocalFree (paclACLNew);
				return NULL;
			}

			// Set the original ACE flags on the new ACE
			((PACE_HEADER) paceACE2)->AceType	=	((PACE_HEADER) paceACE)->AceType;
			((PACE_HEADER) paceACE2)->AceFlags	=	((PACE_HEADER) paceACE)->AceFlags;
		}
	}

	// Check for validity
	if (! IsValidAcl (paclACLNew))
	{
		return NULL;
	}

	// return the new ACL
	return paclACLNew;
}


//
// DeleteACEsByHeaderFlags: Delete all ACEs from an ACL that have certain header flags set
//
DWORD CSD::DeleteACEsByHeaderFlags (DWORD nWhere, BYTE nFlags, bool fFlagsSet)
{
	DWORD									nError			=	RTN_OK;
	ACL_SIZE_INFORMATION				asiACLSize;
	ACCESS_ALLOWED_ACE*				paceACE			=	NULL;
	DWORD									nACECount;
	PACL									paclACL			=	NULL;

	if (nWhere == ACL_DACL)
	{
		paclACL			=	m_paclDACL;
	}
	else if (nWhere == ACL_SACL)
	{
		paclACL			=	m_paclSACL;
	}
	else
	{
		return RTN_ERR_PARAMS;
	}

	// If this is a NULL ACL, do nothing
	if (! paclACL)
	{
		return nError;
	}

	// Get the number of entries in the ACL
	if (! GetAclInformation (paclACL, &asiACLSize, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		m_nAPIError		=	GetLastError ();
		return RTN_ERR_LOOP_ACL;
	}

	nACECount			=	asiACLSize.AceCount;

	// Loop through the ACEs
	for (DWORD i = 0; i < nACECount; i++)
	{
		bool fDelete	=	false;

		// Get the current ACE
		if (! GetAce (paclACL, i, (LPVOID*) &paceACE))
		{
			m_nAPIError	=	GetLastError ();
			return RTN_ERR_LOOP_ACL;
		}

		// Determine whether to delete the ACE depending on the parameters passed
		if (fFlagsSet)
		{
			fDelete		=	(paceACE->Header.AceFlags & nFlags) > 0;
		}
		else
		{
			fDelete		=	! (paceACE->Header.AceFlags & nFlags);
		}

		if (fDelete)
		{
			if (! DeleteAce (paclACL, i))
			{
				m_nAPIError	=	GetLastError ();
				return RTN_ERR_DEL_ACE;
			}

			// The ACECount is now reduced by one!
			nACECount--;
			i--;
		}
	}

	return nError;
}


//
// Set the inherited flag for all pseudo-inherited ACEs
//
DWORD CSD::ConvertPseudoInheritedACEsToInheritedACEs (PACL paclACL, bool fIsContainer, PACL paclParent, DWORD& nRemainingACEs)
{
	DWORD									nError			=	RTN_OK;
	ACL_SIZE_INFORMATION				asiACLSize;
	ACCESS_ALLOWED_ACE*				paceACE			=	NULL;

	// If this is a NULL ACL, do nothing
	if (! paclACL)
	{
		return RTN_OK;
	}

	// Get the number of entries in the ACL
	if (! GetAclInformation (paclACL, &asiACLSize, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		m_nAPIError		=	GetLastError ();
		return RTN_ERR_LOOP_ACL;
	}

	// Store the number of ACEs for our caller
	nRemainingACEs		=	asiACLSize.AceCount;

	// Loop through the ACEs
	for (DWORD i = 0; i < nRemainingACEs; i++)
	{
		// Get the current ACE
		if (! GetAce (paclACL, i, (LPVOID*) &paceACE))
		{
			m_nAPIError	=	GetLastError ();
			return RTN_ERR_LOOP_ACL;
		}

		// Make this ACE inherited if it is pseudo-inherited
		if (m_setaclMain->IsACEPseudoInherited (paceACE, fIsContainer, paclParent))
		{
			paceACE->Header.AceFlags |= INHERITED_ACE;
		}
	}

	return nError;
}


//
// Delete all pseudo-inherited ACEs from an ACL
//
DWORD CSD::DeletePseudoInheritedACEs (DWORD nWhere, bool fIsContainer, PACL paclParent, DWORD& nRemainingACEs)
{
	DWORD									nError			=	RTN_OK;
	ACL_SIZE_INFORMATION				asiACLSize;
	ACCESS_ALLOWED_ACE*				paceACE			=	NULL;
	PACL									paclACL			=	NULL;

	if (paclParent == NULL)
		return RTN_OK;

	if (nWhere == ACL_DACL)
	{
		paclACL			=	m_paclDACL;
	}
	else if (nWhere == ACL_SACL)
	{
		paclACL			=	m_paclSACL;
	}
	else
	{
		return RTN_ERR_PARAMS;
	}

	// If this is a NULL ACL, do nothing
	if (! paclACL)
	{
		return RTN_OK;
	}

	// Get the number of entries in the ACL
	if (! GetAclInformation (paclACL, &asiACLSize, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		m_nAPIError		=	GetLastError ();
		return RTN_ERR_LOOP_ACL;
	}

	// Store the number of ACEs for our caller
	nRemainingACEs		=	asiACLSize.AceCount;

	// Loop through the ACEs
	for (DWORD i = 0; i < nRemainingACEs; i++)
	{
		// Get the current ACE
		if (! GetAce (paclACL, i, (LPVOID*) &paceACE))
		{
			m_nAPIError	=	GetLastError ();
			return RTN_ERR_LOOP_ACL;
		}

		// Delete this ACE if it is pseudo-inherited
		if (m_setaclMain->IsACEPseudoInherited (paceACE, fIsContainer, paclParent))
		{
			if (! DeleteAce (paclACL, i))
			{
				m_nAPIError	=	GetLastError ();
				return RTN_ERR_DEL_ACE;
			}

			// The ACECount is now reduced by one!
			nRemainingACEs--;
			i--;
		}
	}

	return nError;
}


//
// GetPermissions: Return a string with the permissions in an access mask
//
CString CSetACL::GetPermissions (ACCESS_MASK nAccessMask)
{
	CString	sPermissions;

	// Map generic rights to standard plus specific rights
	nAccessMask	= MapGenericRight (nAccessMask, m_nObjectType);

	if (m_nObjectType	==	SE_FILE_OBJECT)
	{
		// Mask out the SYNCHRONIZE flag which is not always set (ie. not on audit ACEs)
		nAccessMask			&=	~SYNCHRONIZE;

		if ((nAccessMask & MY_DIR_FULL_ACCESS) == MY_DIR_FULL_ACCESS)
		{
			sPermissions	+=	L"full+";
			nAccessMask		&= ~MY_DIR_FULL_ACCESS;
		}
		if ((nAccessMask & MY_DIR_CHANGE_ACCESS) == MY_DIR_CHANGE_ACCESS)
		{
			sPermissions	+=	L"change+";
			nAccessMask		&= ~MY_DIR_CHANGE_ACCESS;
		}
		if ((nAccessMask & MY_DIR_READ_EXECUTE_ACCESS) == MY_DIR_READ_EXECUTE_ACCESS)
		{
			sPermissions	+=	L"read_execute+";
			nAccessMask		&= ~MY_DIR_READ_EXECUTE_ACCESS;
		}
		if ((nAccessMask & MY_DIR_WRITE_ACCESS) == MY_DIR_WRITE_ACCESS)
		{
			sPermissions	+=	L"write+";
			nAccessMask		&= ~MY_DIR_WRITE_ACCESS;
		}
		if ((nAccessMask & MY_DIR_READ_ACCESS) == MY_DIR_READ_ACCESS)
		{
			sPermissions	+=	L"read+";
			nAccessMask		&= ~MY_DIR_READ_ACCESS;
		}

		if (nAccessMask & FILE_LIST_DIRECTORY)
			  sPermissions += L"FILE_LIST_DIRECTORY+";
		if (nAccessMask & FILE_ADD_FILE)
			  sPermissions += L"FILE_ADD_FILE+";
		if (nAccessMask & FILE_ADD_SUBDIRECTORY)
			  sPermissions += L"FILE_ADD_SUBDIRECTORY+";
		if (nAccessMask & FILE_READ_EA)
			  sPermissions += L"FILE_READ_EA+";
		if (nAccessMask & FILE_WRITE_EA)
			  sPermissions += L"FILE_WRITE_EA+";
		if (nAccessMask & FILE_TRAVERSE)
			  sPermissions += L"FILE_TRAVERSE+";
		if (nAccessMask & FILE_DELETE_CHILD)
			  sPermissions += L"FILE_DELETE_CHILD+";
		if (nAccessMask & FILE_READ_ATTRIBUTES)
			  sPermissions += L"FILE_READ_ATTRIBUTES+";
		if (nAccessMask & FILE_WRITE_ATTRIBUTES)
			  sPermissions += L"FILE_WRITE_ATTRIBUTES+";
	}
	else if (m_nObjectType	==	SE_REGISTRY_KEY)
	{
		if ((nAccessMask & MY_REG_FULL_ACCESS) == MY_REG_FULL_ACCESS)
		{
			sPermissions	+=	L"full+";
			nAccessMask		&= ~MY_REG_FULL_ACCESS;
		}
		if ((nAccessMask & MY_REG_READ_ACCESS) == MY_REG_READ_ACCESS)
		{
			sPermissions	+=	L"read+";
			nAccessMask		&= ~MY_REG_READ_ACCESS;
		}

		if (nAccessMask & KEY_CREATE_LINK)
			  sPermissions += L"KEY_CREATE_LINK+";
		if (nAccessMask & KEY_CREATE_SUB_KEY)
			  sPermissions += L"KEY_CREATE_SUB_KEY+";
		if (nAccessMask & KEY_ENUMERATE_SUB_KEYS)
			  sPermissions += L"KEY_ENUMERATE_SUB_KEYS+";
		if (nAccessMask & KEY_EXECUTE)
			  sPermissions += L"KEY_EXECUTE+";
		if (nAccessMask & KEY_NOTIFY)
			  sPermissions += L"KEY_NOTIFY+";
		if (nAccessMask & KEY_QUERY_VALUE)
			  sPermissions += L"KEY_QUERY_VALUE+";
		if (nAccessMask & KEY_READ)
			  sPermissions += L"KEY_READ+";
		if (nAccessMask & KEY_SET_VALUE)
			  sPermissions += L"KEY_SET_VALUE+";
		if (nAccessMask & KEY_WRITE)
			  sPermissions += L"KEY_WRITE+";
	}
	else if (m_nObjectType	==	SE_SERVICE)
	{
		if ((nAccessMask & MY_SVC_FULL_ACCESS) == MY_SVC_FULL_ACCESS)
		{
			sPermissions	+=	L"full+";
			nAccessMask		&= ~MY_SVC_FULL_ACCESS;
		}
		if ((nAccessMask & MY_SVC_STARTSTOP_ACCESS) == MY_SVC_STARTSTOP_ACCESS)
		{
			sPermissions	+=	L"start_stop+";
			nAccessMask		&= ~MY_SVC_STARTSTOP_ACCESS;
		}
		if ((nAccessMask & MY_SVC_READ_ACCESS) == MY_SVC_READ_ACCESS)
		{
			sPermissions	+=	L"read+";
			nAccessMask		&= ~MY_SVC_READ_ACCESS;
		}

		if (nAccessMask & SERVICE_CHANGE_CONFIG)
			  sPermissions += L"SERVICE_CHANGE_CONFIG+";
		if (nAccessMask & SERVICE_ENUMERATE_DEPENDENTS)
			  sPermissions += L"SERVICE_ENUMERATE_DEPENDENTS+";
		if (nAccessMask & SERVICE_INTERROGATE)
			  sPermissions += L"SERVICE_INTERROGATE+";
		if (nAccessMask & SERVICE_PAUSE_CONTINUE)
			  sPermissions += L"SERVICE_PAUSE_CONTINUE+";
		if (nAccessMask & SERVICE_QUERY_CONFIG)
			  sPermissions += L"SERVICE_QUERY_CONFIG+";
		if (nAccessMask & SERVICE_QUERY_STATUS)
			  sPermissions += L"SERVICE_QUERY_STATUS+";
		if (nAccessMask & SERVICE_START)
			  sPermissions += L"SERVICE_START+";
		if (nAccessMask & SERVICE_STOP)
			  sPermissions += L"SERVICE_STOP+";
		if (nAccessMask & SERVICE_USER_DEFINED_CONTROL)
			  sPermissions += L"SERVICE_USER_DEFINED_CONTROL+";
	}
	else if (m_nObjectType	==	SE_PRINTER)
	{
		if ((nAccessMask & MY_PRINTER_MAN_PRINTER_ACCESS) == MY_PRINTER_MAN_PRINTER_ACCESS)
		{
			sPermissions	+=	L"manage_printer+";
			nAccessMask		&= ~MY_PRINTER_MAN_PRINTER_ACCESS;
		}
		if ((nAccessMask & MY_PRINTER_MAN_DOCS_ACCESS) == MY_PRINTER_MAN_DOCS_ACCESS)
		{
			sPermissions	+=	L"manage_documents+";
			nAccessMask		&= ~MY_PRINTER_MAN_DOCS_ACCESS;
		}
		if ((nAccessMask & MY_PRINTER_PRINT_ACCESS) == MY_PRINTER_PRINT_ACCESS)
		{
			sPermissions	+=	L"print+";
			nAccessMask		&= ~MY_PRINTER_PRINT_ACCESS;
		}

		if (nAccessMask & PRINTER_ACCESS_ADMINISTER)
			  sPermissions += L"PRINTER_ACCESS_ADMINISTER+";
		if (nAccessMask & PRINTER_ACCESS_USE)
			  sPermissions += L"PRINTER_ACCESS_USE+";
		if (nAccessMask & JOB_ACCESS_ADMINISTER)
			  sPermissions += L"JOB_ACCESS_ADMINISTER+";
		if (nAccessMask & JOB_ACCESS_READ)
			  sPermissions += L"JOB_ACCESS_READ+";
	}
	else if (m_nObjectType	==	SE_LMSHARE)
	{
		if ((nAccessMask & MY_SHARE_FULL_ACCESS) == MY_SHARE_FULL_ACCESS)
		{
			sPermissions	+=	L"full+";
			nAccessMask		&= ~MY_SHARE_FULL_ACCESS;
		}
		if ((nAccessMask & MY_SHARE_CHANGE_ACCESS) == MY_SHARE_CHANGE_ACCESS)
		{
			sPermissions	+=	L"change+";
			nAccessMask		&= ~MY_SHARE_CHANGE_ACCESS;
		}
		if ((nAccessMask & MY_SHARE_READ_ACCESS) == MY_SHARE_READ_ACCESS)
		{
			sPermissions	+=	L"read+";
			nAccessMask		&= ~MY_SHARE_READ_ACCESS;
		}

		if (nAccessMask & SHARE_READ)
			  sPermissions += L"SHARE_READ+";
		if (nAccessMask & SHARE_CHANGE)
			  sPermissions += L"SHARE_CHANGE+";
		if (nAccessMask & SHARE_WRITE)
			  sPermissions += L"SHARE_WRITE+";
	}
	else if (m_nObjectType	==	SE_WMIGUID_OBJECT)
	{
		if ((nAccessMask & MY_WMI_FULL_ACCESS) == MY_WMI_FULL_ACCESS)
		{
			sPermissions	+=	L"full+";
			nAccessMask		&= ~MY_WMI_FULL_ACCESS;
		}
		if ((nAccessMask & MY_WMI_EXECUTE_ACCESS) == MY_WMI_EXECUTE_ACCESS)
		{
			sPermissions	+=	L"execute+";
			nAccessMask		&= ~MY_WMI_EXECUTE_ACCESS;
		}
		if ((nAccessMask & MY_WMI_REMOTE_ACCESS) == MY_WMI_REMOTE_ACCESS)
		{
			sPermissions	+=	L"remote_access+";
			nAccessMask		&= ~MY_WMI_REMOTE_ACCESS;
		}
		if ((nAccessMask & MY_WMI_ENABLE_ACCOUNT) == MY_WMI_ENABLE_ACCOUNT)
		{
			sPermissions	+=	L"enable_account+";
			nAccessMask		&= ~MY_WMI_ENABLE_ACCOUNT;
		}

		if (nAccessMask & WBEM_ENABLE)
			  sPermissions += L"WBEM_ENABLE+";
		if (nAccessMask & WBEM_METHOD_EXECUTE)
			  sPermissions += L"WBEM_METHOD_EXECUTE+";
		if (nAccessMask & WBEM_FULL_WRITE_REP)
			  sPermissions += L"WBEM_FULL_WRITE_REP+";
		if (nAccessMask & WBEM_PARTIAL_WRITE_REP)
			  sPermissions += L"WBEM_PARTIAL_WRITE_REP+";
		if (nAccessMask & WBEM_WRITE_PROVIDER)
			  sPermissions += L"WBEM_WRITE_PROVIDER+";
		if (nAccessMask & WBEM_REMOTE_ACCESS)
			  sPermissions += L"WBEM_REMOTE_ACCESS+";
		if (nAccessMask & WBEM_RIGHT_SUBSCRIBE)
			  sPermissions += L"WBEM_RIGHT_SUBSCRIBE+";
		if (nAccessMask & WBEM_RIGHT_PUBLISH)
			  sPermissions += L"WBEM_RIGHT_PUBLISH+";
	}

	if (nAccessMask & READ_CONTROL)
			sPermissions += L"READ_CONTROL+";
	if (nAccessMask & WRITE_OWNER)
			sPermissions += L"WRITE_OWNER+";
	if (nAccessMask & WRITE_DAC)
			sPermissions += L"WRITE_DAC+";
	if (nAccessMask & DELETE)
			sPermissions += L"DELETE+";
	if (nAccessMask & SYNCHRONIZE)
			sPermissions += L"SYNCHRONIZE+";
	if (nAccessMask & ACCESS_SYSTEM_SECURITY)
			sPermissions += L"ACCESS_SYSTEM_SECURITY+";
	if (nAccessMask & GENERIC_ALL)
			sPermissions += L"GENERIC_ALL+";
	if (nAccessMask & GENERIC_EXECUTE)
			sPermissions += L"GENERIC_EXECUTE+";
	if (nAccessMask & GENERIC_READ)
			sPermissions += L"GENERIC_READ+";
	if (nAccessMask & GENERIC_WRITE)
			sPermissions += L"GENERIC_WRITE+";

	sPermissions.TrimRight (L"+");

	return sPermissions;
}


//
// GetACEType: Return a string with the type of an ACE
//
CString CSetACL::GetACEType (BYTE nACEType)
{
	CString	sACEType;

	switch(nACEType)
	{
	case ACCESS_ALLOWED_ACE_TYPE:
		sACEType	=	L"allow";
		break;
	case ACCESS_ALLOWED_CALLBACK_ACE_TYPE:
		sACEType	=	L"allow_callback";
		break;
	case ACCESS_ALLOWED_CALLBACK_OBJECT_ACE_TYPE:
		sACEType	=	L"allow_callback_object";
		break;
	case ACCESS_ALLOWED_COMPOUND_ACE_TYPE:
		sACEType	=	L"allow_compound";
		break;
	case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
		sACEType	=	L"allow_object";
		break;
	case ACCESS_DENIED_ACE_TYPE:
		sACEType	=	L"deny";
		break;
	case ACCESS_DENIED_CALLBACK_ACE_TYPE:
		sACEType	=	L"deny_callback";
		break;
	case ACCESS_DENIED_CALLBACK_OBJECT_ACE_TYPE:
		sACEType	=	L"deny_callback_object";
		break;
	case ACCESS_DENIED_OBJECT_ACE_TYPE:
		sACEType	=	L"deny_object";
		break;
	case SYSTEM_ALARM_ACE_TYPE:
		sACEType	=	L"alarm";
		break;
	case SYSTEM_ALARM_CALLBACK_ACE_TYPE:
		sACEType	=	L"alarm_callback";
		break;
	case SYSTEM_ALARM_CALLBACK_OBJECT_ACE_TYPE:
		sACEType	=	L"alarm_callback_object";
		break;
	case SYSTEM_ALARM_OBJECT_ACE_TYPE:
		sACEType	=	L"alarm_object";
		break;
	case SYSTEM_AUDIT_ACE_TYPE:
		sACEType	=	L"audit";
		break;
	case SYSTEM_AUDIT_CALLBACK_ACE_TYPE:
		sACEType	=	L"audit_callback";
		break;
	case SYSTEM_AUDIT_CALLBACK_OBJECT_ACE_TYPE:
		sACEType	=	L"audit_callback_object";
		break;
	case SYSTEM_AUDIT_OBJECT_ACE_TYPE:
		sACEType	=	L"audit_object";
		break;
	case SYSTEM_MANDATORY_LABEL_ACE_TYPE:
		sACEType	=	L"mandatory_label";
		break;
	}

	return sACEType;
}


//
// GetACEFlags: Return a string with the flags of an ACE
//
CString CSetACL::GetACEFlags (BYTE nACEFlags, bool fACEIsPseudoInherited)
{
	CString	sFlags;

	if (fACEIsPseudoInherited)
		 sFlags	+=	L"pseudo_inherited+";
	if (nACEFlags & CONTAINER_INHERIT_ACE)
		 sFlags	+=	L"container_inherit+";
	if (nACEFlags & OBJECT_INHERIT_ACE)
		 sFlags	+= L"object_inherit+";
	if (nACEFlags & INHERIT_ONLY_ACE)
		 sFlags	+= L"inherit_only+";
	if (nACEFlags & NO_PROPAGATE_INHERIT_ACE)
		 sFlags	+= L"no_propagate_inherit+";
	if (nACEFlags & INHERITED_ACE)
		 sFlags	+= L"inherited+";
	if (nACEFlags & SUCCESSFUL_ACCESS_ACE_FLAG)
		 sFlags	+= L"audit_success+";
	if (nACEFlags & FAILED_ACCESS_ACE_FLAG)
		 sFlags	+= L"audit_fail+";

	sFlags.TrimRight (L"+");

	if (sFlags.IsEmpty ())
	{
		sFlags	=	L"no_inheritance";
	}

	return sFlags;
}


//
// SetPrivilege: Enable a privilege (user right) for the current process
//
DWORD CSetACL::SetPrivilege (CString sPrivilege, bool fEnable, bool bLogErrors)
{
	HANDLE				hToken	=	NULL;		// handle to process token
	TOKEN_PRIVILEGES	tkp;						// pointer to token structure

	// Get the current process token handle
	if (! OpenProcessToken (GetCurrentProcess (), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
	{
		return RTN_ERR_EN_PRIV;
	}

	// Get the LUID for the privilege on the target (or local) system
	if (! LookupPrivilegeValue (m_sTargetSystemName, sPrivilege, &tkp.Privileges[0].Luid))
	{
		CloseHandle (hToken);
		return RTN_ERR_EN_PRIV;
	}

	// One privilege to set
	tkp.PrivilegeCount					=	1;
	tkp.Privileges[0].Attributes		=	(fEnable ? SE_PRIVILEGE_ENABLED : 0);

	// Enable the privilege
	if (! AdjustTokenPrivileges (hToken, false, &tkp, 0, (PTOKEN_PRIVILEGES) NULL, 0) || GetLastError ())
	{
		if (bLogErrors)
		{
			// Log in detail what went wrong
			CString sEnDis;
			if (fEnable)
				sEnDis = L"Enabling";
			else
				sEnDis = L"Disabling";
			LogMessage (Error, sEnDis + L" the privilege " + sPrivilege + L" failed with: " + GetLastErrorMessage());
		}

		CloseHandle (hToken);
		return RTN_ERR_EN_PRIV;
	}

	CloseHandle (hToken);
	return RTN_OK;
}


//
// Check if an ACE is pseudo-inherited from its parent
//
bool CSetACL::IsACEPseudoInherited (ACCESS_ALLOWED_ACE* paceObject, bool fObjectIsContainer, PACL paclParent)
{
	ACL_SIZE_INFORMATION		asiParent;
	ACCESS_ALLOWED_ACE*		paceParent			=	NULL;
	DWORD							nAccessMaskChild	=	0;
	DWORD							nAccessMaskParent	=	0;

	// Check parameters
	if (paceObject == NULL || paclParent == NULL)
	{
		return false;
	}

	// Get the number of entries in the parent's ACL
	if (! GetAclInformation (paclParent, &asiParent, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		return false;
	}

	// Loop through the ACEs in the parent's ACL
	for (WORD i = 0; i < asiParent.AceCount; i++)
	{
		// Get the current ACE
		if (! GetAce (paclParent, i, (LPVOID*) &paceParent))
		{
			continue;
		}

		// Check if the object ACE passed in and the current parent ACE pertain the same SID
		if (! EqualSid ((PSID) &(paceObject->SidStart), (PSID) &(paceParent->SidStart)))
		{
			continue;
		}

		// Check if both ACEs have the same type (access allowed, denied, etc.)
		if (paceObject->Header.AceType != paceParent->Header.AceType)
		{
			continue;
		}

		// Check if both ACEs have the same access mask (permission)
		// First replace (map) generic rights with their standard+specific counterparts
		nAccessMaskChild	= MapGenericRight (paceObject->Mask, m_nObjectType);
		nAccessMaskParent	= MapGenericRight (paceParent->Mask, m_nObjectType);
		// Now compare the resulting masks
		if (nAccessMaskChild != nAccessMaskParent)
		{
			continue;
		}

		if (fObjectIsContainer)
		{
			// The object is a container (e.g. a directory or registry key)

			if (((paceParent->Header.AceFlags & OBJECT_INHERIT_ACE) == false &&
				(paceParent->Header.AceFlags & CONTAINER_INHERIT_ACE) == false))
			{
				// The parent object does not propagate anything
				continue;
			}

			if (paceParent->Header.AceFlags & NO_PROPAGATE_INHERIT_ACE)
			{
				// This parent ACE is inheritable to child containers, but not to be propagated further ("one level only")

				if ((paceObject->Header.AceFlags & OBJECT_INHERIT_ACE) == false && 
					(paceObject->Header.AceFlags & CONTAINER_INHERIT_ACE) == false &&
					(paceObject->Header.AceFlags & INHERIT_ONLY_ACE) == false &&
					(paceObject->Header.AceFlags & NO_PROPAGATE_INHERIT_ACE) == false)
				{
					// The child does not have inheritance flags set -> this ACe is pseudo-inherited
					return true;
				}
			}
			else
			{
				// This parent ACE has object and/or container inherit flags set (the most common case)

				// Compare the relevant inheritance flags of parent and child
				if ((paceObject->Header.AceFlags & (OBJECT_INHERIT_ACE | CONTAINER_INHERIT_ACE | INHERIT_ONLY_ACE)) ==
					(paceParent->Header.AceFlags & (OBJECT_INHERIT_ACE | CONTAINER_INHERIT_ACE | INHERIT_ONLY_ACE)))
				{
					// Flags are identical -> this ACE is pseudo-inherited
					return true;
				}
			}

			if (paceParent->Header.AceFlags & INHERIT_ONLY_ACE && 
						paceParent->Header.AceFlags & CONTAINER_INHERIT_ACE)
			{
				// The parent has an ACE like the following: Everyone   full   allow   container_inherit+object_inherit+inherit_only
				// One level deeper, this becomes two (!) ACEs: a copy of itself plus the following: Everyone   full   allow   no_inheritance

				if ((paceObject->Header.AceFlags & OBJECT_INHERIT_ACE) == false && 
					(paceObject->Header.AceFlags & CONTAINER_INHERIT_ACE) == false &&
					(paceObject->Header.AceFlags & INHERIT_ONLY_ACE) == false &&
					(paceObject->Header.AceFlags & NO_PROPAGATE_INHERIT_ACE) == false)
				{
					// The child does not have inheritance flags set -> this ACe is pseudo-inherited
					return true;
				}
			}
		}
		else
		{
			// The object is not a container (e.g. a file)

			if ((paceParent->Header.AceFlags & OBJECT_INHERIT_ACE) == false)
			{
				// The parent object does not propagate anything
				continue;
			}

			// The ACE is inheritable to child objects (at least one level - NO_PROPAGATE_INHERIT_ACE is of no relevance here)
			// -> this ACE is pseudo-inherited
			return true;
		}

	}

	return false;
}


//
// Check if an ACL is pseudo-protected: the parent contains inheritable ACEs not present in the child
//
bool CSetACL::IsACLPseudoProtected (PACL paclObject, bool fObjectIsContainer, PSID psidObjectOwner, PACL paclParent)
{
	ACL_SIZE_INFORMATION		asiObject;
	ACL_SIZE_INFORMATION		asiParent;
	ACCESS_ALLOWED_ACE*		paceObject			=	NULL;
	ACCESS_ALLOWED_ACE*		paceParent			=	NULL;
	bool							fFound				=	false;
	DWORD							nAccessMaskChild	=	0;
	DWORD							nAccessMaskParent	=	0;
	PSID							psidParentOrOwner	=	0;

	// Check parameters
	if (paclObject == NULL || paclParent == NULL)
	{
		return false;
	}

	// Get the number of entries in the object's ACL...
	if (! GetAclInformation (paclObject, &asiObject, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		return false;
	}
	// ...and in the parent's ACL
	if (! GetAclInformation (paclParent, &asiParent, sizeof (ACL_SIZE_INFORMATION), AclSizeInformation))
	{
		return false;
	}

	//
	// Loop through the parent's ACL and look for inheritable ACEs not present in the child
	//
	for (WORD i = 0; i < asiParent.AceCount; i++)
	{
		// Get the current parent ACE
		if (! GetAce (paclParent, i, (LPVOID*) &paceParent))
		{
			continue;
		}

		// Is this an inheritable ACE?
		if (fObjectIsContainer)
		{
			if (((paceParent->Header.AceFlags & OBJECT_INHERIT_ACE) == false &&
				(paceParent->Header.AceFlags & CONTAINER_INHERIT_ACE) == false))
			{
				// Not an inheritable ACE
				continue;
			}
		}
		else
		{
			if ((paceParent->Header.AceFlags & OBJECT_INHERIT_ACE) == false)
			{
				// Not an inheritable ACE
				continue;
			}
		}

		// The parent ACE might pertain to CREATOR_OWNER. In that case it is replaced by the SID of the owner of the child.
		// In the constellation we have here this only happens for files.
		if (EqualSid ((PSID) &(paceParent->SidStart), m_WellKnownSIDCreatorOwner.m_psidTrustee) && fObjectIsContainer == false)
		{
			// CREATOR_OWNER stored in parent ACE: use the object owner's SID instead of the CREATOR_OWNER
			psidParentOrOwner	=	psidObjectOwner;
		}
		else
		{
			// Something other than CREATOR_OWNER stored in parent ACE: use whatever is stored in the ACE
			psidParentOrOwner	=	(PSID) &(paceParent->SidStart);
		}

		// Now loop through the child's ACEs and check if the parent's inheritable ACE is present (i.e. has been inherited)
		fFound	= false;
		for (WORD j = 0; j < asiObject.AceCount; j++)
		{
			// Get the current child ACE
			if (! GetAce (paclObject, j, (LPVOID*) &paceObject))
			{
				continue;
			}

			// Check if the object ACE and the parent ACE pertain the same SID
			if (! EqualSid ((PSID) &(paceObject->SidStart), psidParentOrOwner))
			{
				continue;
			}

			// Check if both ACEs have the same type (access allowed, denied, etc.)
			if (paceObject->Header.AceType != paceParent->Header.AceType)
			{
				continue;
			}

			// Check if both ACEs have the same access mask (permission)
			// First replace (map) generic rights with their standard+specific counterparts
			nAccessMaskChild	= MapGenericRight (paceObject->Mask, m_nObjectType);
			nAccessMaskParent	= MapGenericRight (paceParent->Mask, m_nObjectType);
			// Now compare the resulting masks
			if (nAccessMaskChild != nAccessMaskParent)
			{
				continue;
			}

			if (fObjectIsContainer)
			{
				// The object is a container (e.g. a directory or registry key)

				if (paceParent->Header.AceFlags & NO_PROPAGATE_INHERIT_ACE)
				{
					// This parent ACE is inheritable to child containers, but not to be propagated further ("one level only")

					if ((paceObject->Header.AceFlags & OBJECT_INHERIT_ACE) == false && 
						(paceObject->Header.AceFlags & CONTAINER_INHERIT_ACE) == false &&
						(paceObject->Header.AceFlags & INHERIT_ONLY_ACE) == false &&
						(paceObject->Header.AceFlags & NO_PROPAGATE_INHERIT_ACE) == false)
					{
						// The child does not have inheritance flags set -> this ACe is (pseudo-) inherited
						fFound	= true;
						break;
					}
				}
				else
				{
					// This parent ACE has object and/or container inherit flags set (the most common case)

					// Compare the relevant inheritance flags of parent and child
					if ((paceObject->Header.AceFlags & (OBJECT_INHERIT_ACE | CONTAINER_INHERIT_ACE)) ==
						(paceParent->Header.AceFlags & (OBJECT_INHERIT_ACE | CONTAINER_INHERIT_ACE)))
					{
						// Check flag INHERIT_ONLY_ACE: if the child is inherit only, the parent must also be inherit only
						if ((paceObject->Header.AceFlags & INHERIT_ONLY_ACE) && (! (paceParent->Header.AceFlags & INHERIT_ONLY_ACE)))
						{
							continue;
						}

						// Flags are identical -> this ACE is (pseudo-) inherited
						fFound	= true;
						break;
					}
				}
			}
			else
			{
				// The object is not a container (e.g. a file)

				// The ACE is inheritable to child objects (at least one level - NO_PROPAGATE_INHERIT_ACE is of no relevance here)
				// -> this ACE is (pseudo-) inherited
				fFound	= true;
				break;
			}
		}

		// If the parent's inheritable ACE was not found in the child, then it is pseudo-protected
		if (! fFound)
		{
			return true;
		}
	}

	return false;
}


//
// Retrieve the last output generated by the list function
//
CString CSetACL::GetLastListOutput (void)
{
	return m_sListOutput;
}


//
// Set if raw mode is to be used (where no ACL prettying up is done)
//
void CSetACL::SetRawMode (bool fRawMode)
{
	m_fRawMode	= fRawMode;
}

//////////////////////////////////////////////////////////////////////
//
// Class CTrustee
//
//////////////////////////////////////////////////////////////////////


//
// Constructor: initialize all member variables
//
CTrustee::CTrustee (CString sTrustee, bool fTrusteeIsSID, DWORD nAction, bool fDACL, bool fSACL)
{
	m_sTrustee			=	sTrustee;
	m_fTrusteeIsSID	=	fTrusteeIsSID;
	m_psidTrustee		=	NULL;
	m_nAction			=	nAction;
	m_fDACL				=	fDACL;
	m_fSACL				=	fSACL;
	m_oNewTrustee		=	NULL;
}


CTrustee::CTrustee ()
{
	m_sTrustee			=	TEXT ("");
	m_fTrusteeIsSID	=	false;
	m_psidTrustee		=	NULL;
	m_nAction			=	0;
	m_fDACL				=	false;
	m_fSACL				=	false;
	m_oNewTrustee		=	NULL;
}


//
// Destructor: clean up
//
CTrustee::~CTrustee ()
{
	if (m_psidTrustee)
	{
		LocalFree (m_psidTrustee);
	}
}


//
// LookupSID: Lookup/calculate the binary SID of a given trustee
//
DWORD CTrustee::LookupSID ()
{
	// Is there a valid trustee object?
	if (m_sTrustee.IsEmpty ())
	{
		return RTN_OK;
	}

	// If there already is a binary SID, delete it (playing it safe)
	if (m_psidTrustee)
	{
		LocalFree (m_psidTrustee);
		m_psidTrustee		=	NULL;
	}

	// If the trustee was specified as a string SID, it has to be converted into a binary SID.
	if (m_fTrusteeIsSID)
	{
		if (! ConvertStringSidToSid (m_sTrustee, &m_psidTrustee))
		{
			m_psidTrustee	=	NULL;

			return RTN_ERR_LOOKUP_SID;
		}

		return RTN_OK;
	}

	//
	// The trustee is not a (string) SID. Look up the SID for the account name.
	//
	CString							sSystemName;
	CString							sDomainNameFound;
	SID_NAME_USE					eUse;
	PDOMAIN_CONTROLLER_INFO		pdcInfo				=	NULL;

	// The trustee name can have a preceding machine/domain name, or not. Try to isolate the machine/domain.
	int nBackslashPos		= m_sTrustee.Find (L"\\");
	if (nBackslashPos != -1)
	{
		sSystemName			=	m_sTrustee.Left (nBackslashPos);
	}

	//
	// If a machine/domain name was specified, we assume it is a domain name, and search for a DC.
	// But first check if the domain name is identical to the local computer name. In that case, ignore it.
	//
	// Get the computer name
	CString sComputername;
	sComputername.GetEnvironmentVariable(L"computername");

	// Now check
	if (sSystemName.IsEmpty () == false && sComputername.CompareNoCase(sSystemName) != 0)
	{
		// If a DC can be found, we use it. If not, the name might be a computer name.
		if (DsGetDcName (NULL, sSystemName, NULL, NULL, 0, &pdcInfo) == ERROR_SUCCESS)
		{
			// We now have the name of a domain controller.
			// The names returned have the same format as those passed in, i.e. if we pass in DNS names, we get back DNS names, and vice versa.
			sSystemName		=	pdcInfo->DomainControllerName;

			// Remove the leading backslashes in the name
			if (sSystemName.Left (2) == L"\\\\")
			{
				sSystemName	=	sSystemName.Right (sSystemName.GetLength () - 2);
			}

			// Free the buffer allocated by the API function
			if (pdcInfo)
			{
				NetApiBufferFree (pdcInfo);
				pdcInfo	=	NULL;
			}
		}
	}

	//
	// We'll search for the account on either a DC found, a remote machine specified or on the local computer.
	//
	// First get buffer sizes
	DWORD	nBufTrustee	=	0;
	DWORD	nBufDomain	=	0;
	LookupAccountName (sSystemName.IsEmpty() ? NULL : (LPCTSTR) sSystemName, m_sTrustee, NULL, &nBufTrustee, NULL, &nBufDomain, &eUse);

	// Second, allocate memory for the SID
	PSID	psidTrustee	=	LocalAlloc (LPTR, nBufTrustee);

	// Third, look up the name
	if (LookupAccountName (sSystemName.IsEmpty() ? NULL : (LPCTSTR) sSystemName, m_sTrustee, psidTrustee, &nBufTrustee, 
									sDomainNameFound.GetBuffer(nBufDomain), &nBufDomain, &eUse))
	{
		// We found the correct SID. Store it.
		m_psidTrustee			=	CopySID (psidTrustee);
	}

	// Free buffers
	sDomainNameFound.ReleaseBuffer ();
	if (psidTrustee)
	{
		LocalFree (psidTrustee);
		psidTrustee	= NULL;
	}

	if (m_psidTrustee)
	{
		return RTN_OK;
	}
	else
	{
		return RTN_ERR_LOOKUP_SID;
	}
}


//////////////////////////////////////////////////////////////////////
//
// Class CDomain
//
//////////////////////////////////////////////////////////////////////


//
// Constructor: initialize all member variables
//
CDomain::CDomain ()
{
	m_sDomain.Empty ();
	m_nAction			=	0;
	m_fDACL				=	false;
	m_fSACL				=	false;
	m_oNewDomain		=	NULL;
}


//
// Destructor: clean up
//
CDomain::~CDomain ()
{
}


//
// SetDomain: Set the domain and do some initialization
//
DWORD CDomain::SetDomain (CString sDomain, DWORD nAction, bool fDACL, bool fSACL)
{
	PDOMAIN_CONTROLLER_INFO		pdcInfo				=	NULL;

	// Reset all member variables
	m_sDomain.Empty ();
	m_nAction			=	0;
	m_fDACL				=	false;
	m_fSACL				=	false;

	delete m_oNewDomain;
	m_oNewDomain = NULL;

	if (sDomain.IsEmpty ())
	{
		return RTN_ERR_INV_DOMAIN;
	}

	// Get the NetBIOS name of the domain
	if (DsGetDcName (NULL, sDomain, NULL, NULL, DS_RETURN_FLAT_NAME, &pdcInfo) == ERROR_SUCCESS)
	{
		m_sDomain		=	pdcInfo->DomainName;

		// Free the buffer allocated by the API function
		if (pdcInfo)
		{
			NetApiBufferFree (pdcInfo);
		}
	}
	else
	{
		return RTN_ERR_INV_DOMAIN;
	}

	m_nAction			=	nAction;
	m_fDACL				=	fDACL;
	m_fSACL				=	fSACL;
	m_oNewDomain		=	NULL;

	return RTN_OK;
}


//////////////////////////////////////////////////////////////////////
//
// Class CACE
//
//////////////////////////////////////////////////////////////////////


//
// Constructor: initialize all member variables
//
CACE::CACE (CTrustee* pTrustee, CString sPermission, DWORD nInheritance, bool fInhSpecified, ACCESS_MODE nAccessMode, DWORD nACLType)
{
	if (pTrustee)
	{
		m_pTrustee		=	pTrustee;
	}
	else
	{
		m_pTrustee		=	NULL;
	}

	m_sPermission		=	sPermission;
	m_nInheritance		=	nInheritance;
	m_fInhSpecified	=	fInhSpecified;
	m_nAccessMode		=	nAccessMode;
	m_nACLType			=	nACLType;
	m_nAccessMask		=	0;
}


CACE::CACE ()
{
	m_pTrustee			=	NULL;
	m_sPermission		=	TEXT ("");
	m_nInheritance		=	0;
	m_fInhSpecified	=	false;
	m_nAccessMode		=	NOT_USED_ACCESS;
	m_nACLType			=	0;
	m_nAccessMask		=	0;
}


//
// Destructor: clean up
//
CACE::~CACE ()
{
	delete m_pTrustee;
	m_pTrustee = NULL;
}


//////////////////////////////////////////////////////////////////////
//
// Class CSD
//
//////////////////////////////////////////////////////////////////////


//
// Constructor: initialize all member variables
//
CSD::CSD (CSetACL* setaclMain)
{
	m_setaclMain				=	setaclMain;

	m_sObjectPath				=	TEXT ("");
	m_nAPIError					=	0;
	m_nObjectType				=	SE_UNKNOWN_OBJECT_TYPE;
	m_siSecInfo					=	NULL;
	m_psdSD						=	NULL;
	m_paclDACL					=	NULL;
	m_paclSACL					=	NULL;
	m_psidOwner					=	NULL;
	m_psidGroup					=	NULL;
	m_fBufSDAlloc				=	false;
	m_fBufDACLAlloc			=	false;
	m_fBufSACLAlloc			=	false;
	m_fBufOwnerAlloc			=	false;
	m_fBufGroupAlloc			=	false;
}


//
// Destructor: clean up
//
CSD::~CSD ()
{
	Reset ();
}


//
// Clear the object
//
void CSD::Reset ()
{
	DeleteBufSD ();
	DeleteBufDACL ();
	DeleteBufSACL ();
	DeleteBufOwner ();
	DeleteBufGroup ();
}


//
// DeleteBufSD: Delete the buffer for the SD
//
void CSD::DeleteBufSD ()
{
	if (m_psdSD && m_fBufSDAlloc)
	{
		LocalFree (m_psdSD);
		m_psdSD			=	NULL;
		m_fBufSDAlloc	=	false;
	}
}


//
// DeleteBufDACL: Delete the buffer for the DACL
//
void CSD::DeleteBufDACL ()
{
	if (m_paclDACL && m_fBufDACLAlloc)
	{
		LocalFree (m_paclDACL);
		m_paclDACL			=	NULL;
		m_fBufDACLAlloc	=	false;
	}
}


//
// DeleteBufSACL: Delete the buffer for the SACL
//
void CSD::DeleteBufSACL ()
{
	if (m_paclSACL && m_fBufSACLAlloc)
	{
		LocalFree (m_paclSACL);
		m_paclSACL			=	NULL;
		m_fBufSACLAlloc	= false;
	}
}


//
// DeleteBufOwner: Delete the buffer for the Owner
//
void CSD::DeleteBufOwner ()
{
	if (m_psidOwner && m_fBufOwnerAlloc)
	{
		LocalFree (m_psidOwner);
		m_psidOwner			=	NULL;
		m_fBufOwnerAlloc	=	false;
	}
}


//
// DeleteBufGroup: Delete the buffer for the Group
//
void CSD::DeleteBufGroup ()
{
	if (m_psidGroup && m_fBufGroupAlloc)
	{
		LocalFree (m_psidGroup);
		m_psidGroup			=	NULL;
		m_fBufGroupAlloc	=	false;
	}
}


//
// GetSD: Get the SD and other information indicated by siSecInfo
//
DWORD CSD::GetSD (CString sObjectPath, SE_OBJECT_TYPE nObjectType, SECURITY_INFORMATION siSecInfo)
{
	bool									fSDRead			=	false;
	BOOL									fOK				=	true;
	DWORD									nError			=	RTN_OK;
	DWORD									nDesiredAccess	=	READ_CONTROL;
	HANDLE								hFile				=	NULL;
	HKEY									hKey				=	NULL;
	HANDLE								hAny				=	NULL;
	PSECURITY_DESCRIPTOR				pSDSelfRel		=	NULL;
	DWORD									nBufSD			=	0;
	DWORD									nBufDACL			=	0;
	DWORD									nBufSACL			=	0;
	DWORD									nBufOwner		=	0;
	DWORD									nBufGroup		=	0;

	// Check the parameters
	if (sObjectPath.IsEmpty ())
	{
		return RTN_ERR_PARAMS;
	}
	if (nObjectType == SE_UNKNOWN_OBJECT_TYPE)
	{
		return RTN_ERR_PARAMS;
	}

	m_sObjectPath	=	sObjectPath;
	m_nObjectType	=	nObjectType;
	m_siSecInfo		=	siSecInfo;

	if (siSecInfo & SACL_SECURITY_INFORMATION)
	{
		// Enable a privilege needed to read the SACL
		if (m_setaclMain->SetPrivilege (SE_SECURITY_NAME, true, true) != ERROR_SUCCESS)
		{
			return RTN_ERR_EN_PRIV;
		}

		// We need to have more access
		nDesiredAccess	|=	ACCESS_SYSTEM_SECURITY;
	}

	//
	// Use black magic (like backup programs) to read ALL SDs even if access is (normally) denied.
	//
	if (m_nObjectType	==	SE_FILE_OBJECT)
	{
		// Get a handle to the file/dir with special backup access (privilege SeBackupName has to be enabled for this).
		// This fails, of course, if the file is locked exclusively (like, for example, hiberfil.sys).
		hFile		=	CreateFile (m_sObjectPath, nDesiredAccess, 0, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT, NULL);

		if (hFile == INVALID_HANDLE_VALUE || hFile == NULL)
		{
			hFile		=	NULL;
			goto GetSD2ndTry;
		}
		else
		{
			hAny		=	hFile;
		}
	}
	else if (m_nObjectType == SE_REGISTRY_KEY)
	{
		// Open the key
		nError	=	m_setaclMain->RegKeyFixPathAndOpen (m_sObjectPath, hKey, false, READ_CONTROL);
		if (nError != RTN_OK)
		{
			hKey		=	NULL;
			goto GetSD2ndTry;
		}
		else
		{
			hAny		=	hKey;
		}
	}
	else if(m_nObjectType == SE_WMIGUID_OBJECT)
	{
		m_nAPIError	= AccessWMINamespaceSD (m_sObjectPath, &pSDSelfRel, 0);
		if (m_nAPIError != ERROR_SUCCESS)
			return RTN_ERR_GETSECINFO;
		else
			goto GotSD;
	}

	if (! hAny)
	{
		goto GetSD2ndTry;
	}

	// Get the size of the buffer needed
	GetKernelObjectSecurity (hAny, siSecInfo, &pSDSelfRel, 0, &nBufSD);
	if (! nBufSD)
	{
		goto GetSD2ndTry;
	}

	// Allocate a buffer for the SD.
	pSDSelfRel			=	LocalAlloc (LPTR, nBufSD);
	if (pSDSelfRel == NULL)
	{
		m_nAPIError	=	GetLastError ();

		return RTN_ERR_OUT_OF_MEMORY;
	}

	// Get the SD
	if (! GetKernelObjectSecurity (hAny, siSecInfo, pSDSelfRel, nBufSD, &nBufSD))
	{
		LocalFree (pSDSelfRel);
		pSDSelfRel		=	NULL;
		fSDRead			=	false;

		goto GetSD2ndTry;
	}
	else
	{
		fSDRead			=	true;
	}

GetSD2ndTry:

	// Close the open handles
	if (hFile)
	{
		CloseHandle (hFile);
		hFile			=	NULL;
	}
	if (hKey)
	{
		RegCloseKey (hKey);
		hKey			=	NULL;
	}

	if (! fSDRead)
	{
		// Above procedure did not work or this is an object type other than file system or registry: use the regular method
		m_nAPIError	=	GetNamedSecurityInfo (m_sObjectPath.GetBuffer (m_sObjectPath.GetLength () + 1), m_nObjectType, siSecInfo, NULL, NULL, NULL, NULL, &pSDSelfRel);
		m_sObjectPath.ReleaseBuffer ();

		if (m_nAPIError != ERROR_SUCCESS)
		{
			return RTN_ERR_GETSECINFO;
		}
	}

	// Does the object have an SD?
	if (pSDSelfRel == NULL)
	{
		// The SD is NULL -> no need to do further processing
		return RTN_OK;
	}

GotSD:
	//
	// We have got the SD.
	//
	// Convert the self-relative SD into an absolute SD:
	//
	// 1. Determine the size of the buffers needed
	nBufSD			=	0;

	MakeAbsoluteSD (pSDSelfRel, NULL, &nBufSD, NULL, &nBufDACL, NULL, &nBufSACL, NULL, &nBufOwner, NULL, &nBufGroup);
	if (GetLastError () != ERROR_INSUFFICIENT_BUFFER)
	{
		nError		=	RTN_ERR_CONVERT_SD;
		m_nAPIError	=	GetLastError ();

		return nError;
	}

	// 2.1 Delete old buffers
	DeleteBufSD ();
	DeleteBufDACL ();
	DeleteBufSACL ();
	DeleteBufOwner ();
	DeleteBufGroup ();

	// 2.2 Allocate new buffers
	m_psdSD					=	(PSECURITY_DESCRIPTOR) LocalAlloc (LPTR, nBufSD);
	m_fBufSDAlloc			=	true;

	if (nBufDACL)
	{
		m_paclDACL			=	(PACL) LocalAlloc (LPTR, nBufDACL);
		m_fBufDACLAlloc	=	true;
	}
	if (nBufSACL)
	{
		m_paclSACL			=	(PACL) LocalAlloc (LPTR, nBufSACL);
		m_fBufSACLAlloc	=	true;
	}
	if (nBufOwner)
	{
		m_psidOwner			=	(PSID) LocalAlloc (LPTR, nBufOwner);
		m_fBufOwnerAlloc	=	true;
	}
	if (nBufGroup)
	{
		m_psidGroup			=	(PSID) LocalAlloc (LPTR, nBufGroup);
		m_fBufGroupAlloc	=	true;
	}

	// 3. Do the conversion
	fOK	=	MakeAbsoluteSD (pSDSelfRel, m_psdSD, &nBufSD, m_paclDACL, &nBufDACL, m_paclSACL, &nBufSACL, m_psidOwner, &nBufOwner, m_psidGroup, &nBufGroup);

	// Free the temporary self-relative SD
	LocalFree (pSDSelfRel);
	pSDSelfRel				=	NULL;

	if (! fOK || ! IsValidSecurityDescriptor (m_psdSD))
	{
		nError				=	RTN_ERR_CONVERT_SD;
		m_nAPIError			=	GetLastError ();

		return nError;
	}

	return nError;
}


//
// SetSD: Set the SD and other information indicated by siSecInfo
//
DWORD CSD::SetSD (CString sObjectPath, SE_OBJECT_TYPE nObjectType, SECURITY_INFORMATION siSecInfo, 
						PACL paclDACL, PACL paclSACL, PSID psidOwner, PSID psidGroup, bool fLowLevel, bool fIsContainer)
{
	DWORD									nError			=	RTN_OK;
	HANDLE								hFile				=	NULL;
	HKEY									hKey				=	NULL;
	HANDLE								hAny				=	NULL;
	CString								sParentPath;
	CSD									csdSDParent(m_setaclMain);

	// Check the parameters
	if (sObjectPath.IsEmpty ())
	{
		return RTN_ERR_PARAMS;
	}
	if (nObjectType == SE_UNKNOWN_OBJECT_TYPE || nObjectType == SE_WMIGUID_OBJECT) 
	{
		return RTN_ERR_PARAMS;
	}

	if (siSecInfo & SACL_SECURITY_INFORMATION)
	{
		// Enable a privilege needed to read the SACL
		if (m_setaclMain->SetPrivilege (SE_SECURITY_NAME, true, true) != ERROR_SUCCESS)
		{
			return RTN_ERR_EN_PRIV;
		}
	}

	//
	// The code below should be used only after VERY thorough consideration:
	//
	// It tries to set the SD using the flag FILE_FLAG_BACKUP_SEMANTICS, which does not handle inheritance at all: 
	// permissions are not propagated to sub-objects, and the "inherited" flag of the object's ACEs  is not set/cleared 
	// as necessary. Since SetACL does not do this either, this normally results in incorrect ACLs.
	//
	// The BIG advantage of FILE_FLAG_BACKUP_SEMANTICS is, that it does not perform access checks but sets 
	// the SD on any object, regardless of it's current permissions and owner.
	//
	// There are three situations where we can safely use these low level writes: if only owner and/or 
	// primary group are to be set (no inheritance problems with those).
	//
	if (fLowLevel || 	
		(siSecInfo == OWNER_SECURITY_INFORMATION) ||
		(siSecInfo == GROUP_SECURITY_INFORMATION) ||
		(siSecInfo == (OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION)))
	{
		// Use black magic (like backup programs) to write ALL SDs even if access is (normally) denied.

		if (nObjectType == SE_FILE_OBJECT)
		{
			// Get a handle to the file/dir with special backup access (privilege SeBackupName has to be enabled for this)
			hFile		=	CreateFile (sObjectPath, WRITE_OWNER | WRITE_DAC, 0, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT, NULL);

			if (hFile == INVALID_HANDLE_VALUE || hFile == NULL)
			{
            nError = GetLastError();
				hFile		=	NULL;
				goto SetSD2ndTry;
			}
			else
			{
				hAny		=	hFile;
			}
		}
		else if (nObjectType == SE_REGISTRY_KEY)
		{
			// Open the key
			nError	=	m_setaclMain->RegKeyFixPathAndOpen (sObjectPath, hKey, false, WRITE_DAC);

			if (nError != RTN_OK)
			{
				hKey		=	NULL;
				goto SetSD2ndTry;
			}
			else
			{
				hAny		=	hKey;
			}
		}

		if (! hAny)
		{
			goto SetSD2ndTry;
		}

		// Set the SD using a handle instead of the name
		m_nAPIError	=	SetSecurityInfo (hAny, nObjectType, siSecInfo, psidOwner, psidGroup, paclDACL, paclSACL);
		if (m_nAPIError != ERROR_SUCCESS)
		{
			goto SetSD2ndTry;
		}
		else
		{
			nError	=	RTN_OK;
			goto CleanUp;
		}

	}	// if (fLowLevel)

SetSD2ndTry:

	// fLowLevel is set to false (the default), above procedure did not work, or this is an object type other than file system or registry: use the regular method to set the SD

	//
	// Convert all pseudo-inherited ACEs to correctly inherited ACEs
	//

	// Get the parent object's SD so we can determine pseudo-inheritance
	if (fIsContainer && (siSecInfo & DACL_SECURITY_INFORMATION || siSecInfo & SACL_SECURITY_INFORMATION))
	{
		// Get the path to the parent object
		if (GetParentObject (sObjectPath, nObjectType, sParentPath))
		{
			// Get the parent's SD
			nError			=	csdSDParent.GetSD (sParentPath, nObjectType, siSecInfo);
			m_nAPIError		=	csdSDParent.m_nAPIError;

			if (nError != RTN_OK || csdSDParent.m_psdSD == NULL)
			{
				// Clear the SD object
				csdSDParent.Reset ();
				// Reset error codes
				nError		=	RTN_OK;
				m_nAPIError	=	NO_ERROR;
			}
		}
	}

	if (! m_setaclMain->m_fRawMode)
	{
		// Convert pseudo-inherited ACEs to correctly inherited ACEs
		DWORD	nDummy	= 0;
		if (siSecInfo & DACL_SECURITY_INFORMATION && csdSDParent.m_paclDACL)
		{
			nError		=	ConvertPseudoInheritedACEsToInheritedACEs (paclDACL, fIsContainer, csdSDParent.m_paclDACL, nDummy);
			if (nError != RTN_OK)
			{
				goto CleanUp;
			}
		}
		if (siSecInfo & SACL_SECURITY_INFORMATION && csdSDParent.m_paclSACL)
		{
			nError		=	ConvertPseudoInheritedACEsToInheritedACEs (paclSACL, fIsContainer, csdSDParent.m_paclSACL, nDummy);
			if (nError != RTN_OK)
			{
				goto CleanUp;
			}
		}
	}

	//
	// Set the SD
	//
	m_nAPIError	=SetNamedSecurityInfo (sObjectPath.GetBuffer (sObjectPath.GetLength () + 1), nObjectType, siSecInfo, psidOwner, psidGroup, paclDACL, paclSACL);
	sObjectPath.ReleaseBuffer ();
	if (m_nAPIError != ERROR_SUCCESS)
	{
		nError	=	RTN_ERR_SETSECINFO;
	}
	else
	{
		nError	=	RTN_OK;
	}

CleanUp:

	// Close the open handles
	if (hFile)
	{
		CloseHandle (hFile);
		hFile			=	NULL;
	}
	if (hKey)
	{
		RegCloseKey (hKey);
		hKey			=	NULL;
	}

	return nError;
}


//
// Set a WMI SD
//
DWORD CSD::SetWmiSD (CString sObjectPath, SE_OBJECT_TYPE nObjectType, PACL paclDACL, PACL paclSACL, PSID psidOwner, PSID psidGroup)
{
	DWORD		nError			=	RTN_OK;

	// Check the parameters
	if (sObjectPath.IsEmpty ())
	{
		return RTN_ERR_PARAMS;
	}

	if (nObjectType != SE_WMIGUID_OBJECT) 
	{
		return RTN_ERR_PARAMS;
	}

	DWORD pSelfLen  = 0;

	if(psidOwner!=NULL)
	{
		nError = SetSecurityDescriptorOwner(m_psdSD, psidOwner, false);
	   if (nError == 0)
		{
		   m_nAPIError	=	GetLastError();
		   return RTN_ERR_SETSECINFO;
	   }
	}

	if(psidGroup!=NULL) {
       nError = SetSecurityDescriptorGroup(m_psdSD, psidGroup, false);
	   if (nError == 0) {
		   m_nAPIError	=	GetLastError();
		   return RTN_ERR_SETSECINFO;
	   }
	}

	if(paclDACL!=NULL) 
	{
        nError = SetSecurityDescriptorDacl(m_psdSD, true, paclDACL, false);
        if (nError == 0) {
	       m_nAPIError	=	GetLastError();
	       return RTN_ERR_SETSECINFO;
        }
	}

	if(paclSACL!=NULL) 
	{
         nError = SetSecurityDescriptorSacl(m_psdSD, true, paclSACL, false);
         if (nError == ERROR_SUCCESS) {
	        m_nAPIError	=	GetLastError();
	        return RTN_ERR_SETSECINFO;
         }
	}

	MakeSelfRelativeSD(m_psdSD, NULL,  &pSelfLen);

	PSECURITY_DESCRIPTOR pSelf = LocalAlloc(LPTR, pSelfLen);

	if(!MakeSelfRelativeSD(m_psdSD, pSelf,  &pSelfLen)) 
	{
		   m_nAPIError = GetLastError();
		   return RTN_ERR_CONVERT_SD;
	}

	nError = AccessWMINamespaceSD(sObjectPath, &pSelf, pSelfLen);
	if(nError != RTN_OK) {
		return nError;
	}

	return RTN_OK;
}



//////////////////////////////////////////////////////////////////////
//
// Global functions
//
//////////////////////////////////////////////////////////////////////


//
// Split: Called by various CSetACL functions. Emulation of Perl's split function.
//
bool Split (CString sDelimiter, CString sInput, CStringArray& asOutput)
{
	try
	{
		// Delete contents of the output array
		asOutput.RemoveAll ();

		// Remove leading and trailing whitespace
		sInput.TrimLeft ();
		sInput.TrimRight ();

		// Find each substring and add it to the output array
		int i = 0, j;
		while ((j = sInput.Find (sDelimiter, i)) != -1)
		{
			asOutput.Add (sInput.Mid (i, j - i));
			i = j + 1;
		}

		// Is there an element left at the end?
		if (sInput.GetLength () > i)
		{
			asOutput.Add (sInput.Mid (i, sInput.GetLength () - i));
		}

		return true;
	}
	catch (CMemoryException* exc)
	{
		exc->Delete ();

		return false;
	}
}


//
// CopySID: Copy a SID
//
PSID CopySID (PSID pSID)
{
	PSID		pSIDNew				=	NULL;

	if (! pSID)
	{
		return NULL;
	}

	// Allocate memory for the SID
	pSIDNew	= (PSID) LocalAlloc (LPTR, GetLengthSid (pSID));

	if (! pSIDNew)
	{
		return NULL;
	}

	// Copy the SID
	if (! CopySid (GetLengthSid (pSID), pSIDNew, pSID))
	{
		LocalFree (pSIDNew);
		return NULL;
	}

	return pSIDNew;
}


//
// BuildLongUnicodePath: Take any path any turn it into a path that is not limited to MAX_PATH, if possible
//
void BuildLongUnicodePath (CString& sPath)
{
	// Check this is not already a "long" path
	if (sPath.Left (3) == L"\\\\?")
	{
		return;
	}

	#ifdef UNICODE
		if (sPath.Left (2) == L"\\\\")
		{
			// This is a UNC path. Build a path in the following format: \\?\UNC\<server>\<share>
			sPath.Insert (2, L"?\\UNC\\");
		}
		else
		{
			// Check this is an absolute path
			if (sPath.Mid (1, 2) == L":\\")
			{
				// This is an absolute non-UNC. Build a path in the following format: \\?\<drive letter>:\<path>
				sPath.Insert (0, L"\\\\?\\");
			}
		}
	#endif
}


//
//	IsDirectory: Check if a given path points to an (existing) directory
//
bool IsDirectory (CString sPath)
{
	bool	fExists		= false;
	DWORD	nAttributes	= INVALID_FILE_ATTRIBUTES;
	
	nAttributes	= GetFileAttributes (sPath);
	if (nAttributes != INVALID_FILE_ATTRIBUTES)
	{
		if (nAttributes & FILE_ATTRIBUTE_DIRECTORY)
		{
			fExists	= true;
		}
	}

	return fExists;
}


//
//	GetParentObject: Determine the path of an object's parent
//
bool GetParentObject (CString sObject, SE_OBJECT_TYPE nObjectType, CString& sParent)
{
	bool	fParentPathIsDriveRoot	= false;

	// Empty the out parameter
	sParent.Empty ();

	// Check parameters
	if (sObject.IsEmpty ())
	{
		return false;
	}
	if (nObjectType != SE_FILE_OBJECT && nObjectType != SE_REGISTRY_KEY)
	{
		return false;
	}

	// Handle special cases
	if (sObject.GetLength() <= 2)
	{
		return false;
	}
	if (sObject.Mid(1,1) == ':' && sObject.GetLength() == 3)
	{
		// This is a path like: "c:\"
		return false;
	}

	// Trim trailing backslashes
	sObject.TrimRight ('\\');

	// Count the number of backslashes in the path
	int	nPos				= -1;
	DWORD nBackslashes	= 0;
	while ((nPos = sObject.Find ('\\', nPos + 1)) != -1)
	{
		nBackslashes++;
	}

	if (nObjectType == SE_FILE_OBJECT)
	{
		if (sObject.Left(8).CompareNoCase(L"\\\\?\\UNC\\") == 0)
		{
			// This is a long UNC path, e.g.: \\?\UNC\<server>\<share>\dir1\dir2\file

			// We can determine the parent only if it has at least 6 backslashes
			if (nBackslashes < 6)
			{
				return false;
			}
		}
		else if (sObject.Left(4) == L"\\\\?\\")
		{
			// This is a long drive letter path, e.g.: \\?\<drive letter>:\dir1\dir2\file

			// We can determine the parent only if it has at least 4 backslashes
			if (nBackslashes < 4)
			{
				return false;
			}
			if (nBackslashes == 4)
			{
				fParentPathIsDriveRoot	= true;
			}
		}
		else if (sObject.Left(2) == L"\\\\")
		{
			// This is a short UNC path, e.g.: \\server\share\dir1\dir2\file

			// We can determine the parent only if it has at least 4 backslashes
			if (nBackslashes < 4)
			{
				return false;
			}
		}
		else if (sObject.Mid(1,1) == ':')
		{
			// This is a short drive letter path, e.g.: c:\dir1\dir2\file

			// We can determine the parent only if it has at least 1 backslash
			if (nBackslashes < 1)
			{
				return false;
			}
			if (nBackslashes == 1)
			{
				fParentPathIsDriveRoot	= true;
			}
		}
		else
		{
			// This is a relative path

			// We can determine the parent only if it has at least 1 backslash
			if (nBackslashes < 1)
			{
				return false;
			}
		}
	}
	else if (nObjectType == SE_REGISTRY_KEY  || nObjectType == SE_WMIGUID_OBJECT)
	{
		// This is a registry or WMI path

		// We can determine the parent only if it has at least 1 backslash
		if (nBackslashes < 1)
		{
			return false;
		}
	}

	// With all the checks out of the way, let's determine the parent
	sParent		=	sObject.Left (sObject.ReverseFind('\\'));

	// If the parent path is the root of a drive, we need to append a backslash
	if (fParentPathIsDriveRoot)
	{
		sParent	+=	L"\\";
	}

	return true;
}

//
// Dummy used to get the module address in memory
//
void DummyFunction ()
{
	// Do something silly just to make sure the compiler does not optimize this away
	int i = 1;
	int j = i + 1;
	int k = i + j;
	i = k;
}


//
// Map a generic right to a set of standard and specific rights
//
DWORD MapGenericRight (DWORD nAccessMask, SE_OBJECT_TYPE nObjectType)
{
	// Generic mapping seem to be available for files and keys only??

	// These mappings are defined in GENERIC_MAPPING. Order: read, writer, execute, all

	DWORD	aGenericMappingFile[4]			= {0x120089, 0x120116, 0x1200A0, 0x1F01FF};
	// Note: this is the newer registry mapping introduced with Vista. Older OS mapped GENERIC_EXECUTE to 0x20019 just like GENERIC_READ.
	DWORD	aGenericMappingRegistry[4]		= {0x20019, 0x20006, 0x20039, 0xF003F};

	if (nObjectType == SE_FILE_OBJECT)
	{
		MapGenericMask (&nAccessMask, (GENERIC_MAPPING*) aGenericMappingFile);
	}
	else if (nObjectType == SE_REGISTRY_KEY)
	{
		MapGenericMask (&nAccessMask, (GENERIC_MAPPING*) aGenericMappingRegistry);
	}

	return nAccessMask;
}


