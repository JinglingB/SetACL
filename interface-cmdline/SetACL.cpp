/////////////////////////////////////////////////////////////////////////////
//
//
//	SetACL.cpp
//
//
//	Description:	Command line interface for SetACL
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
#include	"SetACL.h"


//////////////////////////////////////////////////////////////////////
//
// Defines
//
//////////////////////////////////////////////////////////////////////


#define	TXTPRINTHELP	TEXT("\nType 'SetACL -help' for help.\n")


//////////////////////////////////////////////////////////////////////
//
// Global variables
//
//////////////////////////////////////////////////////////////////////


// The application object
CWinApp	theApp;

// Are we in silent mode?
bool		fSilent;


//////////////////////////////////////////////////////////////////////
//
// Functions
//
//////////////////////////////////////////////////////////////////////


//
// main
//
int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	UNUSED_ALWAYS(envp);		// suppress warning (unreferenced formal parameter)

	// Change locale from minimum ("C") to the user-default ANSI code page obtained from the operating system 
	// (e.g. "German_Germany.1252") (see http://msdn.microsoft.com/en-us/library/x99tb11d%28v=VS.100%29.aspx)
	_tsetlocale(LC_ALL, _T(""));
	
	// Our SetACL class object
	CSetACL	oSetACL (PrintMsg);

	// Return code to pass to caller
	int		nRetCode	= 0;

	// Other variable initializations
	fSilent	=	false;


	// Initialize MFC
	if (!AfxWinInit (::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
	{
		printf ("ERROR: MFC initialization failed!\n");

		return RTN_ERR_GENERAL;
	}

	// Process the command line
	if ((nRetCode = ProcessCmdLine (argc, argv, PrintMsg, &oSetACL)) != RTN_OK)
	{
		return nRetCode;
	}

	// Run the program
	nRetCode	=	oSetACL.Run ();

	// Return status
	return nRetCode;
}


//
// ProcessCmdLine: Process the command line
//
DWORD ProcessCmdLine (int argc, TCHAR* argv[], void (* funcNotify) (CString), CSetACL* oSetACL)
{
	// Check arguments
	if (! oSetACL || ! argv)
	{
		return RTN_ERR_GENERAL;
	}

	int		nRetCode				=	0;

	CString	sObjectName;
	DWORD		nObjectType			=	0;
	bool		fActionClear		=	false;
	bool		fActionSetProt		=	false;
	bool		fActionReset		=	false;
	DWORD		nRecursionType		=	RECURSE_NO;
	bool		fDACLReset			=	false;
	bool		fSACLReset			=	false;
	bool		fClearDACL			=	false;
	bool		fClearSACL			=	false;
	DWORD		nDACLProtected		=	INHPARNOCHANGE;
	DWORD		nSACLProtected		=	INHPARNOCHANGE;
	CString	sArg;
	CString	sParam;


	// Invocation without arguments -> message and quit
	if (argc	==	1)
	{
		if (funcNotify)
		{
			(*funcNotify) (TXTPRINTHELP);
		}
		return RTN_ERR_PARAMS;
	}

	// Should we print the help page?
	if (argc > 1)
	{
		sArg			=	argv[1];

		if (sArg.CompareNoCase (TEXT ("-help")) == 0 || sArg.CompareNoCase (TEXT ("-?")) == 0)
		{
			PrintHelp ();

			return RTN_USAGE;
		}
	}

	//
	// Loop over all arguments
	//
	for (int i = 1; i < argc; i++)
	{
		sArg			=	argv[i];

		// Is this an argument that must be followed by a parameter?
		if (sArg.CompareNoCase (TEXT ("-silent")) == 0)
		{
			fSilent	=	true;
		}
		else if (sArg.CompareNoCase (TEXT ("-ignoreerr")) == 0)
		{
			oSetACL->SetIgnoreErrors (true);
		}
		else if (sArg.CompareNoCase (TEXT ("-raw")) == 0)
		{
			oSetACL->SetRawMode (true);
		}
		else
		{
			if (i == argc - 1)
			{
				if (funcNotify)
				{
					(*funcNotify) (TEXT ("ERROR in command line: No parameter found for option ") + sArg + TEXT ("!\n") + TXTPRINTHELP);
				}
				return RTN_ERR_PARAMS;
			}
			else
			{
				sParam	=	argv[i + 1];

				// Check if the parameter contains quotes (which is probably an unintentional escape, e.g.: "C:\" NextParam -> <C:" Next Param>
				if (sParam.Find ('"') != -1)
				{
					if (funcNotify)
					{
						CString sMessage;
						sMessage.Format (L"WARNING: The parameter <%s> contains a double quotation mark (\"). Did you unintentionally escape a double quote? Hint: use <\"C:\\\\\"> instead of <\"C:\\\">.", (LPCTSTR) sParam);
						(*funcNotify) (sMessage);
					}
				}
			}

			// Check if the following argument begins with a '-'. If yes -> error
			if (sParam.Left (1) == TEXT ("-"))
			{
				if (funcNotify)
				{
					(*funcNotify) (TEXT ("ERROR in command line: No parameter found for option ") + sArg + TEXT ("!\n") + TXTPRINTHELP);
				}
				return RTN_ERR_PARAMS;
			}

			// Do not process the parameter during the next loop - this is done now
			i++;

			//
			// Process the parameter
			//
			if (sArg.CompareNoCase (TEXT ("-on")) == 0)
			{
				sObjectName		=	sParam;

				if (nObjectType)
				{
					if (oSetACL->SetObject (sObjectName, (SE_OBJECT_TYPE) nObjectType) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR while processing command line: object (name, type): ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-ot")) == 0)
			{
				if (sParam.CompareNoCase (TEXT ("file")) == 0)
				{
					nObjectType	=	SE_FILE_OBJECT;
				}
				else if (sParam.CompareNoCase (TEXT ("reg")) == 0)
				{
					nObjectType	=	SE_REGISTRY_KEY;
				}
				else if (sParam.CompareNoCase (TEXT ("srv")) == 0)
				{
					nObjectType	=	SE_SERVICE;
				}
				else if (sParam.CompareNoCase (TEXT ("prn")) == 0)
				{
					nObjectType	=	SE_PRINTER;
				}
				else if (sParam.CompareNoCase (TEXT ("shr")) == 0)
				{
					nObjectType	=	SE_LMSHARE;
				}
				else if(sParam.CompareNoCase (TEXT ("wmi")) == 0) 
				{
					nObjectType	=	SE_WMIGUID_OBJECT;
				}
				else
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid object type specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				if (! sObjectName.IsEmpty ())
				{
					if (oSetACL->SetObject (sObjectName, (SE_OBJECT_TYPE) nObjectType) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR while processing command line: object (name, type): ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-actn")) == 0)
			{
				if (sParam.CompareNoCase (TEXT ("ace")) == 0)
				{
					if (oSetACL->AddAction (ACTN_ADDACE) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else if (sParam.CompareNoCase (TEXT ("trustee")) == 0)
				{
					if (oSetACL->AddAction (ACTN_TRUSTEE) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else if (sParam.CompareNoCase (TEXT ("domain")) == 0)
				{
					if (oSetACL->AddAction (ACTN_DOMAIN) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else if (sParam.CompareNoCase (TEXT ("list")) == 0)
				{
					if (oSetACL->AddAction (ACTN_LIST) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else if (sParam.CompareNoCase (TEXT ("restore")) == 0)
				{
					if (funcNotify)
					{
						(*funcNotify) (L"INFO: The restore action ignores the object name parameter (paths are read from the backup file). However, other actions that require the object name may be combined with -restore.");
					}
					if (oSetACL->AddAction (ACTN_RESTORE) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else if (sParam.CompareNoCase (TEXT ("setowner")) == 0)
				{
					if (oSetACL->AddAction (ACTN_SETOWNER) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else if (sParam.CompareNoCase (TEXT ("setgroup")) == 0)
				{
					if (oSetACL->AddAction (ACTN_SETGROUP) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: ") + sParam + TEXT (" could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else if (sParam.CompareNoCase (TEXT ("clear")) == 0)
				{
					fActionClear	=	true;

					if (fClearDACL)
					{
						if (oSetACL->AddAction (ACTN_CLEARDACL) != RTN_OK)
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: CLEARDACL could not be set!\n"));
							}
							return RTN_ERR_PARAMS;
						}
					}

					if (fClearSACL)
					{
						if (oSetACL->AddAction (ACTN_CLEARSACL) != RTN_OK)
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: CLEARSACL could not be set!\n"));
							}
							return RTN_ERR_PARAMS;
						}
					}
				}
				else if (sParam.CompareNoCase (TEXT ("setprot")) == 0)
				{
					fActionSetProt	=	true;

					if (oSetACL->AddAction (ACTN_SETINHFROMPAR) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: SETINHFROMPAR could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else if (sParam.CompareNoCase (TEXT ("rstchldrn")) == 0)
				{
					fActionReset	=	true;

					if (oSetACL->AddAction (ACTN_RESETCHILDPERMS) != RTN_OK)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: RESETCHILDPERMS could not be set!\n"));
						}
						return RTN_ERR_PARAMS;
					}
				}
				else
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid action specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-rec")) == 0)
			{
				if (sParam.CompareNoCase (TEXT ("no")) == 0)
				{
					nRecursionType	=	RECURSE_NO;
				}
				else if (sParam.CompareNoCase (TEXT ("cont")) == 0)
				{
					nRecursionType	=	RECURSE_CONT;
				}
				else if (sParam.CompareNoCase (TEXT ("yes")) == 0)
				{
					nRecursionType	=	RECURSE_CONT;
				}
				else if (sParam.CompareNoCase (TEXT ("obj")) == 0)
				{
					nRecursionType	=	RECURSE_OBJ;
				}
				else if (sParam.CompareNoCase (TEXT ("cont_obj")) == 0)
				{
					nRecursionType	=	RECURSE_CONT_OBJ;
				}
				else
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid recursion type specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				if (oSetACL->SetRecursion (nRecursionType) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR (internal) while processing command line: recursion type could not be set!\n"));
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-rst")) == 0)
			{
				if (sParam.CompareNoCase (TEXT ("dacl")) == 0)
				{
					fDACLReset		=	true;
				}
				else if (sParam.CompareNoCase (TEXT ("sacl")) == 0)
				{
					fSACLReset		=	true;
				}
				else if (sParam.CompareNoCase (TEXT ("dacl,sacl")) == 0 || sParam.CompareNoCase (TEXT ("sacl,dacl")) == 0)
				{
					fDACLReset		=	true;
					fSACLReset		=	true;
				}
				else
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -rst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Now set the object flags
				if (oSetACL->SetObjectFlags (nDACLProtected, nSACLProtected, fDACLReset, fSACLReset) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Backup/Restore file: ") + sParam + TEXT (" could not be set!\n"));
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-bckp")) == 0)
			{
				if (oSetACL->SetBackupRestoreFile (sParam) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Backup/Restore file: ") + sParam + TEXT (" could not be set!\n"));
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-log")) == 0)
			{
				oSetACL->SetLogFile (sParam);
			}
			else if (sArg.CompareNoCase (TEXT ("-fltr")) == 0)
			{
				oSetACL->AddObjectFilter (sParam);
			}
			else if (sArg.CompareNoCase (TEXT ("-clr")) == 0)
			{
				if (sParam.CompareNoCase (TEXT ("dacl")) == 0)
				{
					fClearDACL		=	true;

					if (fActionClear)
					{
						if (oSetACL->AddAction (ACTN_CLEARDACL) != RTN_OK)
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: CLEARDACL could not be set!\n"));
							}
							return RTN_ERR_PARAMS;
						}
					}
				}
				else if (sParam.CompareNoCase (TEXT ("sacl")) == 0)
				{
					fClearSACL		=	true;

					if (fActionClear)
					{
						if (oSetACL->AddAction (ACTN_CLEARSACL) != RTN_OK)
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: CLEARSACL could not be set!\n"));
							}
							return RTN_ERR_PARAMS;
						}
					}
				}
				else if (sParam.CompareNoCase (TEXT ("dacl,sacl")) == 0 || sParam.CompareNoCase (TEXT ("sacl,dacl")) == 0)
				{
					fClearDACL		=	true;
					fClearSACL		=	true;

					if (fActionClear)
					{
						if (oSetACL->AddAction (ACTN_CLEARDACL) != RTN_OK)
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: CLEARDACL could not be set!\n"));
							}
							return RTN_ERR_PARAMS;
						}

						if (oSetACL->AddAction (ACTN_CLEARSACL) != RTN_OK)
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR (internal) while processing command line: Action: CLEARSACL could not be set!\n"));
							}
							return RTN_ERR_PARAMS;
						}
					}
				}
				else
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -clr specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-ace")) == 0)
			{
				CStringArray	asElements;
				CString			sTrusteeName;
				CString			sPermission;
				bool				fSID				=	false;
				DWORD				nInheritance	=	0;
				DWORD				nAccessMode		=	SET_ACCESS;
				DWORD				nACLType			=	ACL_DACL;

				// Split the parameter into its components
				if (! Split (TEXT (";"), sParam, asElements))
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Check the number of entries in the parameter
				if (asElements.GetSize () < 2 || asElements.GetSize () > 6)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid number of entries in parameter for option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Loop through the entries in the parameter
				for (int j = 0; j < asElements.GetSize (); j++)
				{
					CStringArray	asEntry;

					// Split the entry into value and data
					if (! Split (TEXT (":"), asElements[j], asEntry))
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry.GetSize () != 2)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry[0].CompareNoCase (TEXT ("n")) == 0)
					{
						sTrusteeName	=	asEntry[1];
					}
					else if (asEntry[0].CompareNoCase (TEXT ("p")) == 0)
					{
						sPermission	=	asEntry[1];
					}
					else if (asEntry[0].CompareNoCase (TEXT ("s")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("y")) == 0)
						{
							fSID			=	true;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("n")) == 0)
						{
							fSID			=	false;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid SID entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("i")) == 0)
					{
						CStringArray	asInheritance;

						// Split the inheritance list
						if (! Split (TEXT (","), asEntry[1], asInheritance))
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid inheritance entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}

						// Loop through the inheritance list
						for (int k = 0; k < asInheritance.GetSize (); k++)
						{
							if (asInheritance[k].CompareNoCase (TEXT ("so")) == 0)
							{
								nInheritance	|=	SUB_OBJECTS_ONLY_INHERIT;
							}
							else if (asInheritance[k].CompareNoCase (TEXT ("sc")) == 0)
							{
								nInheritance	|=	SUB_CONTAINERS_ONLY_INHERIT;
							}
							else if (asInheritance[k].CompareNoCase (TEXT ("np")) == 0)
							{
								nInheritance	|=	INHERIT_NO_PROPAGATE;
							}
							else if (asInheritance[k].CompareNoCase (TEXT ("io")) == 0)
							{
								nInheritance	|=	INHERIT_ONLY;
							}
							else
							{
								if (funcNotify)
								{
									(*funcNotify) (TEXT ("ERROR in command line: Invalid inheritance entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
								}
								return RTN_ERR_PARAMS;
							}
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("m")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("set")) == 0)
						{
							nAccessMode			=	SET_ACCESS;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("grant")) == 0)
						{
							nAccessMode			=	GRANT_ACCESS;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("deny")) == 0)
						{
							nAccessMode			=	DENY_ACCESS;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("revoke")) == 0)
						{
							nAccessMode			=	REVOKE_ACCESS;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("aud_succ")) == 0)
						{
							nAccessMode			=	SET_AUDIT_SUCCESS;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("aud_fail")) == 0)
						{
							nAccessMode			=	SET_AUDIT_FAILURE;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("aud_fail,aud_succ")) == 0 || asEntry[1].CompareNoCase (TEXT ("aud_succ,aud_fail")) == 0)
						{
							nAccessMode			=	SET_AUDIT_FAILURE + SET_AUDIT_SUCCESS;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid access mode entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("w")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("dacl")) == 0)
						{
							nACLType			=	ACL_DACL;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("sacl")) == 0)
						{
							nACLType			=	ACL_SACL;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid ACL type (where) entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}
				}

				// Now add the ace
				if (oSetACL->AddACE (sTrusteeName, fSID, sPermission, nInheritance, nInheritance != 0, nAccessMode, nACLType) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR while processing command line: ACE: ") + sParam + TEXT (" could not be set!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-trst")) == 0)
			{
				CStringArray	asElements;
				CString			sTrusteeName1;
				CString			sTrusteeName2;
				bool				fSID1					=	false;
				bool				fSID2					=	false;
				bool				fDACL					=	false;
				bool				fSACL					=	false;
				DWORD				nTrusteeAction		=	0;

				// Split the parameter into its components
				if (! Split (TEXT (";"), sParam, asElements))
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Check the number of entries in the parameter
				if (asElements.GetSize () < 3 || asElements.GetSize () > 6)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid number of entries in parameter for option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Loop through the entries in the parameter
				for (int j = 0; j < asElements.GetSize (); j++)
				{
					CStringArray	asEntry;

					// Split the entry into value and data
					if (! Split (TEXT (":"), asElements[j], asEntry))
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry.GetSize () != 2)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry[0].CompareNoCase (TEXT ("n1")) == 0)
					{
						sTrusteeName1	=	asEntry[1];
					}
					else if (asEntry[0].CompareNoCase (TEXT ("n2")) == 0)
					{
						sTrusteeName2	=	asEntry[1];
					}
					else if (asEntry[0].CompareNoCase (TEXT ("s1")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("y")) == 0)
						{
							fSID1			=	true;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("n")) == 0)
						{
							fSID1			=	false;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid SID entry ") + asElements[j] + TEXT (" in a parameter option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("s2")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("y")) == 0)
						{
							fSID2			=	true;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("n")) == 0)
						{
							fSID2			=	false;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid SID entry ") + asElements[j] + TEXT (" in a parameter option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("ta")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("remtrst")) == 0)
						{
							nTrusteeAction		=	ACTN_REMOVETRUSTEE;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("repltrst")) == 0)
						{
							nTrusteeAction		=	ACTN_REPLACETRUSTEE;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("cpytrst")) == 0)
						{
							nTrusteeAction		=	ACTN_COPYTRUSTEE;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid trustee action entry ") + asElements[j] + TEXT (" in a parameter option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("w")) == 0)
					{
						CStringArray	asACLType;

						// Split the ACL type list
						if (! Split (TEXT (","), asEntry[1], asACLType))
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid ACL type (where) entry ") + asElements[j] + TEXT (" in a parameter option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}

						// Loop through the ACL type list
						for (int k = 0; k < asACLType.GetSize (); k++)
						{
							if (asACLType[k].CompareNoCase (TEXT ("dacl")) == 0)
							{
								fDACL				=	true;
							}
							else if (asACLType[k].CompareNoCase (TEXT ("sacl")) == 0)
							{
								fSACL				=	true;
							}
							else
							{
								if (funcNotify)
								{
									(*funcNotify) (TEXT ("ERROR in command line: Invalid ACL type (where) entry ") + asElements[j] + TEXT (" in a parameter option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
								}
								return RTN_ERR_PARAMS;
							}
						}
					}
					else
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -trst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}
				}

				// Now add the trustee
				if (oSetACL->AddTrustee (sTrusteeName1, sTrusteeName2, fSID1, fSID2, nTrusteeAction, fDACL, fSACL) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR while processing command line: Trustee: ") + sParam + TEXT (" could not be set!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-dom")) == 0)
			{
				CStringArray	asElements;
				CString			sDomainName1;
				CString			sDomainName2;
				bool				fDACL					=	false;
				bool				fSACL					=	false;
				DWORD				nDomainAction		=	0;

				// Split the parameter into its components
				if (! Split (TEXT (";"), sParam, asElements))
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -dom specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Check the number of entries in the parameter
				if (asElements.GetSize () < 3 || asElements.GetSize () > 4)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid number of entries in parameter for option -dom specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Loop through the entries in the parameter
				for (int j = 0; j < asElements.GetSize (); j++)
				{
					CStringArray	asEntry;

					// Split the entry into value and data
					if (! Split (TEXT (":"), asElements[j], asEntry))
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -dom specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry.GetSize () != 2)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -dom specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry[0].CompareNoCase (TEXT ("n1")) == 0)
					{
						sDomainName1	=	asEntry[1];
					}
					else if (asEntry[0].CompareNoCase (TEXT ("n2")) == 0)
					{
						sDomainName2	=	asEntry[1];
					}
					else if (asEntry[0].CompareNoCase (TEXT ("da")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("remdom")) == 0)
						{
							nDomainAction		=	ACTN_REMOVEDOMAIN;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("repldom")) == 0)
						{
							nDomainAction		=	ACTN_REPLACEDOMAIN;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("cpydom")) == 0)
						{
							nDomainAction		=	ACTN_COPYDOMAIN;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid domain action entry ") + asElements[j] + TEXT (" in a parameter option -dom specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("w")) == 0)
					{
						CStringArray	asACLType;

						// Split the ACL type list
						if (! Split (TEXT (","), asEntry[1], asACLType))
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid ACL type (where) entry ") + asElements[j] + TEXT (" in a parameter option -dom specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}

						// Loop through the ACL type list
						for (int k = 0; k < asACLType.GetSize (); k++)
						{
							if (asACLType[k].CompareNoCase (TEXT ("dacl")) == 0)
							{
								fDACL				=	true;
							}
							else if (asACLType[k].CompareNoCase (TEXT ("sacl")) == 0)
							{
								fSACL				=	true;
							}
							else
							{
								if (funcNotify)
								{
									(*funcNotify) (TEXT ("ERROR in command line: Invalid ACL type (where) entry ") + asElements[j] + TEXT (" in a parameter option -dom specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
								}
								return RTN_ERR_PARAMS;
							}
						}
					}
					else
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -dom specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}
				}

				// Now add the domain
				if (oSetACL->AddDomain (sDomainName1, sDomainName2, nDomainAction, fDACL, fSACL) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR while processing command line: Domain: ") + sParam + TEXT (" could not be set!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-ownr")) == 0)
			{
				CStringArray	asElements;
				CString			sTrusteeName;
				bool				fSID					=	false;

				// Split the parameter into its components
				if (! Split (TEXT (";"), sParam, asElements))
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -ownr specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Check the number of entries in the parameter
				if (asElements.GetSize () < 1 || asElements.GetSize () > 2)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid number of entries in parameter for option -ownr specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Loop through the entries in the parameter
				for (int j = 0; j < asElements.GetSize (); j++)
				{
					CStringArray	asEntry;

					// Split the entry into value and data
					if (! Split (TEXT (":"), asElements[j], asEntry))
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -ownr specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry.GetSize () != 2)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -ownr specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry[0].CompareNoCase (TEXT ("n")) == 0)
					{
						sTrusteeName	=	asEntry[1];
					}
					else if (asEntry[0].CompareNoCase (TEXT ("s")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("y")) == 0)
						{
							fSID			=	true;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("n")) == 0)
						{
							fSID			=	false;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid SID entry ") + asElements[j] + TEXT (" in a parameter option -ownr specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -ownr specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}
				}

				// Now set the owner
				if (oSetACL->SetOwner (sTrusteeName, fSID) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR while processing command line: Owner: ") + sParam + TEXT (" could not be set!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-grp")) == 0)
			{
				CStringArray	asElements;
				CString			sTrusteeName;
				bool				fSID					=	false;

				// Split the parameter into its components
				if (! Split (TEXT (";"), sParam, asElements))
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -grp specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Check the number of entries in the parameter
				if (asElements.GetSize () < 1 || asElements.GetSize () > 2)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid number of entries in parameter for option -grp specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Loop through the entries in the parameter
				for (int j = 0; j < asElements.GetSize (); j++)
				{
					CStringArray	asEntry;

					// Split the entry into value and data
					if (! Split (TEXT (":"), asElements[j], asEntry))
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -grp specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry.GetSize () != 2)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -grp specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry[0].CompareNoCase (TEXT ("n")) == 0)
					{
						sTrusteeName	=	asEntry[1];
					}
					else if (asEntry[0].CompareNoCase (TEXT ("s")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("y")) == 0)
						{
							fSID			=	true;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("n")) == 0)
						{
							fSID			=	false;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid SID entry ") + asElements[j] + TEXT (" in a parameter option -grp specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -grp specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}
				}

				// Now set the primary group
				if (oSetACL->SetPrimaryGroup (sTrusteeName, fSID) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR while processing command line: primary group: ") + sParam + TEXT (" could not be set!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-op")) == 0)
			{
				CStringArray	asElements;

				// Split the parameter into its components
				if (! Split (TEXT (";"), sParam, asElements))
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -op specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Check the number of entries in the parameter
				if (asElements.GetSize () < 1 || asElements.GetSize () > 2)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid number of entries in parameter for option -op specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Loop through the entries in the parameter
				for (int j = 0; j < asElements.GetSize (); j++)
				{
					CStringArray	asEntry;

					// Split the entry into value and data
					if (! Split (TEXT (":"), asElements[j], asEntry))
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -op specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry.GetSize () != 2)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -op specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry[0].CompareNoCase (TEXT ("dacl")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("nc")) == 0)
						{
							nDACLProtected	=	INHPARNOCHANGE;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("np")) == 0)
						{
							nDACLProtected	=	INHPARYES;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("p_c")) == 0)
						{
							nDACLProtected	=	INHPARCOPY;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("p_nc")) == 0)
						{
							nDACLProtected	=	INHPARNOCOPY;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid protection entry ") + asElements[j] + TEXT (" in a parameter option -op specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("sacl")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("nc")) == 0)
						{
							nSACLProtected	=	INHPARNOCHANGE;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("np")) == 0)
						{
							nSACLProtected	=	INHPARYES;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("p_c")) == 0)
						{
							nSACLProtected	=	INHPARCOPY;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("p_nc")) == 0)
						{
							nSACLProtected	=	INHPARNOCOPY;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid protection entry ") + asElements[j] + TEXT (" in a parameter option -op specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -op specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}
				}

				// Now set the object flags
				if (oSetACL->SetObjectFlags (nDACLProtected, nSACLProtected, fDACLReset, fSACLReset) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR (internal) while processing command line: object flags: ") + sParam + TEXT (" could not be set!\n"));
					}
					return RTN_ERR_PARAMS;
				}
			}
			else if (sArg.CompareNoCase (TEXT ("-lst")) == 0)
			{
				CStringArray	asElements;
				DWORD				nListFormat		=	LIST_CSV;
				DWORD				nListWhat		=	ACL_DACL;
				DWORD				nListNameSID	=	LIST_NAME;
				bool				fListInherited	=	false;

				// Split the parameter into its components
				if (! Split (TEXT (";"), sParam, asElements))
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR in command line: Invalid parameter for option -lst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
					}
					return RTN_ERR_PARAMS;
				}

				// Loop through the entries in the parameter
				for (int j = 0; j < asElements.GetSize (); j++)
				{
					CStringArray	asEntry;

					// Split the entry into value and data
					if (! Split (TEXT (":"), asElements[j], asEntry))
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -lst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry.GetSize () != 2)
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -lst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}

					if (asEntry[0].CompareNoCase (TEXT ("f")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("sddl")) == 0)
						{
							nListFormat		=	LIST_SDDL;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("csv")) == 0 || asEntry[1].CompareNoCase (TEXT ("own")) == 0)
						{
							nListFormat		=	LIST_CSV;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("tab")) == 0)
						{
							nListFormat		=	LIST_TAB;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid list format entry ") + asElements[j] + TEXT (" in a parameter option -lst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("i")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("y")) == 0)
						{
							fListInherited	=	true;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("n")) == 0)
						{
							fListInherited	=	false;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid list format entry ") + asElements[j] + TEXT (" in a parameter option -lst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("s")) == 0)
					{
						if (asEntry[1].CompareNoCase (TEXT ("y")) == 0)
						{
							nListNameSID	=	LIST_SID;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("n")) == 0)
						{
							nListNameSID	=	LIST_NAME;
						}
						else if (asEntry[1].CompareNoCase (TEXT ("b")) == 0)
						{
							nListNameSID	=	LIST_NAME_SID;
						}
						else
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid list format entry ") + asElements[j] + TEXT (" in a parameter option -lst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}
					}
					else if (asEntry[0].CompareNoCase (TEXT ("w")) == 0)
					{
						CStringArray	asListWhat;

						// Split the list
						if (! Split (TEXT (","), asEntry[1], asListWhat))
						{
							if (funcNotify)
							{
								(*funcNotify) (TEXT ("ERROR in command line: Invalid list what entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
							}
							return RTN_ERR_PARAMS;
						}

						// Reset the "what" variable
						nListWhat	= 0;

						// Loop through the list
						for (int k = 0; k < asListWhat.GetSize (); k++)
						{
							if (asListWhat[k].CompareNoCase (TEXT ("d")) == 0)
							{
								nListWhat		|=	ACL_DACL;
							}
							else if (asListWhat[k].CompareNoCase (TEXT ("s")) == 0)
							{
								nListWhat		|=	ACL_SACL;
							}
							else if (asListWhat[k].CompareNoCase (TEXT ("o")) == 0)
							{
								nListWhat		|=	SD_OWNER;
							}
							else if (asListWhat[k].CompareNoCase (TEXT ("g")) == 0)
							{
								nListWhat		|=	SD_GROUP;
							}
							else
							{
								if (funcNotify)
								{
									(*funcNotify) (TEXT ("ERROR in command line: Invalid list what entry ") + asElements[j] + TEXT (" in a parameter option -ace specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
								}
								return RTN_ERR_PARAMS;
							}
						}
					}
					else
					{
						if (funcNotify)
						{
							(*funcNotify) (TEXT ("ERROR in command line: Invalid entry ") + asElements[j] + TEXT (" in a parameter option -lst specified: ") + sParam + TEXT ("!\n") + TXTPRINTHELP);
						}
						return RTN_ERR_PARAMS;
					}
				}

				// Now set the list options
				if (oSetACL->SetListOptions (nListFormat, nListWhat, fListInherited, nListNameSID, false) != RTN_OK)
				{
					if (funcNotify)
					{
						(*funcNotify) (TEXT ("ERROR (internal) while processing command line: list options: ") + sParam + TEXT (" could not be set!\n"));
					}
					return RTN_ERR_PARAMS;
				}
			}
			else
			{
				if (funcNotify)
				{
					(*funcNotify) (TEXT ("ERROR in command line: Invalid option specified: ") + sArg + TEXT ("!\n") + TXTPRINTHELP);
				}
				return RTN_ERR_PARAMS;
			}
		}
	}	//	for (int i = 1; i < argc; i++)

	return nRetCode;
}


//
// PrintMsg: Called by various CSetACL functions. Prints status text to stdout and/or log file
//
void PrintMsg (CString sMessage)
{
	if (! fSilent)
	{
		// Print to the screen
		_tprintf (TEXT ("%s\n"), sMessage.GetString ());
	}
}


//
// PrintHelp: Print help text
//
void PrintHelp ()
{
	printf ("\nSetACL by Helge Klein\n");
	printf ("\n");
	printf ("Homepage:        http://helgeklein.com/\n");
	printf ("Version:         2.3.2.0\n");
	printf ("Copyright:       Helge Klein\n");
	printf ("License:         LGPL\n");
	printf ("\n");
	printf ("Syntax:\n");
	printf ("=======\n");
	printf ("\n");
	printf ("SetACL -on ObjectName -ot ObjectType -actn Action1 ParametersForAction1\n");
	printf ("       [-actn Action2 ParametersForAction2] [Options]\n");
	printf ("\n");
	printf ("Documentation:\n");
	printf ("==============\n");
	printf ("\n");
	printf ("Documentation and examples are maintained at\n");
	printf ("http://helgeklein.com/.\n");
	printf ("The usage reference can be found at\n");
	printf ("http://helgeklein.com/setacl/documentation/command-line-version-setacl-exe/\n");
	printf ("\n");
}