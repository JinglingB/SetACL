'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
' This is a sample script that demonstrates how to use the SetACL ActiveX control from Visual Basic Script (VBS).
' This script is included in the download package of ActiveX version of SetACL.
' 
' Author: Helge Klein
' Tested on: Windows 7 RTM
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
' Constants
' 
' Since VBScript cannot import type libraries, we need to define all constants again here
' Note: these were copied from the original source code file SetACL.ODL and modified to fit the VBS const syntax
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ACL/SD information types
' Enum SDINFO
' The ACL is a DACL (permission information)
const ACL_DACL = 1
' The ACL is a SACL (auditing information)
const ACL_SACL = 2
' Owner information
const SD_OWNER = 4
' Primary group information
const SD_GROUP = 8

' Actions the program can perform
' Enum ACTIONS
' Add an ACE
const ACTN_ADDACE = 1
' List the entries in the security descriptor
const ACTN_LIST = 2
' Set the owner
const ACTN_SETOWNER = 4
' Set the primary group
const ACTN_SETGROUP = 8
' Clear the DACL of any non-inherited ACEs
const ACTN_CLEARDACL = 16
' Clear the SACL of any non-inherited ACEs
const ACTN_CLEARSACL = 32
' Set the flag 'allow inheritable permissions from the parent object to propagate to this object'
const ACTN_SETINHFROMPAR = 64
' Reset permissions on all sub-objects and enable propagation of inherited permissions
const ACTN_RESETCHILDPERMS = 128
' Replace one trustee by another in all ACEs
const ACTN_REPLACETRUSTEE = 256
' Remove all ACEs belonging to a certain trustee
const ACTN_REMOVETRUSTEE = 512
' Copy the permissions for one trustee to another
const ACTN_COPYTRUSTEE = 1024
' Replace one domain by another in all ACEs
const ACTN_REPLACEDOMAIN = 256
' Remove all ACEs belonging to a certain domain
const ACTN_REMOVEDOMAIN = 512
' Copy the permissions for one domain to another
const ACTN_COPYDOMAIN = 1024
' Restore entire security descriptors backup up with the list function
const ACTN_RESTORE = 2048
' Process all trustee actions
const ACTN_TRUSTEE = 4096
' Process all domain actions
const ACTN_DOMAIN = 8192

' Return/Error codes
' Enum RETCODES
' OK
const RTN_OK = 0
' Usage instructions were printed
const RTN_USAGE = 1
' General error
const RTN_ERR_GENERAL = 2
' Parameter(s) incorrect
const RTN_ERR_PARAMS = 3
' The object was not set
const RTN_ERR_OBJECT_NOT_SET = 4
' The call to GetNamedSecurityInfo () failed
const RTN_ERR_GETSECINFO = 5
' The SID for a trustee could not be found
const RTN_ERR_LOOKUP_SID = 6
' Directory permissions specified are invalid
const RTN_ERR_INV_DIR_PERMS = 7
' Printer permissions specified are invalid
const RTN_ERR_INV_PRN_PERMS = 8
' Registry permissions specified are invalid
const RTN_ERR_INV_REG_PERMS = 9
' Service permissions specified are invalid
const RTN_ERR_INV_SVC_PERMS = 10
' Share permissions specified are invalid
const RTN_ERR_INV_SHR_PERMS = 11
' A privilege could not be enabled
const RTN_ERR_EN_PRIV = 12
' A privilege could not be disabled
const RTN_ERR_DIS_PRIV = 13
' No notification function was given
const RTN_ERR_NO_NOTIFY = 14
' An error occured in the list function
const RTN_ERR_LIST_FAIL = 15
' FindFile reported an error
const RTN_ERR_FINDFILE = 16
' GetSecurityDescriptorControl () failed
const RTN_ERR_GET_SD_CONTROL = 17
' An internal program error occured
const RTN_ERR_INTERNAL = 18
' SetEntriesInAcl () failed
const RTN_ERR_SETENTRIESINACL = 19
' A registry path is incorrect
const RTN_ERR_REG_PATH = 20
' Connect to a remote registry failed
const RTN_ERR_REG_CONNECT = 21
' Opening a registry key failed
const RTN_ERR_REG_OPEN = 22
' Enumeration of registry keys failed
const RTN_ERR_REG_ENUM = 23
' Preparation failed
const RTN_ERR_PREPARE = 24
' The call to SetNamedSecurityInfo () failed
const RTN_ERR_SETSECINFO = 25
' Incorrect list options specified
const RTN_ERR_LIST_OPTIONS = 26
' A SD could not be converted to/from string format
const RTN_ERR_CONVERT_SD = 27
' ACL listing failed
const RTN_ERR_LIST_ACL = 28
' Looping through an ACL failed
const RTN_ERR_LOOP_ACL = 29
' Deleting an ACE failed
const RTN_ERR_DEL_ACE = 30
' Copying an ACL failed
const RTN_ERR_COPY_ACL = 31
' Adding an ACE failed
const RTN_ERR_ADD_ACE = 32
' No backup/restore file was specified
const RTN_ERR_NO_LOGFILE = 33
' The backup/restore file could not be opened
const RTN_ERR_OPEN_LOGFILE = 34
' A read operation from the backup/restore file failed
const RTN_ERR_READ_LOGFILE = 35
' A write operation from the backup/restore file failed
const RTN_ERR_WRITE_LOGFILE = 36
' The operating system is not supported
const RTN_ERR_OS_NOT_SUPPORTED = 37
' The security descriptor is invalid
const RTN_ERR_INVALID_SD = 38
' The call to SetSecurityDescriptorDacl () failed
const RTN_ERR_SET_SD_DACL = 39
' The call to SetSecurityDescriptorSacl () failed
const RTN_ERR_SET_SD_SACL = 40
' The call to SetSecurityDescriptorOwner () failed
const RTN_ERR_SET_SD_OWNER = 41
' The call to SetSecurityDescriptorGroup () failed
const RTN_ERR_SET_SD_GROUP = 42
' The domain specified is invalid
const RTN_ERR_INV_DOMAIN = 43
' An error occured, but it was ignored
const RTN_ERR_IGNORED = 44
' The creation of an SD failed
const RTN_ERR_CREATE_SD = 45
' Memory allocation failed
const RTN_ERR_OUT_OF_MEMORY = 46

' For inheritance from the parent
' Enum INHERITANCE
' Do not change settings
const INHPARNOCHANGE = 0
' Inherit from parent
const INHPARYES = 1
' Do not inherit, copy inheritable permissions
const INHPARCOPY = 2
' Do not inherit, do not copy inheritable permissions
const INHPARNOCOPY = 4

' For propagation to child objects
' Enum PROPAGATION
' The specific access permissions will only be applied to the container, and will not be inherited by objects created within the container.
const NO_INHERITANCE = 0
' The specific access permissions will only be inherited by objects created within the specific container. The access permissions will not be applied to the container itself.
const SUB_OBJECTS_ONLY_INHERIT = 1
' The specific access permissions will be inherited by containers created within the specific container, will be applied to objects created within the container, but will not be applied to the container itself.
const SUB_CONTAINERS_ONLY_INHERIT = 2
' Combination of SUB_OBJECTS_ONLY_INHERIT and SUB_CONTAINERS_ONLY_INHERIT.
const SUB_CONTAINERS_AND_OBJECTS_INHERIT = 3
' Do not propogate permissions, only the direct descendent gets permissions.
const INHERIT_NO_PROPAGATE = 4
' The specific access permissions will not affect the object they are set on but its children only (depending on other propagation flags).
const INHERIT_ONLY = 8

' List formats
' Enum LISTFORMATS
' SDDL format
const LIST_SDDL = 0
' CSV format
const LIST_CSV = 1
' Tabular format
const LIST_TAB = 2

' How to list trustees
' Enum LISTNAMES
' List names
const LIST_NAME = 1
' List SIDs
const LIST_SID = 2
' List names and SIDs
const LIST_NAME_SID = 3

' For recursion
' Enum RECURSION
' No recursion
const RECURSE_NO = 1
' Recurse containers
const RECURSE_CONT = 2
' Recurse objects
const RECURSE_OBJ = 4
' Recurse containers and objects
const RECURSE_CONT_OBJ = 6

' Specify the type of object to process
' Enum SE_OBJECT_TYPE
' Files/directories
const SE_FILE_OBJECT = 1
' Services
const SE_SERVICE = 2
' Printers
const SE_PRINTER = 3
' Registry keys
const SE_REGISTRY_KEY = 4
' Network shares
const SE_LMSHARE = 5

' Specify what type of ACE to modify
' Enum ACCESS_MODE
' Grant access (add an ACE)
const GRANT_ACCESS = 1
' Set access (replace other ACEs for the trustee specified by another ACE)
const SET_ACCESS = 2
' Deny access (add an access denied ACE)
const DENY_ACCESS = 3
' Revoke access
const REVOKE_ACCESS = 4
' Audit success
const SET_AUDIT_SUCCESS = 5
' Audit failure
const SET_AUDIT_FAILURE = 6

' Is a trustee (user/group) specified by SID or by name?
' Trustee is name (e.g. Administrators)
const TRUSTEE_IS_NAME = 0
' Trustee is SID (e.g. S-1-5-...)
const TRUSTEE_IS_SID = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
' Script body
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Create the SetACL object. Events are mapped to subroutines with the prefix "SetACL_" (see http://msdn.microsoft.com/en-us/library/ms974564.aspx)
set objSetACL = WScript.CreateObject("SETACL.SetACLCtrl.1", "SetACL_")

' Tell SetACL which object to process (path and object type)
' In this example, process d:\temp
retCode = objSetACL.SetObject ("D:\temp", SE_FILE_OBJECT)

' Error checking
if retCode <> RTN_OK then
	PrintSetACLError "SetACL SetObject", retCode
	WScript.Quit (retCode)
end if

' Tell SetACL which action(s) to perform
' In this example, set the owner
retCode = objSetACL.SetAction (ACTN_SETOWNER)

' Error checking
if retCode <> RTN_OK then
	PrintSetACLError "SetACL SetAction", retCode
	WScript.Quit (retCode)
end if

' Set parameters for the action(s) we configured earlier
' In this example, set the owner to the group "Administratoren" (which is a name, not a SID)
retCode = objSetACL.SetOwner ("Administratoren", TRUSTEE_IS_NAME)

' Error checking
if retCode <> RTN_OK then
	PrintSetACLError "SetACL SetOwner", retCode
	WScript.Quit (retCode)
end if

' Configure recursion
' In this example, recurse from the top directory given (d:\temp) and process every directory and file in the tree
' In other words set the owner of every object you find in the file system below d:\temp
retCode = objSetACL.SetRecursion (RECURSE_CONT_OBJ)

' Error checking
if retCode <> RTN_OK then
	PrintSetACLError "SetACL SetRecursion", retCode
	WScript.Quit (retCode)
end if

' Status message
WScript.Echo "Starting SetACL..."

' Tell SetACL to start working
retCode = objSetACL.Run

' Status message
if retCode = RTN_OK then
	WScript.Echo "SetACL finished successfully!"
else
	PrintSetACLError "SetACL run", retCode
end if

' End of script
WScript.Quit (retCode)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
' Functions / subroutines
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Print an error message if we get a return code indicating failure
sub PrintSetACLError (WhatFailed, retCode)
	WScript.Echo WhatFailed & " failed: " & objSetACL.GetResourceString(retCode) & vbCrLf & "OS error: " & objSetACL.GetLastAPIErrorMessage ()
end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
' Event handlers
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Catch and print messages from the SetACL ActiveX control which are passed (fired) as events
' The name postfix of this function (MessageEvent) must be identical to the event name as defined by SetACL
' The prefix (SetACL_) can be set freely in the call to WScript.CreateObject
sub SetACL_MessageEvent (Message)
	WScript.Echo Message
end sub