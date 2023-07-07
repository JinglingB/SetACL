//////////////////////////////////////////////////////////////////////
//
// Definitions that must come before anything else, normally in stdafx.h
//
//////////////////////////////////////////////////////////////////////


#pragma once


// Windows >= XP needed
#define	_WIN32_WINNT				0x05010000


// Do not include rarely used parts of the Windows headers
#define VC_EXTRALEAN


// We ONLY use unicode
#ifndef UNICODE
#define UNICODE
#endif

#ifndef _UNICODE
#define _UNICODE
#endif

//////////////////////////////////////////////////////////////////////
//
// Includes
//
//////////////////////////////////////////////////////////////////////


// Includes normally in stdafx.h
#ifdef I_AM_AN_OCX
	#include <afxctl.h>
	#include <afxext.h>
#else
	#include <afxwin.h>
#endif

// Always needed for certain "advanced" MFC classes
#include <afxtempl.h>

// More includes
#include <windows.h>
#include <iostream>
#include <aclapi.h>
#include <lmerr.h>
#include <winspool.h>
#include <winsvc.h>
#include <lmaccess.h>
#include <lmapibuf.h>
#include <lmdfs.h>
#include <sddl.h>
#include <dsgetdc.h>
#include <dsrole.h>
#include <objbase.h>
#include <stdio.h>

#include <wbemidl.h> 
#include <comdef.h>
#include <strsafe.h>

#include <new>

