// **************************************************************************
//
// Copyright (c) Microsoft Corporation, All Rights Reserved
//
// File:  Security.cpp 
//
// Description:
//        WMI Client Security Sample.  This sample shows how various how to 
// handle various DCOM security issues.  In particular, it shows how to call
// CoSetProxyBlanket in order to deal with common situations.  The sample also
// shows how to access the security descriptors that wmi uses to control
// namespace access. 
// **************************************************************************

#include "Stdafx.h"
#include "CSetACL.h"

#pragma comment (lib, "WbemUuid.Lib")
 

//***************************************************************************
//
// SetProxySecurity
//
// Purpose: Calls CoSetProxyBlanket in order to control the security on a 
// particular interface proxy.
//
//***************************************************************************

HRESULT SetProxySecurity(IUnknown * pProxy)
{
    HRESULT hr;
    DWORD dwAuthnSvc, dwAuthzSvc ,dwAuthnLevel, dwImpLevel, dwCapabilities;

    //  There are various reasons to set security.  An application that can 
    //  call CoInitializeSecurity and doesnt use an alternative user\password,
    //  need not bother with call CoSetProxyBlanket.  There are at least
    //  three cases that do need it though.
    //
    //  1) Dlls cannot call CoInitializeSecurity and will need to call it just
    //     to raise the impersonation level.  This does NOT require that the 
    //     RPC_AUTH_IDENTITY_HANDLE to be set nor does this require setting
    //     it on the IUnknown pointer.
    //  2) Any time that an alternative user\password are set as is the case
    //     in this simple sample.  Note that it is necessary in that case to 
    //     also set the information on the IUnknown.
    //  3) If the caller has a thread token from a remote call to itself and
    //     it wants to use that identity.  That would be the case of an RPC/COM
    //     server which is handling a call from one of its clients and it wants
    //     to use the client's identity in calls to WMI.  In this case, then 
    //     the RPC_AUTH_IDENTITY_HANDLE does not need to be set, but the 
    //     dwCapabilities arguement should be set for cloaking.  This is also
    //     required for the IUnknown pointer.
    

    // In this case, nothing needs to be done.  But for the sake of 
    // example the code will set the impersonation level.  

    // For backwards compatibility, retrieve the previous authentication level, so
    // we can echo-back the value when we set the blanket

	hr = CoQueryProxyBlanket(
				pProxy,           //Location for the proxy to query
				&dwAuthnSvc,      //Location for the current authentication service
				&dwAuthzSvc,      //Location for the current authorization service
				NULL,             //Location for the current principal name
				&dwAuthnLevel,    //Location for the current authentication level
				&dwImpLevel,      //Location for the current impersonation level
				NULL,             //Location for the value passed to IClientSecurity::SetBlanket
				&dwCapabilities   //Location for flags indicating further capabilities of the proxy
			);
	if(FAILED(hr))
		return hr;

	hr = CoSetProxyBlanket(pProxy, RPC_C_AUTHN_DEFAULT, RPC_C_AUTHZ_DEFAULT, COLE_DEFAULT_PRINCIPAL, dwAuthnLevel, RPC_C_IMP_LEVEL_IMPERSONATE, COLE_DEFAULT_AUTHINFO, EOAC_DEFAULT);

	return hr;
}

//***************************************************************************
//
// StoreSD
//
// Purpose: Writes back the ACL which controls access to the namespace.
//
//***************************************************************************


bool StoreSD(IWbemServices * pSession, PSECURITY_DESCRIPTOR pSD)
{
    bool bRet = false;
    HRESULT hr;

	if (!IsValidSecurityDescriptor(pSD))
		return false;

    // Get the class object

    IWbemClassObject * pClass = NULL;
    _bstr_t InstPath(L"__systemsecurity=@");
    _bstr_t ClassPath(L"__systemsecurity");
    hr = pSession->GetObject(ClassPath, 0, NULL, &pClass, NULL);
    if(FAILED(hr))
        return false;
	if (pClass == NULL)
		return false;

    // Get the input parameter class

    _bstr_t MethName(L"SetSD");
    IWbemClassObject * pInClassSig = NULL;
    hr = pClass->GetMethod(MethName,0, &pInClassSig, NULL);
    pClass->Release();
    if(FAILED(hr))
        return false;

    // spawn an instance of the input parameter class

    IWbemClassObject * pInArg = NULL;
    pInClassSig->SpawnInstance(0, &pInArg);
    pInClassSig->Release();
    if(FAILED(hr))
        return false;


    // move the SD into a variant.

    SAFEARRAY FAR* psa;
    SAFEARRAYBOUND rgsabound[1];    rgsabound[0].lLbound = 0;
    long lSize = GetSecurityDescriptorLength(pSD);
    rgsabound[0].cElements = lSize;
    psa = SafeArrayCreate( VT_UI1, 1 , rgsabound );
    if(psa == NULL)
    {
        pInArg->Release();
        return false;
    }

    char * pData = NULL;
    hr = SafeArrayAccessData(psa, (void HUGEP* FAR*)&pData);
    if(FAILED(hr))
    {
        pInArg->Release();
        return false;
    }

    memcpy(pData, pSD, lSize);

    SafeArrayUnaccessData(psa);
    _variant_t var;
    var.vt = VT_I4|VT_ARRAY;
    var.parray = psa;

    // put the property

    hr = pInArg->Put(L"SD" , 0, &var, 0);      
    if(FAILED(hr))
    {
        pInArg->Release();
        return false;
    }

    // Execute the method
    hr = pSession->ExecMethod(InstPath,
            MethName,
            0,
            NULL, pInArg,
            NULL, NULL);
    if(FAILED(hr))
        printf("\nPut failed, returned 0x%x",hr);

    return bRet;
}

//***************************************************************************
//
// ReadACL
//
// Purpose: Reads the ACL which controls access to the namespace.
//
//***************************************************************************

DWORD ReadACL(IWbemServices *pNamespace, PSECURITY_DESCRIPTOR* pSD, DWORD pLen)
{
	UNUSED_ALWAYS(pLen);

	_bstr_t InstPath(L"__systemsecurity=@");
	_bstr_t MethName(L"GetSD");
	IWbemClassObject * pOutParams = NULL;

	// The security descriptor is returned via the GetSD method

	HRESULT hr = pNamespace->ExecMethod(InstPath, MethName, 0, NULL, NULL, &pOutParams, NULL);
	if(FAILED(hr))
		return hr;
	if (pOutParams == NULL)
		return ERROR_FUNCTION_FAILED;

	// The output parameters has a property which has the descriptor

	_bstr_t prop(L"sd");
	_variant_t var;
	hr = pOutParams->Get(prop, 0, &var, NULL, NULL);
	if(FAILED(hr))
		return hr;
	if(var.vt != (VT_ARRAY | VT_UI1))
		return ERROR_INVALID_DATA;

	SAFEARRAY * psa = var.parray;
	PSECURITY_DESCRIPTOR tmpSD;
	long lBound, uBound;

	// Get size
	hr = SafeArrayGetLBound(psa, 1, &lBound);
	if(FAILED(hr))
		return hr;
	hr = SafeArrayGetUBound(psa, 1, &uBound);
	if(FAILED(hr))
		return hr;

	int nBufSD = (int)(uBound - lBound + 1);

	hr = SafeArrayAccessData(psa, (void HUGEP* FAR*)&tmpSD);
	if(FAILED(hr))
		return hr;

	// Allocate a buffer for the SD.
	*pSD = LocalAlloc (LPTR, nBufSD);
	if (*pSD == NULL)
		return ERROR_NOT_ENOUGH_MEMORY;

	memcpy(*pSD, tmpSD, nBufSD);

	SafeArrayUnaccessData(psa);

	return ERROR_SUCCESS;
}

//***************************************************************************
//
// AccessWMINamespaceSD
//
// Purpose: Main entry point.
//
//***************************************************************************

DWORD AccessWMINamespaceSD(CString sObjectPath, PSECURITY_DESCRIPTOR* pSD, DWORD pSdLen)
{
	bool bComAlreadyInitialized = false;

	// Make sure we have a UNC path
	if (sObjectPath.Left(2) != L"\\\\")
		sObjectPath = L"\\\\.\\" + sObjectPath;

	IWbemLocator *pLocator = NULL;
	IWbemServices *pNamespace = NULL;

	// Initialize COM
	HRESULT hr = CoInitialize(0);
	if (FAILED(hr))
		if (hr != RPC_E_CHANGED_MODE)
			return hr;
		else
			bComAlreadyInitialized = true;

	// This sets the default impersonation level to "Impersonate" which is what WMI 
	// providers will generally require. 
	//
	// DLLs cannot call this function. To bump up the impersonation level, they must 
	// call CoSetProxyBlanket which is illustrated later on.  
	//  
	// When using asynchronous WMI API's remotely in an environment where the "Local System" account 
	// has no network identity (such as non-Kerberos domains), the authentication level of 
	// RPC_C_AUTHN_LEVEL_NONE is needed. However, lowering the authentication level to 
	// RPC_C_AUTHN_LEVEL_NONE makes your application less secure. It is wise to
	//	use semi-synchronous API's for accessing WMI data and events instead of the asynchronous ones.
    
	hr = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_IMPERSONATE, NULL,
	   								EOAC_SECURE_REFS, //change to EOAC_NONE if you change dwAuthnLevel to RPC_C_AUTHN_LEVEL_NONE
										0);
	if (FAILED(hr) && hr != RPC_E_TOO_LATE)
	{
		// RPC_E_TOO_LATE means that Security has already been initialized. Ignore that.
		return hr;
	}
                                
	hr = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *) &pLocator);
	if (FAILED(hr))
		return hr;
   
	BSTR path = sObjectPath.AllocSysString();
	hr = pLocator->ConnectServer(path, NULL, NULL, NULL, 0, NULL, NULL, &pNamespace);
	SysFreeString(path);
	if(FAILED(hr))
		return hr;

	hr = SetProxySecurity(pNamespace);
	if(!SUCCEEDED(hr))
	{
		pNamespace->Release();
		return hr;
	}

	if(pSdLen == 0)
	{
		hr = ReadACL(pNamespace, pSD, pSdLen);
		if (hr != ERROR_SUCCESS)
			return hr;
	}
	else 
	{
		StoreSD(pNamespace, *pSD);
	}

	pNamespace->Release();

	if (! bComAlreadyInitialized)
		CoUninitialize();

	 return ERROR_SUCCESS;
}

