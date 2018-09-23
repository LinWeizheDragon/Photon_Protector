

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 6.00.0361 */
/* at Tue Oct 21 22:48:30 2008
 */
/* Compiler settings for .\ProcProtectCtrl.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __ProcProtectCtrl_h__
#define __ProcProtectCtrl_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IProcProtect_FWD_DEFINED__
#define __IProcProtect_FWD_DEFINED__
typedef interface IProcProtect IProcProtect;
#endif 	/* __IProcProtect_FWD_DEFINED__ */


#ifndef __ProcProtect_FWD_DEFINED__
#define __ProcProtect_FWD_DEFINED__

#ifdef __cplusplus
typedef class ProcProtect ProcProtect;
#else
typedef struct ProcProtect ProcProtect;
#endif /* __cplusplus */

#endif 	/* __ProcProtect_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 

#ifndef __IProcProtect_INTERFACE_DEFINED__
#define __IProcProtect_INTERFACE_DEFINED__

/* interface IProcProtect */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IProcProtect;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("277546D6-AC56-439D-8273-EA8F9B6946D2")
    IProcProtect : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Protect( 
            /* [in] */ LONG lProcId,
            /* [in] */ BYTE bIsProtect,
            /* [out] */ DWORD *pdwResult) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Register( 
            /* [in] */ CHAR *pszRegStr) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IProcProtectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IProcProtect * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IProcProtect * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IProcProtect * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IProcProtect * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IProcProtect * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IProcProtect * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IProcProtect * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Protect )( 
            IProcProtect * This,
            /* [in] */ LONG lProcId,
            /* [in] */ BYTE bIsProtect,
            /* [out] */ DWORD *pdwResult);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Register )( 
            IProcProtect * This,
            /* [in] */ CHAR *pszRegStr);
        
        END_INTERFACE
    } IProcProtectVtbl;

    interface IProcProtect
    {
        CONST_VTBL struct IProcProtectVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IProcProtect_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IProcProtect_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IProcProtect_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IProcProtect_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IProcProtect_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IProcProtect_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IProcProtect_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IProcProtect_Protect(This,lProcId,bIsProtect,pdwResult)	\
    (This)->lpVtbl -> Protect(This,lProcId,bIsProtect,pdwResult)

#define IProcProtect_Register(This,pszRegStr)	\
    (This)->lpVtbl -> Register(This,pszRegStr)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProcProtect_Protect_Proxy( 
    IProcProtect * This,
    /* [in] */ LONG lProcId,
    /* [in] */ BYTE bIsProtect,
    /* [out] */ DWORD *pdwResult);


void __RPC_STUB IProcProtect_Protect_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IProcProtect_Register_Proxy( 
    IProcProtect * This,
    /* [in] */ CHAR *pszRegStr);


void __RPC_STUB IProcProtect_Register_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IProcProtect_INTERFACE_DEFINED__ */



#ifndef __ProcProtectCtrlLib_LIBRARY_DEFINED__
#define __ProcProtectCtrlLib_LIBRARY_DEFINED__

/* library ProcProtectCtrlLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ProcProtectCtrlLib;

EXTERN_C const CLSID CLSID_ProcProtect;

#ifdef __cplusplus

class DECLSPEC_UUID("2F6816C3-620C-4EA5-B5A2-FE5F979BEB95")
ProcProtect;
#endif
#endif /* __ProcProtectCtrlLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


