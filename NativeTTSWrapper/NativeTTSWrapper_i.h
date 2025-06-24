

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 03:14:07 2038
 */
/* Compiler settings for NativeTTSWrapper.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0628 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __NativeTTSWrapper_i_h__
#define __NativeTTSWrapper_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef DECLSPEC_XFGVIRT
#if defined(_CONTROL_FLOW_GUARD_XFG)
#define DECLSPEC_XFGVIRT(base, func) __declspec(xfg_virtual(base, func))
#else
#define DECLSPEC_XFGVIRT(base, func)
#endif
#endif

/* Forward Declarations */ 

#ifndef __INativeTTSWrapper_FWD_DEFINED__
#define __INativeTTSWrapper_FWD_DEFINED__
typedef interface INativeTTSWrapper INativeTTSWrapper;

#endif 	/* __INativeTTSWrapper_FWD_DEFINED__ */


#ifndef __CNativeTTSWrapper_FWD_DEFINED__
#define __CNativeTTSWrapper_FWD_DEFINED__

#ifdef __cplusplus
typedef class CNativeTTSWrapper CNativeTTSWrapper;
#else
typedef struct CNativeTTSWrapper CNativeTTSWrapper;
#endif /* __cplusplus */

#endif 	/* __CNativeTTSWrapper_FWD_DEFINED__ */


#ifndef __OpenSpeechSpVoice_FWD_DEFINED__
#define __OpenSpeechSpVoice_FWD_DEFINED__

#ifdef __cplusplus
typedef class OpenSpeechSpVoice OpenSpeechSpVoice;
#else
typedef struct OpenSpeechSpVoice OpenSpeechSpVoice;
#endif /* __cplusplus */

#endif 	/* __OpenSpeechSpVoice_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "sapi.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __INativeTTSWrapper_INTERFACE_DEFINED__
#define __INativeTTSWrapper_INTERFACE_DEFINED__

/* interface INativeTTSWrapper */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_INativeTTSWrapper;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E")
    INativeTTSWrapper : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct INativeTTSWrapperVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            INativeTTSWrapper * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            INativeTTSWrapper * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            INativeTTSWrapper * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            INativeTTSWrapper * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            INativeTTSWrapper * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            INativeTTSWrapper * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            INativeTTSWrapper * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } INativeTTSWrapperVtbl;

    interface INativeTTSWrapper
    {
        CONST_VTBL struct INativeTTSWrapperVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define INativeTTSWrapper_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define INativeTTSWrapper_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define INativeTTSWrapper_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define INativeTTSWrapper_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define INativeTTSWrapper_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define INativeTTSWrapper_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define INativeTTSWrapper_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __INativeTTSWrapper_INTERFACE_DEFINED__ */



#ifndef __NativeTTSWrapperLib_LIBRARY_DEFINED__
#define __NativeTTSWrapperLib_LIBRARY_DEFINED__

/* library NativeTTSWrapperLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_NativeTTSWrapperLib;

EXTERN_C const CLSID CLSID_CNativeTTSWrapper;

#ifdef __cplusplus

class DECLSPEC_UUID("E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B")
CNativeTTSWrapper;
#endif

EXTERN_C const CLSID CLSID_OpenSpeechSpVoice;

#ifdef __cplusplus

class DECLSPEC_UUID("F2E8B6A1-3C4D-4E5F-8A7B-9C1D2E3F4A5B")
OpenSpeechSpVoice;
#endif
#endif /* __NativeTTSWrapperLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


