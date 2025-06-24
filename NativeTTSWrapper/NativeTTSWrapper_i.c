

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 03:14:07 2038
 */
/* Compiler settings for NativeTTSWrapper.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0628 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        EXTERN_C __declspec(selectany) const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif // !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_INativeTTSWrapper,0xA74D7C8E,0x4CC5,0x4F2F,0xA6,0xEB,0x80,0x4D,0xEE,0x18,0x50,0x0E);


MIDL_DEFINE_GUID(IID, LIBID_NativeTTSWrapperLib,0xB8F4A8E2,0x9C3D,0x4A5E,0x8F,0x7C,0x2D,0x1B,0x3E,0x4F,0x5A,0x6B);


MIDL_DEFINE_GUID(CLSID, CLSID_CNativeTTSWrapper,0xE1C4A8F2,0x9B3D,0x4A5E,0x8F,0x7C,0x2D,0x1B,0x3E,0x4F,0x5A,0x6B);


MIDL_DEFINE_GUID(CLSID, CLSID_OpenSpeechSpVoice,0xF2E8B6A1,0x3C4D,0x4E5F,0x8A,0x7B,0x9C,0x1D,0x2E,0x3F,0x4A,0x5B);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



