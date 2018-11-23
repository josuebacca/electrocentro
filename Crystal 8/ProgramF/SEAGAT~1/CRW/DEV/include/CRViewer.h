/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Tue Feb 08 01:13:31 2000
 */
/* Compiler settings for ReportAXViewer.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
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

#ifndef __ReportAXViewer_h__
#define __ReportAXViewer_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __ICRVTrackCursorInfo_FWD_DEFINED__
#define __ICRVTrackCursorInfo_FWD_DEFINED__
typedef interface ICRVTrackCursorInfo ICRVTrackCursorInfo;
#endif 	/* __ICRVTrackCursorInfo_FWD_DEFINED__ */


#ifndef __ICRViewer_FWD_DEFINED__
#define __ICRViewer_FWD_DEFINED__
typedef interface ICRViewer ICRViewer;
#endif 	/* __ICRViewer_FWD_DEFINED__ */


#ifndef __ICRViewer2_FWD_DEFINED__
#define __ICRViewer2_FWD_DEFINED__
typedef interface ICRViewer2 ICRViewer2;
#endif 	/* __ICRViewer2_FWD_DEFINED__ */


#ifndef __ICrystalReportViewer_FWD_DEFINED__
#define __ICrystalReportViewer_FWD_DEFINED__
typedef interface ICrystalReportViewer ICrystalReportViewer;
#endif 	/* __ICrystalReportViewer_FWD_DEFINED__ */


#ifndef __ICrystalReportViewer2_FWD_DEFINED__
#define __ICrystalReportViewer2_FWD_DEFINED__
typedef interface ICrystalReportViewer2 ICrystalReportViewer2;
#endif 	/* __ICrystalReportViewer2_FWD_DEFINED__ */


#ifndef __ICrystalReportViewer3_FWD_DEFINED__
#define __ICrystalReportViewer3_FWD_DEFINED__
typedef interface ICrystalReportViewer3 ICrystalReportViewer3;
#endif 	/* __ICrystalReportViewer3_FWD_DEFINED__ */


#ifndef __ICRVEventInfo_FWD_DEFINED__
#define __ICRVEventInfo_FWD_DEFINED__
typedef interface ICRVEventInfo ICRVEventInfo;
#endif 	/* __ICRVEventInfo_FWD_DEFINED__ */


#ifndef __ICrystalReportSourceRouter_FWD_DEFINED__
#define __ICrystalReportSourceRouter_FWD_DEFINED__
typedef interface ICrystalReportSourceRouter ICrystalReportSourceRouter;
#endif 	/* __ICrystalReportSourceRouter_FWD_DEFINED__ */


#ifndef __ICRFields_FWD_DEFINED__
#define __ICRFields_FWD_DEFINED__
typedef interface ICRFields ICRFields;
#endif 	/* __ICRFields_FWD_DEFINED__ */


#ifndef __ICRField_FWD_DEFINED__
#define __ICRField_FWD_DEFINED__
typedef interface ICRField ICRField;
#endif 	/* __ICRField_FWD_DEFINED__ */


#ifndef __IEnumCRFields_FWD_DEFINED__
#define __IEnumCRFields_FWD_DEFINED__
typedef interface IEnumCRFields IEnumCRFields;
#endif 	/* __IEnumCRFields_FWD_DEFINED__ */


#ifndef __ICollectCRFields_FWD_DEFINED__
#define __ICollectCRFields_FWD_DEFINED__
typedef interface ICollectCRFields ICollectCRFields;
#endif 	/* __ICollectCRFields_FWD_DEFINED__ */


#ifndef ___ICRViewerEvents_FWD_DEFINED__
#define ___ICRViewerEvents_FWD_DEFINED__
typedef interface _ICRViewerEvents _ICRViewerEvents;
#endif 	/* ___ICRViewerEvents_FWD_DEFINED__ */


#ifndef __CRViewer_FWD_DEFINED__
#define __CRViewer_FWD_DEFINED__

#ifdef __cplusplus
typedef class CRViewer CRViewer;
#else
typedef struct CRViewer CRViewer;
#endif /* __cplusplus */

#endif 	/* __CRViewer_FWD_DEFINED__ */


#ifndef __CRVTrackCursorInfo_FWD_DEFINED__
#define __CRVTrackCursorInfo_FWD_DEFINED__

#ifdef __cplusplus
typedef class CRVTrackCursorInfo CRVTrackCursorInfo;
#else
typedef struct CRVTrackCursorInfo CRVTrackCursorInfo;
#endif /* __cplusplus */

#endif 	/* __CRVTrackCursorInfo_FWD_DEFINED__ */


#ifndef __CRFields_FWD_DEFINED__
#define __CRFields_FWD_DEFINED__

#ifdef __cplusplus
typedef class CRFields CRFields;
#else
typedef struct CRFields CRFields;
#endif /* __cplusplus */

#endif 	/* __CRFields_FWD_DEFINED__ */


#ifndef __CRField_FWD_DEFINED__
#define __CRField_FWD_DEFINED__

#ifdef __cplusplus
typedef class CRField CRField;
#else
typedef struct CRField CRField;
#endif /* __cplusplus */

#endif 	/* __CRField_FWD_DEFINED__ */


#ifndef __CRVEventInfo_FWD_DEFINED__
#define __CRVEventInfo_FWD_DEFINED__

#ifdef __cplusplus
typedef class CRVEventInfo CRVEventInfo;
#else
typedef struct CRVEventInfo CRVEventInfo;
#endif /* __cplusplus */

#endif 	/* __CRVEventInfo_FWD_DEFINED__ */


#ifndef __ReportSourceRouter_FWD_DEFINED__
#define __ReportSourceRouter_FWD_DEFINED__

#ifdef __cplusplus
typedef class ReportSourceRouter ReportSourceRouter;
#else
typedef struct ReportSourceRouter ReportSourceRouter;
#endif /* __cplusplus */

#endif 	/* __ReportSourceRouter_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "rsidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

/* interface __MIDL_itf_ReportAXViewer_0000 */
/* [local] */ 

typedef 
enum CRLoadingType
    {	crLoadingNothing	= 0,
	crLoadingPages	= crLoadingNothing + 1,
	crLoadingTotaller	= crLoadingPages + 1,
	crLoadingQueryInfo	= crLoadingTotaller + 1
    }	CRLoadingType;

typedef 
enum CRObjectType
    {	crUnknownFieldDefType	= 0,
	crDatabaseFieldType	= crUnknownFieldDefType + 1,
	crOLAPDimensionFieldType	= crDatabaseFieldType + 1,
	crOLAPDataFieldType	= crOLAPDimensionFieldType + 1,
	crOLAPCrossTabFieldType	= crOLAPDataFieldType + 1,
	crFormulaFieldType	= crOLAPCrossTabFieldType + 1,
	crSummaryFieldType	= crFormulaFieldType + 1,
	crSpecialVarFieldType	= crSummaryFieldType + 1,
	crGroupNameFieldType	= crSpecialVarFieldType + 1,
	crPromptingVarFieldType	= crGroupNameFieldType + 1,
	crText	= 100,
	crOLEObject	= crText + 1,
	crSubreport	= crOLEObject + 1,
	crBitmap	= crSubreport + 1,
	crBlob	= crBitmap + 1,
	crLine	= crBlob + 1,
	crBox	= crLine + 1,
	crGroupChart	= crBox + 1,
	crCrosstabChart	= crGroupChart + 1,
	crDetailChart	= crCrosstabChart + 1,
	crCrossTab	= crDetailChart + 1,
	crGraphic	= crCrossTab + 1,
	crOOPSubreport	= crGraphic + 1,
	crOLAPChart	= crOOPSubreport + 1,
	crGroupMap	= crOLAPChart + 1,
	crCrosstabMap	= crGroupMap + 1,
	crDetailMap	= crCrosstabMap + 1,
	crOLAPMap	= crDetailMap + 1,
	crGroupHeaderSection	= 200,
	crGroupFooterSection	= crGroupHeaderSection + 1,
	crDetailSection	= crGroupFooterSection + 1,
	crReportHeaderSection	= crDetailSection + 1,
	crReportFooterSection	= crReportHeaderSection + 1,
	crPageHeaderSection	= crReportFooterSection + 1,
	crPageFooterSection	= crPageHeaderSection + 1
    }	CRObjectType;

typedef 
enum CRTrackCursor
    {	crDefaultCursor	= 0,
	crArrowCursor	= 1,
	crCrossCursor	= 2,
	crIBeamCursor	= 3,
	crNoCursor	= 10,
	crWaitCursor	= 11,
	crAppStartingCursor	= 12,
	crHelpCursor	= 13,
	crMagnifyCursor	= 99
    }	CRTrackCursor;

typedef 
enum CRDrillType
    {	crDrillOnGroup	= 0,
	crDrillOnGroupTree	= crDrillOnGroup + 1,
	crDrillOnGraph	= crDrillOnGroupTree + 1,
	crDrillOnMap	= crDrillOnGraph + 1,
	crDrillOnSubreport	= crDrillOnMap + 1
    }	CRDrillType;

typedef 
enum CRFieldType
    {	crInt8	= 0,
	crInt16	= crInt8 + 1,
	crInt32	= crInt16 + 1,
	crNumber	= crInt32 + 1,
	crCurrency	= crNumber + 1,
	crBoolean	= crCurrency + 1,
	crDate	= crBoolean + 1,
	crTime	= crDate + 1,
	crDateTime	= crTime + 1,
	crString	= crDateTime + 1,
	crUnknownFieldType	= 255
    }	CRFieldType;

typedef /* [hidden][uuid] */ 
enum Server_Type
    {	ReportSource	= -1,
	ReportWebServer	= 0,
	CInfoServer	= 1,
	CIWebServer	= 3
    }	Server_Type;



extern RPC_IF_HANDLE __MIDL_itf_ReportAXViewer_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_ReportAXViewer_0000_v0_0_s_ifspec;

#ifndef __ICRVTrackCursorInfo_INTERFACE_DEFINED__
#define __ICRVTrackCursorInfo_INTERFACE_DEFINED__

/* interface ICRVTrackCursorInfo */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICRVTrackCursorInfo;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("13FA5946-561C-11D1-BE3F-00A0C95A6A5C")
    ICRVTrackCursorInfo : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DetailAreaCursor( 
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DetailAreaCursor( 
            /* [in] */ CRTrackCursor newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DetailAreaFieldCursor( 
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DetailAreaFieldCursor( 
            /* [in] */ CRTrackCursor newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_GraphCursor( 
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_GraphCursor( 
            /* [in] */ CRTrackCursor newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_GroupAreaCursor( 
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_GroupAreaCursor( 
            /* [in] */ CRTrackCursor newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_GroupAreaFieldCursor( 
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_GroupAreaFieldCursor( 
            /* [in] */ CRTrackCursor newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICRVTrackCursorInfoVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICRVTrackCursorInfo __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICRVTrackCursorInfo __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DetailAreaCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DetailAreaCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ CRTrackCursor newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DetailAreaFieldCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DetailAreaFieldCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ CRTrackCursor newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_GraphCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_GraphCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ CRTrackCursor newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_GroupAreaCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_GroupAreaCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ CRTrackCursor newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_GroupAreaFieldCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_GroupAreaFieldCursor )( 
            ICRVTrackCursorInfo __RPC_FAR * This,
            /* [in] */ CRTrackCursor newVal);
        
        END_INTERFACE
    } ICRVTrackCursorInfoVtbl;

    interface ICRVTrackCursorInfo
    {
        CONST_VTBL struct ICRVTrackCursorInfoVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICRVTrackCursorInfo_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICRVTrackCursorInfo_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICRVTrackCursorInfo_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICRVTrackCursorInfo_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICRVTrackCursorInfo_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICRVTrackCursorInfo_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICRVTrackCursorInfo_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICRVTrackCursorInfo_get_DetailAreaCursor(This,pVal)	\
    (This)->lpVtbl -> get_DetailAreaCursor(This,pVal)

#define ICRVTrackCursorInfo_put_DetailAreaCursor(This,newVal)	\
    (This)->lpVtbl -> put_DetailAreaCursor(This,newVal)

#define ICRVTrackCursorInfo_get_DetailAreaFieldCursor(This,pVal)	\
    (This)->lpVtbl -> get_DetailAreaFieldCursor(This,pVal)

#define ICRVTrackCursorInfo_put_DetailAreaFieldCursor(This,newVal)	\
    (This)->lpVtbl -> put_DetailAreaFieldCursor(This,newVal)

#define ICRVTrackCursorInfo_get_GraphCursor(This,pVal)	\
    (This)->lpVtbl -> get_GraphCursor(This,pVal)

#define ICRVTrackCursorInfo_put_GraphCursor(This,newVal)	\
    (This)->lpVtbl -> put_GraphCursor(This,newVal)

#define ICRVTrackCursorInfo_get_GroupAreaCursor(This,pVal)	\
    (This)->lpVtbl -> get_GroupAreaCursor(This,pVal)

#define ICRVTrackCursorInfo_put_GroupAreaCursor(This,newVal)	\
    (This)->lpVtbl -> put_GroupAreaCursor(This,newVal)

#define ICRVTrackCursorInfo_get_GroupAreaFieldCursor(This,pVal)	\
    (This)->lpVtbl -> get_GroupAreaFieldCursor(This,pVal)

#define ICRVTrackCursorInfo_put_GroupAreaFieldCursor(This,newVal)	\
    (This)->lpVtbl -> put_GroupAreaFieldCursor(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_get_DetailAreaCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);


void __RPC_STUB ICRVTrackCursorInfo_get_DetailAreaCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_put_DetailAreaCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [in] */ CRTrackCursor newVal);


void __RPC_STUB ICRVTrackCursorInfo_put_DetailAreaCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_get_DetailAreaFieldCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);


void __RPC_STUB ICRVTrackCursorInfo_get_DetailAreaFieldCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_put_DetailAreaFieldCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [in] */ CRTrackCursor newVal);


void __RPC_STUB ICRVTrackCursorInfo_put_DetailAreaFieldCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_get_GraphCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);


void __RPC_STUB ICRVTrackCursorInfo_get_GraphCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_put_GraphCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [in] */ CRTrackCursor newVal);


void __RPC_STUB ICRVTrackCursorInfo_put_GraphCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_get_GroupAreaCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);


void __RPC_STUB ICRVTrackCursorInfo_get_GroupAreaCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_put_GroupAreaCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [in] */ CRTrackCursor newVal);


void __RPC_STUB ICRVTrackCursorInfo_put_GroupAreaCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_get_GroupAreaFieldCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [retval][out] */ CRTrackCursor __RPC_FAR *pVal);


void __RPC_STUB ICRVTrackCursorInfo_get_GroupAreaFieldCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICRVTrackCursorInfo_put_GroupAreaFieldCursor_Proxy( 
    ICRVTrackCursorInfo __RPC_FAR * This,
    /* [in] */ CRTrackCursor newVal);


void __RPC_STUB ICRVTrackCursorInfo_put_GroupAreaFieldCursor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICRVTrackCursorInfo_INTERFACE_DEFINED__ */


#ifndef __ICRViewer_INTERFACE_DEFINED__
#define __ICRViewer_INTERFACE_DEFINED__

/* interface ICRViewer */
/* [unique][helpstring][hidden][dual][uuid][object] */ 


EXTERN_C const IID IID_ICRViewer;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("C4847595-972C-11D0-9567-00A0C9273C2A")
    ICRViewer : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ReportName( 
            /* [retval][out] */ BSTR __RPC_FAR *reportName) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ReportName( 
            /* [in] */ BSTR reportName) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ShowGroupTree( 
            /* [retval][out] */ BOOL __RPC_FAR *showGroupTree) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ShowGroupTree( 
            /* [in] */ BOOL showGroupTree) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ShowToolbar( 
            /* [retval][out] */ BOOL __RPC_FAR *showToolbar) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ShowToolbar( 
            /* [in] */ BOOL showToolbar) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasGroupTree( 
            /* [retval][out] */ BOOL __RPC_FAR *hasGroupTree) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasGroupTree( 
            /* [in] */ BOOL hasGroupTree) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasNavigationControls( 
            /* [retval][out] */ BOOL __RPC_FAR *hasNavigationControls) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasNavigationControls( 
            /* [in] */ BOOL hasNavigationControls) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasStopButton( 
            /* [retval][out] */ BOOL __RPC_FAR *hasStopButton) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasStopButton( 
            /* [in] */ BOOL hasStopButton) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasPrintButton( 
            /* [retval][out] */ BOOL __RPC_FAR *hasPrintButton) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasPrintButton( 
            /* [in] */ BOOL hasPrintButton) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasZoomControl( 
            /* [retval][out] */ BOOL __RPC_FAR *hasZoomControl) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasZoomControl( 
            /* [in] */ BOOL hasZoomControl) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasCloseButton( 
            /* [retval][out] */ BOOL __RPC_FAR *hasCloseButton) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasCloseButton( 
            /* [in] */ BOOL hasCloseButton) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasProgressControl( 
            /* [retval][out] */ BOOL __RPC_FAR *hasProgressControl) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasProgressControl( 
            /* [in] */ BOOL hasProgressControl) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasSearchControl( 
            /* [retval][out] */ BOOL __RPC_FAR *hasSearchControl) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasSearchControl( 
            /* [in] */ BOOL hasSearchControl) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasRefreshButton( 
            /* [retval][out] */ BOOL __RPC_FAR *hasRefreshButton) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasRefreshButton( 
            /* [in] */ BOOL hasRefreshButton) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CanDrillDown( 
            /* [retval][out] */ BOOL __RPC_FAR *canDrillDown) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_CanDrillDown( 
            /* [in] */ BOOL canDrillDown) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasAnimationCtrl( 
            /* [retval][out] */ BOOL __RPC_FAR *hasAnimationCtrl) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasAnimationCtrl( 
            /* [in] */ BOOL hasAnimationCtrl) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICRViewerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICRViewer __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICRViewer __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICRViewer __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ReportName )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *reportName);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ReportName )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BSTR reportName);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ShowGroupTree )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *showGroupTree);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ShowGroupTree )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL showGroupTree);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ShowToolbar )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *showToolbar);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ShowToolbar )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL showToolbar);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasGroupTree )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasGroupTree);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasGroupTree )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasGroupTree);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasNavigationControls )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasNavigationControls);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasNavigationControls )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasNavigationControls);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasStopButton )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasStopButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasStopButton )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasStopButton);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasPrintButton )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasPrintButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasPrintButton )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasPrintButton);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasZoomControl )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasZoomControl);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasZoomControl )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasZoomControl);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasCloseButton )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasCloseButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasCloseButton )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasCloseButton);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasProgressControl )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasProgressControl);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasProgressControl )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasProgressControl);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasSearchControl )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasSearchControl);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasSearchControl )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasSearchControl);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasRefreshButton )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasRefreshButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasRefreshButton )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasRefreshButton);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CanDrillDown )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *canDrillDown);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CanDrillDown )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL canDrillDown);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasAnimationCtrl )( 
            ICRViewer __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasAnimationCtrl);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasAnimationCtrl )( 
            ICRViewer __RPC_FAR * This,
            /* [in] */ BOOL hasAnimationCtrl);
        
        END_INTERFACE
    } ICRViewerVtbl;

    interface ICRViewer
    {
        CONST_VTBL struct ICRViewerVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICRViewer_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICRViewer_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICRViewer_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICRViewer_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICRViewer_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICRViewer_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICRViewer_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICRViewer_get_ReportName(This,reportName)	\
    (This)->lpVtbl -> get_ReportName(This,reportName)

#define ICRViewer_put_ReportName(This,reportName)	\
    (This)->lpVtbl -> put_ReportName(This,reportName)

#define ICRViewer_get_ShowGroupTree(This,showGroupTree)	\
    (This)->lpVtbl -> get_ShowGroupTree(This,showGroupTree)

#define ICRViewer_put_ShowGroupTree(This,showGroupTree)	\
    (This)->lpVtbl -> put_ShowGroupTree(This,showGroupTree)

#define ICRViewer_get_ShowToolbar(This,showToolbar)	\
    (This)->lpVtbl -> get_ShowToolbar(This,showToolbar)

#define ICRViewer_put_ShowToolbar(This,showToolbar)	\
    (This)->lpVtbl -> put_ShowToolbar(This,showToolbar)

#define ICRViewer_get_HasGroupTree(This,hasGroupTree)	\
    (This)->lpVtbl -> get_HasGroupTree(This,hasGroupTree)

#define ICRViewer_put_HasGroupTree(This,hasGroupTree)	\
    (This)->lpVtbl -> put_HasGroupTree(This,hasGroupTree)

#define ICRViewer_get_HasNavigationControls(This,hasNavigationControls)	\
    (This)->lpVtbl -> get_HasNavigationControls(This,hasNavigationControls)

#define ICRViewer_put_HasNavigationControls(This,hasNavigationControls)	\
    (This)->lpVtbl -> put_HasNavigationControls(This,hasNavigationControls)

#define ICRViewer_get_HasStopButton(This,hasStopButton)	\
    (This)->lpVtbl -> get_HasStopButton(This,hasStopButton)

#define ICRViewer_put_HasStopButton(This,hasStopButton)	\
    (This)->lpVtbl -> put_HasStopButton(This,hasStopButton)

#define ICRViewer_get_HasPrintButton(This,hasPrintButton)	\
    (This)->lpVtbl -> get_HasPrintButton(This,hasPrintButton)

#define ICRViewer_put_HasPrintButton(This,hasPrintButton)	\
    (This)->lpVtbl -> put_HasPrintButton(This,hasPrintButton)

#define ICRViewer_get_HasZoomControl(This,hasZoomControl)	\
    (This)->lpVtbl -> get_HasZoomControl(This,hasZoomControl)

#define ICRViewer_put_HasZoomControl(This,hasZoomControl)	\
    (This)->lpVtbl -> put_HasZoomControl(This,hasZoomControl)

#define ICRViewer_get_HasCloseButton(This,hasCloseButton)	\
    (This)->lpVtbl -> get_HasCloseButton(This,hasCloseButton)

#define ICRViewer_put_HasCloseButton(This,hasCloseButton)	\
    (This)->lpVtbl -> put_HasCloseButton(This,hasCloseButton)

#define ICRViewer_get_HasProgressControl(This,hasProgressControl)	\
    (This)->lpVtbl -> get_HasProgressControl(This,hasProgressControl)

#define ICRViewer_put_HasProgressControl(This,hasProgressControl)	\
    (This)->lpVtbl -> put_HasProgressControl(This,hasProgressControl)

#define ICRViewer_get_HasSearchControl(This,hasSearchControl)	\
    (This)->lpVtbl -> get_HasSearchControl(This,hasSearchControl)

#define ICRViewer_put_HasSearchControl(This,hasSearchControl)	\
    (This)->lpVtbl -> put_HasSearchControl(This,hasSearchControl)

#define ICRViewer_get_HasRefreshButton(This,hasRefreshButton)	\
    (This)->lpVtbl -> get_HasRefreshButton(This,hasRefreshButton)

#define ICRViewer_put_HasRefreshButton(This,hasRefreshButton)	\
    (This)->lpVtbl -> put_HasRefreshButton(This,hasRefreshButton)

#define ICRViewer_get_CanDrillDown(This,canDrillDown)	\
    (This)->lpVtbl -> get_CanDrillDown(This,canDrillDown)

#define ICRViewer_put_CanDrillDown(This,canDrillDown)	\
    (This)->lpVtbl -> put_CanDrillDown(This,canDrillDown)

#define ICRViewer_get_HasAnimationCtrl(This,hasAnimationCtrl)	\
    (This)->lpVtbl -> get_HasAnimationCtrl(This,hasAnimationCtrl)

#define ICRViewer_put_HasAnimationCtrl(This,hasAnimationCtrl)	\
    (This)->lpVtbl -> put_HasAnimationCtrl(This,hasAnimationCtrl)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_ReportName_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *reportName);


void __RPC_STUB ICRViewer_get_ReportName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_ReportName_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BSTR reportName);


void __RPC_STUB ICRViewer_put_ReportName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_ShowGroupTree_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *showGroupTree);


void __RPC_STUB ICRViewer_get_ShowGroupTree_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_ShowGroupTree_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL showGroupTree);


void __RPC_STUB ICRViewer_put_ShowGroupTree_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_ShowToolbar_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *showToolbar);


void __RPC_STUB ICRViewer_get_ShowToolbar_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_ShowToolbar_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL showToolbar);


void __RPC_STUB ICRViewer_put_ShowToolbar_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasGroupTree_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasGroupTree);


void __RPC_STUB ICRViewer_get_HasGroupTree_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasGroupTree_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasGroupTree);


void __RPC_STUB ICRViewer_put_HasGroupTree_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasNavigationControls_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasNavigationControls);


void __RPC_STUB ICRViewer_get_HasNavigationControls_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasNavigationControls_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasNavigationControls);


void __RPC_STUB ICRViewer_put_HasNavigationControls_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasStopButton_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasStopButton);


void __RPC_STUB ICRViewer_get_HasStopButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasStopButton_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasStopButton);


void __RPC_STUB ICRViewer_put_HasStopButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasPrintButton_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasPrintButton);


void __RPC_STUB ICRViewer_get_HasPrintButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasPrintButton_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasPrintButton);


void __RPC_STUB ICRViewer_put_HasPrintButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasZoomControl_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasZoomControl);


void __RPC_STUB ICRViewer_get_HasZoomControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasZoomControl_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasZoomControl);


void __RPC_STUB ICRViewer_put_HasZoomControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasCloseButton_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasCloseButton);


void __RPC_STUB ICRViewer_get_HasCloseButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasCloseButton_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasCloseButton);


void __RPC_STUB ICRViewer_put_HasCloseButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasProgressControl_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasProgressControl);


void __RPC_STUB ICRViewer_get_HasProgressControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasProgressControl_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasProgressControl);


void __RPC_STUB ICRViewer_put_HasProgressControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasSearchControl_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasSearchControl);


void __RPC_STUB ICRViewer_get_HasSearchControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasSearchControl_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasSearchControl);


void __RPC_STUB ICRViewer_put_HasSearchControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasRefreshButton_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasRefreshButton);


void __RPC_STUB ICRViewer_get_HasRefreshButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasRefreshButton_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasRefreshButton);


void __RPC_STUB ICRViewer_put_HasRefreshButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_CanDrillDown_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *canDrillDown);


void __RPC_STUB ICRViewer_get_CanDrillDown_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_CanDrillDown_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL canDrillDown);


void __RPC_STUB ICRViewer_put_CanDrillDown_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer_get_HasAnimationCtrl_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasAnimationCtrl);


void __RPC_STUB ICRViewer_get_HasAnimationCtrl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer_put_HasAnimationCtrl_Proxy( 
    ICRViewer __RPC_FAR * This,
    /* [in] */ BOOL hasAnimationCtrl);


void __RPC_STUB ICRViewer_put_HasAnimationCtrl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICRViewer_INTERFACE_DEFINED__ */


#ifndef __ICRViewer2_INTERFACE_DEFINED__
#define __ICRViewer2_INTERFACE_DEFINED__

/* interface ICRViewer2 */
/* [unique][helpstring][hidden][dual][uuid][object] */ 


EXTERN_C const IID IID_ICRViewer2;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("9CC31500-346E-11d1-88D3-00A0C9311D28")
    ICRViewer2 : public ICRViewer
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ServerType( 
            /* [retval][out] */ Server_Type __RPC_FAR *serverType) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ServerType( 
            /* [in] */ Server_Type serverType) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CINFO_APSName( 
            /* [retval][out] */ BSTR __RPC_FAR *apsName) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_CINFO_APSName( 
            /* [in] */ BSTR reportName) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CINFO_UserName( 
            /* [retval][out] */ BSTR __RPC_FAR *userName) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_CINFO_UserName( 
            /* [in] */ BSTR userName) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CINFO_Password( 
            /* [retval][out] */ BSTR __RPC_FAR *password) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_CINFO_Password( 
            /* [in] */ BSTR password) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CINFO_AgentKey( 
            /* [retval][out] */ LONG __RPC_FAR *agentKey) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_CINFO_AgentKey( 
            /* [in] */ LONG agentKey) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CINFO_LoginHandle( 
            /* [retval][out] */ LONG __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_CINFO_LoginHandle( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CINFO_SessionHandle( 
            /* [retval][out] */ LONG __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_CINFO_SessionHandle( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OnCInfoCallBack( 
            /* [in] */ long wParam,
            /* [in] */ long lParam) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE BindToServer( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ViewReport( void) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HasSelectExpertButton( 
            /* [retval][out] */ BOOL __RPC_FAR *hasSelectExpertButton) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HasSelectExpertButton( 
            /* [in] */ BOOL hasSelectExpertButton) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICRViewer2Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICRViewer2 __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICRViewer2 __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICRViewer2 __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ReportName )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *reportName);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ReportName )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BSTR reportName);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ShowGroupTree )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *showGroupTree);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ShowGroupTree )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL showGroupTree);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ShowToolbar )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *showToolbar);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ShowToolbar )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL showToolbar);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasGroupTree )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasGroupTree);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasGroupTree )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasGroupTree);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasNavigationControls )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasNavigationControls);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasNavigationControls )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasNavigationControls);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasStopButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasStopButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasStopButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasStopButton);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasPrintButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasPrintButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasPrintButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasPrintButton);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasZoomControl )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasZoomControl);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasZoomControl )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasZoomControl);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasCloseButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasCloseButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasCloseButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasCloseButton);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasProgressControl )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasProgressControl);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasProgressControl )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasProgressControl);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasSearchControl )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasSearchControl);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasSearchControl )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasSearchControl);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasRefreshButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasRefreshButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasRefreshButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasRefreshButton);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CanDrillDown )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *canDrillDown);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CanDrillDown )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL canDrillDown);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasAnimationCtrl )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasAnimationCtrl);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasAnimationCtrl )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasAnimationCtrl);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ServerType )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ Server_Type __RPC_FAR *serverType);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ServerType )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ Server_Type serverType);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CINFO_APSName )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *apsName);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CINFO_APSName )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BSTR reportName);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CINFO_UserName )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *userName);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CINFO_UserName )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BSTR userName);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CINFO_Password )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *password);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CINFO_Password )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BSTR password);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CINFO_AgentKey )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ LONG __RPC_FAR *agentKey);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CINFO_AgentKey )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ LONG agentKey);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CINFO_LoginHandle )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ LONG __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CINFO_LoginHandle )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CINFO_SessionHandle )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ LONG __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CINFO_SessionHandle )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OnCInfoCallBack )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ long wParam,
            /* [in] */ long lParam);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *BindToServer )( 
            ICRViewer2 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ViewReport )( 
            ICRViewer2 __RPC_FAR * This);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasSelectExpertButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [retval][out] */ BOOL __RPC_FAR *hasSelectExpertButton);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasSelectExpertButton )( 
            ICRViewer2 __RPC_FAR * This,
            /* [in] */ BOOL hasSelectExpertButton);
        
        END_INTERFACE
    } ICRViewer2Vtbl;

    interface ICRViewer2
    {
        CONST_VTBL struct ICRViewer2Vtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICRViewer2_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICRViewer2_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICRViewer2_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICRViewer2_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICRViewer2_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICRViewer2_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICRViewer2_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICRViewer2_get_ReportName(This,reportName)	\
    (This)->lpVtbl -> get_ReportName(This,reportName)

#define ICRViewer2_put_ReportName(This,reportName)	\
    (This)->lpVtbl -> put_ReportName(This,reportName)

#define ICRViewer2_get_ShowGroupTree(This,showGroupTree)	\
    (This)->lpVtbl -> get_ShowGroupTree(This,showGroupTree)

#define ICRViewer2_put_ShowGroupTree(This,showGroupTree)	\
    (This)->lpVtbl -> put_ShowGroupTree(This,showGroupTree)

#define ICRViewer2_get_ShowToolbar(This,showToolbar)	\
    (This)->lpVtbl -> get_ShowToolbar(This,showToolbar)

#define ICRViewer2_put_ShowToolbar(This,showToolbar)	\
    (This)->lpVtbl -> put_ShowToolbar(This,showToolbar)

#define ICRViewer2_get_HasGroupTree(This,hasGroupTree)	\
    (This)->lpVtbl -> get_HasGroupTree(This,hasGroupTree)

#define ICRViewer2_put_HasGroupTree(This,hasGroupTree)	\
    (This)->lpVtbl -> put_HasGroupTree(This,hasGroupTree)

#define ICRViewer2_get_HasNavigationControls(This,hasNavigationControls)	\
    (This)->lpVtbl -> get_HasNavigationControls(This,hasNavigationControls)

#define ICRViewer2_put_HasNavigationControls(This,hasNavigationControls)	\
    (This)->lpVtbl -> put_HasNavigationControls(This,hasNavigationControls)

#define ICRViewer2_get_HasStopButton(This,hasStopButton)	\
    (This)->lpVtbl -> get_HasStopButton(This,hasStopButton)

#define ICRViewer2_put_HasStopButton(This,hasStopButton)	\
    (This)->lpVtbl -> put_HasStopButton(This,hasStopButton)

#define ICRViewer2_get_HasPrintButton(This,hasPrintButton)	\
    (This)->lpVtbl -> get_HasPrintButton(This,hasPrintButton)

#define ICRViewer2_put_HasPrintButton(This,hasPrintButton)	\
    (This)->lpVtbl -> put_HasPrintButton(This,hasPrintButton)

#define ICRViewer2_get_HasZoomControl(This,hasZoomControl)	\
    (This)->lpVtbl -> get_HasZoomControl(This,hasZoomControl)

#define ICRViewer2_put_HasZoomControl(This,hasZoomControl)	\
    (This)->lpVtbl -> put_HasZoomControl(This,hasZoomControl)

#define ICRViewer2_get_HasCloseButton(This,hasCloseButton)	\
    (This)->lpVtbl -> get_HasCloseButton(This,hasCloseButton)

#define ICRViewer2_put_HasCloseButton(This,hasCloseButton)	\
    (This)->lpVtbl -> put_HasCloseButton(This,hasCloseButton)

#define ICRViewer2_get_HasProgressControl(This,hasProgressControl)	\
    (This)->lpVtbl -> get_HasProgressControl(This,hasProgressControl)

#define ICRViewer2_put_HasProgressControl(This,hasProgressControl)	\
    (This)->lpVtbl -> put_HasProgressControl(This,hasProgressControl)

#define ICRViewer2_get_HasSearchControl(This,hasSearchControl)	\
    (This)->lpVtbl -> get_HasSearchControl(This,hasSearchControl)

#define ICRViewer2_put_HasSearchControl(This,hasSearchControl)	\
    (This)->lpVtbl -> put_HasSearchControl(This,hasSearchControl)

#define ICRViewer2_get_HasRefreshButton(This,hasRefreshButton)	\
    (This)->lpVtbl -> get_HasRefreshButton(This,hasRefreshButton)

#define ICRViewer2_put_HasRefreshButton(This,hasRefreshButton)	\
    (This)->lpVtbl -> put_HasRefreshButton(This,hasRefreshButton)

#define ICRViewer2_get_CanDrillDown(This,canDrillDown)	\
    (This)->lpVtbl -> get_CanDrillDown(This,canDrillDown)

#define ICRViewer2_put_CanDrillDown(This,canDrillDown)	\
    (This)->lpVtbl -> put_CanDrillDown(This,canDrillDown)

#define ICRViewer2_get_HasAnimationCtrl(This,hasAnimationCtrl)	\
    (This)->lpVtbl -> get_HasAnimationCtrl(This,hasAnimationCtrl)

#define ICRViewer2_put_HasAnimationCtrl(This,hasAnimationCtrl)	\
    (This)->lpVtbl -> put_HasAnimationCtrl(This,hasAnimationCtrl)


#define ICRViewer2_get_ServerType(This,serverType)	\
    (This)->lpVtbl -> get_ServerType(This,serverType)

#define ICRViewer2_put_ServerType(This,serverType)	\
    (This)->lpVtbl -> put_ServerType(This,serverType)

#define ICRViewer2_get_CINFO_APSName(This,apsName)	\
    (This)->lpVtbl -> get_CINFO_APSName(This,apsName)

#define ICRViewer2_put_CINFO_APSName(This,reportName)	\
    (This)->lpVtbl -> put_CINFO_APSName(This,reportName)

#define ICRViewer2_get_CINFO_UserName(This,userName)	\
    (This)->lpVtbl -> get_CINFO_UserName(This,userName)

#define ICRViewer2_put_CINFO_UserName(This,userName)	\
    (This)->lpVtbl -> put_CINFO_UserName(This,userName)

#define ICRViewer2_get_CINFO_Password(This,password)	\
    (This)->lpVtbl -> get_CINFO_Password(This,password)

#define ICRViewer2_put_CINFO_Password(This,password)	\
    (This)->lpVtbl -> put_CINFO_Password(This,password)

#define ICRViewer2_get_CINFO_AgentKey(This,agentKey)	\
    (This)->lpVtbl -> get_CINFO_AgentKey(This,agentKey)

#define ICRViewer2_put_CINFO_AgentKey(This,agentKey)	\
    (This)->lpVtbl -> put_CINFO_AgentKey(This,agentKey)

#define ICRViewer2_get_CINFO_LoginHandle(This,pVal)	\
    (This)->lpVtbl -> get_CINFO_LoginHandle(This,pVal)

#define ICRViewer2_put_CINFO_LoginHandle(This,newVal)	\
    (This)->lpVtbl -> put_CINFO_LoginHandle(This,newVal)

#define ICRViewer2_get_CINFO_SessionHandle(This,pVal)	\
    (This)->lpVtbl -> get_CINFO_SessionHandle(This,pVal)

#define ICRViewer2_put_CINFO_SessionHandle(This,newVal)	\
    (This)->lpVtbl -> put_CINFO_SessionHandle(This,newVal)

#define ICRViewer2_OnCInfoCallBack(This,wParam,lParam)	\
    (This)->lpVtbl -> OnCInfoCallBack(This,wParam,lParam)

#define ICRViewer2_BindToServer(This)	\
    (This)->lpVtbl -> BindToServer(This)

#define ICRViewer2_ViewReport(This)	\
    (This)->lpVtbl -> ViewReport(This)

#define ICRViewer2_get_HasSelectExpertButton(This,hasSelectExpertButton)	\
    (This)->lpVtbl -> get_HasSelectExpertButton(This,hasSelectExpertButton)

#define ICRViewer2_put_HasSelectExpertButton(This,hasSelectExpertButton)	\
    (This)->lpVtbl -> put_HasSelectExpertButton(This,hasSelectExpertButton)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer2_get_ServerType_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [retval][out] */ Server_Type __RPC_FAR *serverType);


void __RPC_STUB ICRViewer2_get_ServerType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer2_put_ServerType_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ Server_Type serverType);


void __RPC_STUB ICRViewer2_put_ServerType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer2_get_CINFO_APSName_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *apsName);


void __RPC_STUB ICRViewer2_get_CINFO_APSName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer2_put_CINFO_APSName_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ BSTR reportName);


void __RPC_STUB ICRViewer2_put_CINFO_APSName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer2_get_CINFO_UserName_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *userName);


void __RPC_STUB ICRViewer2_get_CINFO_UserName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer2_put_CINFO_UserName_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ BSTR userName);


void __RPC_STUB ICRViewer2_put_CINFO_UserName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer2_get_CINFO_Password_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *password);


void __RPC_STUB ICRViewer2_get_CINFO_Password_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer2_put_CINFO_Password_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ BSTR password);


void __RPC_STUB ICRViewer2_put_CINFO_Password_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer2_get_CINFO_AgentKey_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [retval][out] */ LONG __RPC_FAR *agentKey);


void __RPC_STUB ICRViewer2_get_CINFO_AgentKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer2_put_CINFO_AgentKey_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ LONG agentKey);


void __RPC_STUB ICRViewer2_put_CINFO_AgentKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer2_get_CINFO_LoginHandle_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [retval][out] */ LONG __RPC_FAR *pVal);


void __RPC_STUB ICRViewer2_get_CINFO_LoginHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer2_put_CINFO_LoginHandle_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ LONG newVal);


void __RPC_STUB ICRViewer2_put_CINFO_LoginHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer2_get_CINFO_SessionHandle_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [retval][out] */ LONG __RPC_FAR *pVal);


void __RPC_STUB ICRViewer2_get_CINFO_SessionHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer2_put_CINFO_SessionHandle_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ LONG newVal);


void __RPC_STUB ICRViewer2_put_CINFO_SessionHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICRViewer2_OnCInfoCallBack_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ long wParam,
    /* [in] */ long lParam);


void __RPC_STUB ICRViewer2_OnCInfoCallBack_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICRViewer2_BindToServer_Proxy( 
    ICRViewer2 __RPC_FAR * This);


void __RPC_STUB ICRViewer2_BindToServer_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICRViewer2_ViewReport_Proxy( 
    ICRViewer2 __RPC_FAR * This);


void __RPC_STUB ICRViewer2_ViewReport_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE ICRViewer2_get_HasSelectExpertButton_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [retval][out] */ BOOL __RPC_FAR *hasSelectExpertButton);


void __RPC_STUB ICRViewer2_get_HasSelectExpertButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE ICRViewer2_put_HasSelectExpertButton_Proxy( 
    ICRViewer2 __RPC_FAR * This,
    /* [in] */ BOOL hasSelectExpertButton);


void __RPC_STUB ICRViewer2_put_HasSelectExpertButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICRViewer2_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportViewer_INTERFACE_DEFINED__
#define __ICrystalReportViewer_INTERFACE_DEFINED__

/* interface ICrystalReportViewer */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICrystalReportViewer;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("C09F8DF1-D331-11d1-A818-00A0C92784CD")
    ICrystalReportViewer : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ReportSource( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ReportSource( 
            /* [in] */ LPUNKNOWN lpUnknown) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DisplayGroupTree( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayGroupTree) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DisplayGroupTree( 
            /* [in] */ VARIANT_BOOL displayGroupTree) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DisplayToolbar( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayToolbar) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DisplayToolbar( 
            /* [in] */ VARIANT_BOOL displayToolbar) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableGroupTree( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableGroupTree) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableGroupTree( 
            /* [in] */ VARIANT_BOOL enableGroupTree) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableNavigationControls( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableNavigationControls) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableNavigationControls( 
            /* [in] */ VARIANT_BOOL enableNavigationControls) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableStopButton( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableStopButton) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableStopButton( 
            /* [in] */ VARIANT_BOOL enableStopButton) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnablePrintButton( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enablePrintButton) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnablePrintButton( 
            /* [in] */ VARIANT_BOOL enablePrintButton) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableZoomControl( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableZoomControl) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableZoomControl( 
            /* [in] */ VARIANT_BOOL enableZoomControl) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableCloseButton( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableCloseButton) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableCloseButton( 
            /* [in] */ VARIANT_BOOL enableCloseButton) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableProgressControl( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableProgressControl) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableProgressControl( 
            /* [in] */ VARIANT_BOOL enableProgressControl) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableSearchControl( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSearchControl) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableSearchControl( 
            /* [in] */ VARIANT_BOOL enableSearchControl) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableRefreshButton( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableRefreshButton) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableRefreshButton( 
            /* [in] */ VARIANT_BOOL enableRefreshButton) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableDrillDown( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableDrillDown) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableDrillDown( 
            /* [in] */ VARIANT_BOOL enableDrillDown) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableAnimationCtrl( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableAnimationCtrl) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableAnimationCtrl( 
            /* [in] */ VARIANT_BOOL enableAnimationCtrl) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableSelectExpertButton( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSelectExpertButton) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableSelectExpertButton( 
            /* [in] */ VARIANT_BOOL enableSelectExpertButton) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ViewReport( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableToolbar( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableToolbar( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DisplayBorder( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DisplayBorder( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DisplayTabs( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DisplayTabs( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DisplayBackgroundEdge( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DisplayBackgroundEdge( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [hidden][helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SelectionFormula( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [hidden][helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_SelectionFormula( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TrackCursorInfo( 
            /* [retval][out] */ ICRVTrackCursorInfo __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ActiveViewIndex( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ViewCount( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ActivateView( 
            /* [in] */ VARIANT Index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddView( 
            /* [in] */ VARIANT GroupPath) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CloseView( 
            /* [in] */ VARIANT Index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetViewPath( 
            /* [in] */ short Index,
            /* [retval][out] */ VARIANT __RPC_FAR *Path) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE PrintReport( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Refresh( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SearchForText( 
            /* [in] */ BSTR Text) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowFirstPage( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowNextPage( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowPreviousPage( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowLastPage( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowNthPage( 
            /* [in] */ short PageNumber) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Zoom( 
            /* [in] */ short ZoomLevel) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetCurrentPageNumber( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowGroup( 
            /* [in] */ VARIANT GroupPath) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsBusy( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnablePopupMenu( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnablePopupMenu( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportViewerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ReportSource )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ReportSource )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ LPUNKNOWN lpUnknown);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayGroupTree )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayGroupTree);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayGroupTree )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL displayGroupTree);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayToolbar )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayToolbar);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayToolbar )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL displayToolbar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableGroupTree )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableGroupTree);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableGroupTree )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableGroupTree);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableNavigationControls )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableNavigationControls);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableNavigationControls )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableNavigationControls);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableStopButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableStopButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableStopButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableStopButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnablePrintButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enablePrintButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnablePrintButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enablePrintButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableZoomControl )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableZoomControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableZoomControl )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableZoomControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableCloseButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableCloseButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableCloseButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableCloseButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableProgressControl )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableProgressControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableProgressControl )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableProgressControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableSearchControl )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSearchControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableSearchControl )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableSearchControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableRefreshButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableRefreshButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableRefreshButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableRefreshButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableDrillDown )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableDrillDown);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableDrillDown )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableDrillDown);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableAnimationCtrl )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableAnimationCtrl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableAnimationCtrl )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableAnimationCtrl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableSelectExpertButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSelectExpertButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableSelectExpertButton )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableSelectExpertButton);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ViewReport )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableToolbar )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableToolbar )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayBorder )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayBorder )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayTabs )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayTabs )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayBackgroundEdge )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayBackgroundEdge )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [hidden][helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SelectionFormula )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [hidden][helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_SelectionFormula )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TrackCursorInfo )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ ICRVTrackCursorInfo __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ActiveViewIndex )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ViewCount )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ActivateView )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddView )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT GroupPath);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CloseView )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetViewPath )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ short Index,
            /* [retval][out] */ VARIANT __RPC_FAR *Path);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *PrintReport )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Refresh )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SearchForText )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ BSTR Text);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowFirstPage )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowNextPage )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowPreviousPage )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowLastPage )( 
            ICrystalReportViewer __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowNthPage )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ short PageNumber);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Zoom )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ short ZoomLevel);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCurrentPageNumber )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowGroup )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT GroupPath);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsBusy )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnablePopupMenu )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnablePopupMenu )( 
            ICrystalReportViewer __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        END_INTERFACE
    } ICrystalReportViewerVtbl;

    interface ICrystalReportViewer
    {
        CONST_VTBL struct ICrystalReportViewerVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportViewer_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportViewer_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportViewer_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportViewer_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICrystalReportViewer_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICrystalReportViewer_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICrystalReportViewer_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICrystalReportViewer_get_ReportSource(This,pVal)	\
    (This)->lpVtbl -> get_ReportSource(This,pVal)

#define ICrystalReportViewer_put_ReportSource(This,lpUnknown)	\
    (This)->lpVtbl -> put_ReportSource(This,lpUnknown)

#define ICrystalReportViewer_get_DisplayGroupTree(This,displayGroupTree)	\
    (This)->lpVtbl -> get_DisplayGroupTree(This,displayGroupTree)

#define ICrystalReportViewer_put_DisplayGroupTree(This,displayGroupTree)	\
    (This)->lpVtbl -> put_DisplayGroupTree(This,displayGroupTree)

#define ICrystalReportViewer_get_DisplayToolbar(This,displayToolbar)	\
    (This)->lpVtbl -> get_DisplayToolbar(This,displayToolbar)

#define ICrystalReportViewer_put_DisplayToolbar(This,displayToolbar)	\
    (This)->lpVtbl -> put_DisplayToolbar(This,displayToolbar)

#define ICrystalReportViewer_get_EnableGroupTree(This,enableGroupTree)	\
    (This)->lpVtbl -> get_EnableGroupTree(This,enableGroupTree)

#define ICrystalReportViewer_put_EnableGroupTree(This,enableGroupTree)	\
    (This)->lpVtbl -> put_EnableGroupTree(This,enableGroupTree)

#define ICrystalReportViewer_get_EnableNavigationControls(This,enableNavigationControls)	\
    (This)->lpVtbl -> get_EnableNavigationControls(This,enableNavigationControls)

#define ICrystalReportViewer_put_EnableNavigationControls(This,enableNavigationControls)	\
    (This)->lpVtbl -> put_EnableNavigationControls(This,enableNavigationControls)

#define ICrystalReportViewer_get_EnableStopButton(This,enableStopButton)	\
    (This)->lpVtbl -> get_EnableStopButton(This,enableStopButton)

#define ICrystalReportViewer_put_EnableStopButton(This,enableStopButton)	\
    (This)->lpVtbl -> put_EnableStopButton(This,enableStopButton)

#define ICrystalReportViewer_get_EnablePrintButton(This,enablePrintButton)	\
    (This)->lpVtbl -> get_EnablePrintButton(This,enablePrintButton)

#define ICrystalReportViewer_put_EnablePrintButton(This,enablePrintButton)	\
    (This)->lpVtbl -> put_EnablePrintButton(This,enablePrintButton)

#define ICrystalReportViewer_get_EnableZoomControl(This,enableZoomControl)	\
    (This)->lpVtbl -> get_EnableZoomControl(This,enableZoomControl)

#define ICrystalReportViewer_put_EnableZoomControl(This,enableZoomControl)	\
    (This)->lpVtbl -> put_EnableZoomControl(This,enableZoomControl)

#define ICrystalReportViewer_get_EnableCloseButton(This,enableCloseButton)	\
    (This)->lpVtbl -> get_EnableCloseButton(This,enableCloseButton)

#define ICrystalReportViewer_put_EnableCloseButton(This,enableCloseButton)	\
    (This)->lpVtbl -> put_EnableCloseButton(This,enableCloseButton)

#define ICrystalReportViewer_get_EnableProgressControl(This,enableProgressControl)	\
    (This)->lpVtbl -> get_EnableProgressControl(This,enableProgressControl)

#define ICrystalReportViewer_put_EnableProgressControl(This,enableProgressControl)	\
    (This)->lpVtbl -> put_EnableProgressControl(This,enableProgressControl)

#define ICrystalReportViewer_get_EnableSearchControl(This,enableSearchControl)	\
    (This)->lpVtbl -> get_EnableSearchControl(This,enableSearchControl)

#define ICrystalReportViewer_put_EnableSearchControl(This,enableSearchControl)	\
    (This)->lpVtbl -> put_EnableSearchControl(This,enableSearchControl)

#define ICrystalReportViewer_get_EnableRefreshButton(This,enableRefreshButton)	\
    (This)->lpVtbl -> get_EnableRefreshButton(This,enableRefreshButton)

#define ICrystalReportViewer_put_EnableRefreshButton(This,enableRefreshButton)	\
    (This)->lpVtbl -> put_EnableRefreshButton(This,enableRefreshButton)

#define ICrystalReportViewer_get_EnableDrillDown(This,enableDrillDown)	\
    (This)->lpVtbl -> get_EnableDrillDown(This,enableDrillDown)

#define ICrystalReportViewer_put_EnableDrillDown(This,enableDrillDown)	\
    (This)->lpVtbl -> put_EnableDrillDown(This,enableDrillDown)

#define ICrystalReportViewer_get_EnableAnimationCtrl(This,enableAnimationCtrl)	\
    (This)->lpVtbl -> get_EnableAnimationCtrl(This,enableAnimationCtrl)

#define ICrystalReportViewer_put_EnableAnimationCtrl(This,enableAnimationCtrl)	\
    (This)->lpVtbl -> put_EnableAnimationCtrl(This,enableAnimationCtrl)

#define ICrystalReportViewer_get_EnableSelectExpertButton(This,enableSelectExpertButton)	\
    (This)->lpVtbl -> get_EnableSelectExpertButton(This,enableSelectExpertButton)

#define ICrystalReportViewer_put_EnableSelectExpertButton(This,enableSelectExpertButton)	\
    (This)->lpVtbl -> put_EnableSelectExpertButton(This,enableSelectExpertButton)

#define ICrystalReportViewer_ViewReport(This)	\
    (This)->lpVtbl -> ViewReport(This)

#define ICrystalReportViewer_get_EnableToolbar(This,pVal)	\
    (This)->lpVtbl -> get_EnableToolbar(This,pVal)

#define ICrystalReportViewer_put_EnableToolbar(This,newVal)	\
    (This)->lpVtbl -> put_EnableToolbar(This,newVal)

#define ICrystalReportViewer_get_DisplayBorder(This,pVal)	\
    (This)->lpVtbl -> get_DisplayBorder(This,pVal)

#define ICrystalReportViewer_put_DisplayBorder(This,newVal)	\
    (This)->lpVtbl -> put_DisplayBorder(This,newVal)

#define ICrystalReportViewer_get_DisplayTabs(This,pVal)	\
    (This)->lpVtbl -> get_DisplayTabs(This,pVal)

#define ICrystalReportViewer_put_DisplayTabs(This,newVal)	\
    (This)->lpVtbl -> put_DisplayTabs(This,newVal)

#define ICrystalReportViewer_get_DisplayBackgroundEdge(This,pVal)	\
    (This)->lpVtbl -> get_DisplayBackgroundEdge(This,pVal)

#define ICrystalReportViewer_put_DisplayBackgroundEdge(This,newVal)	\
    (This)->lpVtbl -> put_DisplayBackgroundEdge(This,newVal)

#define ICrystalReportViewer_get_SelectionFormula(This,pVal)	\
    (This)->lpVtbl -> get_SelectionFormula(This,pVal)

#define ICrystalReportViewer_put_SelectionFormula(This,newVal)	\
    (This)->lpVtbl -> put_SelectionFormula(This,newVal)

#define ICrystalReportViewer_get_TrackCursorInfo(This,pVal)	\
    (This)->lpVtbl -> get_TrackCursorInfo(This,pVal)

#define ICrystalReportViewer_get_ActiveViewIndex(This,pVal)	\
    (This)->lpVtbl -> get_ActiveViewIndex(This,pVal)

#define ICrystalReportViewer_get_ViewCount(This,pVal)	\
    (This)->lpVtbl -> get_ViewCount(This,pVal)

#define ICrystalReportViewer_ActivateView(This,Index)	\
    (This)->lpVtbl -> ActivateView(This,Index)

#define ICrystalReportViewer_AddView(This,GroupPath)	\
    (This)->lpVtbl -> AddView(This,GroupPath)

#define ICrystalReportViewer_CloseView(This,Index)	\
    (This)->lpVtbl -> CloseView(This,Index)

#define ICrystalReportViewer_GetViewPath(This,Index,Path)	\
    (This)->lpVtbl -> GetViewPath(This,Index,Path)

#define ICrystalReportViewer_PrintReport(This)	\
    (This)->lpVtbl -> PrintReport(This)

#define ICrystalReportViewer_Refresh(This)	\
    (This)->lpVtbl -> Refresh(This)

#define ICrystalReportViewer_SearchForText(This,Text)	\
    (This)->lpVtbl -> SearchForText(This,Text)

#define ICrystalReportViewer_ShowFirstPage(This)	\
    (This)->lpVtbl -> ShowFirstPage(This)

#define ICrystalReportViewer_ShowNextPage(This)	\
    (This)->lpVtbl -> ShowNextPage(This)

#define ICrystalReportViewer_ShowPreviousPage(This)	\
    (This)->lpVtbl -> ShowPreviousPage(This)

#define ICrystalReportViewer_ShowLastPage(This)	\
    (This)->lpVtbl -> ShowLastPage(This)

#define ICrystalReportViewer_ShowNthPage(This,PageNumber)	\
    (This)->lpVtbl -> ShowNthPage(This,PageNumber)

#define ICrystalReportViewer_Zoom(This,ZoomLevel)	\
    (This)->lpVtbl -> Zoom(This,ZoomLevel)

#define ICrystalReportViewer_GetCurrentPageNumber(This,pVal)	\
    (This)->lpVtbl -> GetCurrentPageNumber(This,pVal)

#define ICrystalReportViewer_ShowGroup(This,GroupPath)	\
    (This)->lpVtbl -> ShowGroup(This,GroupPath)

#define ICrystalReportViewer_get_IsBusy(This,pVal)	\
    (This)->lpVtbl -> get_IsBusy(This,pVal)

#define ICrystalReportViewer_get_EnablePopupMenu(This,pVal)	\
    (This)->lpVtbl -> get_EnablePopupMenu(This,pVal)

#define ICrystalReportViewer_put_EnablePopupMenu(This,newVal)	\
    (This)->lpVtbl -> put_EnablePopupMenu(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_ReportSource_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_ReportSource_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_ReportSource_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ LPUNKNOWN lpUnknown);


void __RPC_STUB ICrystalReportViewer_put_ReportSource_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_DisplayGroupTree_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayGroupTree);


void __RPC_STUB ICrystalReportViewer_get_DisplayGroupTree_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_DisplayGroupTree_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL displayGroupTree);


void __RPC_STUB ICrystalReportViewer_put_DisplayGroupTree_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_DisplayToolbar_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayToolbar);


void __RPC_STUB ICrystalReportViewer_get_DisplayToolbar_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_DisplayToolbar_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL displayToolbar);


void __RPC_STUB ICrystalReportViewer_put_DisplayToolbar_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableGroupTree_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableGroupTree);


void __RPC_STUB ICrystalReportViewer_get_EnableGroupTree_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableGroupTree_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableGroupTree);


void __RPC_STUB ICrystalReportViewer_put_EnableGroupTree_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableNavigationControls_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableNavigationControls);


void __RPC_STUB ICrystalReportViewer_get_EnableNavigationControls_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableNavigationControls_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableNavigationControls);


void __RPC_STUB ICrystalReportViewer_put_EnableNavigationControls_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableStopButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableStopButton);


void __RPC_STUB ICrystalReportViewer_get_EnableStopButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableStopButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableStopButton);


void __RPC_STUB ICrystalReportViewer_put_EnableStopButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnablePrintButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enablePrintButton);


void __RPC_STUB ICrystalReportViewer_get_EnablePrintButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnablePrintButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enablePrintButton);


void __RPC_STUB ICrystalReportViewer_put_EnablePrintButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableZoomControl_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableZoomControl);


void __RPC_STUB ICrystalReportViewer_get_EnableZoomControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableZoomControl_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableZoomControl);


void __RPC_STUB ICrystalReportViewer_put_EnableZoomControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableCloseButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableCloseButton);


void __RPC_STUB ICrystalReportViewer_get_EnableCloseButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableCloseButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableCloseButton);


void __RPC_STUB ICrystalReportViewer_put_EnableCloseButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableProgressControl_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableProgressControl);


void __RPC_STUB ICrystalReportViewer_get_EnableProgressControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableProgressControl_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableProgressControl);


void __RPC_STUB ICrystalReportViewer_put_EnableProgressControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableSearchControl_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSearchControl);


void __RPC_STUB ICrystalReportViewer_get_EnableSearchControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableSearchControl_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableSearchControl);


void __RPC_STUB ICrystalReportViewer_put_EnableSearchControl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableRefreshButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableRefreshButton);


void __RPC_STUB ICrystalReportViewer_get_EnableRefreshButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableRefreshButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableRefreshButton);


void __RPC_STUB ICrystalReportViewer_put_EnableRefreshButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableDrillDown_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableDrillDown);


void __RPC_STUB ICrystalReportViewer_get_EnableDrillDown_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableDrillDown_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableDrillDown);


void __RPC_STUB ICrystalReportViewer_put_EnableDrillDown_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableAnimationCtrl_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableAnimationCtrl);


void __RPC_STUB ICrystalReportViewer_get_EnableAnimationCtrl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableAnimationCtrl_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableAnimationCtrl);


void __RPC_STUB ICrystalReportViewer_put_EnableAnimationCtrl_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableSelectExpertButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSelectExpertButton);


void __RPC_STUB ICrystalReportViewer_get_EnableSelectExpertButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableSelectExpertButton_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL enableSelectExpertButton);


void __RPC_STUB ICrystalReportViewer_put_EnableSelectExpertButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_ViewReport_Proxy( 
    ICrystalReportViewer __RPC_FAR * This);


void __RPC_STUB ICrystalReportViewer_ViewReport_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnableToolbar_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_EnableToolbar_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnableToolbar_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB ICrystalReportViewer_put_EnableToolbar_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_DisplayBorder_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_DisplayBorder_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_DisplayBorder_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB ICrystalReportViewer_put_DisplayBorder_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_DisplayTabs_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_DisplayTabs_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_DisplayTabs_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB ICrystalReportViewer_put_DisplayTabs_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_DisplayBackgroundEdge_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_DisplayBackgroundEdge_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_DisplayBackgroundEdge_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB ICrystalReportViewer_put_DisplayBackgroundEdge_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [hidden][helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_SelectionFormula_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_SelectionFormula_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [hidden][helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_SelectionFormula_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB ICrystalReportViewer_put_SelectionFormula_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_TrackCursorInfo_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ ICRVTrackCursorInfo __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_TrackCursorInfo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_ActiveViewIndex_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_ActiveViewIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_ViewCount_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_ViewCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_ActivateView_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT Index);


void __RPC_STUB ICrystalReportViewer_ActivateView_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_AddView_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT GroupPath);


void __RPC_STUB ICrystalReportViewer_AddView_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_CloseView_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT Index);


void __RPC_STUB ICrystalReportViewer_CloseView_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_GetViewPath_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ short Index,
    /* [retval][out] */ VARIANT __RPC_FAR *Path);


void __RPC_STUB ICrystalReportViewer_GetViewPath_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_PrintReport_Proxy( 
    ICrystalReportViewer __RPC_FAR * This);


void __RPC_STUB ICrystalReportViewer_PrintReport_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_Refresh_Proxy( 
    ICrystalReportViewer __RPC_FAR * This);


void __RPC_STUB ICrystalReportViewer_Refresh_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_SearchForText_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ BSTR Text);


void __RPC_STUB ICrystalReportViewer_SearchForText_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_ShowFirstPage_Proxy( 
    ICrystalReportViewer __RPC_FAR * This);


void __RPC_STUB ICrystalReportViewer_ShowFirstPage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_ShowNextPage_Proxy( 
    ICrystalReportViewer __RPC_FAR * This);


void __RPC_STUB ICrystalReportViewer_ShowNextPage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_ShowPreviousPage_Proxy( 
    ICrystalReportViewer __RPC_FAR * This);


void __RPC_STUB ICrystalReportViewer_ShowPreviousPage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_ShowLastPage_Proxy( 
    ICrystalReportViewer __RPC_FAR * This);


void __RPC_STUB ICrystalReportViewer_ShowLastPage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_ShowNthPage_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ short PageNumber);


void __RPC_STUB ICrystalReportViewer_ShowNthPage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_Zoom_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ short ZoomLevel);


void __RPC_STUB ICrystalReportViewer_Zoom_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_GetCurrentPageNumber_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_GetCurrentPageNumber_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_ShowGroup_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT GroupPath);


void __RPC_STUB ICrystalReportViewer_ShowGroup_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_IsBusy_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_IsBusy_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_get_EnablePopupMenu_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer_get_EnablePopupMenu_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer_put_EnablePopupMenu_Proxy( 
    ICrystalReportViewer __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB ICrystalReportViewer_put_EnablePopupMenu_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportViewer_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportViewer2_INTERFACE_DEFINED__
#define __ICrystalReportViewer2_INTERFACE_DEFINED__

/* interface ICrystalReportViewer2 */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICrystalReportViewer2;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("2550FAE1-20DE-11d2-BF1D-00A0C95A6A5C")
    ICrystalReportViewer2 : public ICrystalReportViewer
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableExportButton( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableExportButton( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableSearchExpertButton( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableSearchExpertButton( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SearchByFormula( 
            BSTR formula) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetViewName( 
            /* [out] */ BSTR __RPC_FAR *pTabName,
            /* [retval][out] */ BSTR __RPC_FAR *pDocName) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportViewer2Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ReportSource )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ReportSource )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ LPUNKNOWN lpUnknown);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayGroupTree )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayGroupTree);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayGroupTree )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL displayGroupTree);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayToolbar )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayToolbar);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayToolbar )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL displayToolbar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableGroupTree )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableGroupTree);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableGroupTree )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableGroupTree);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableNavigationControls )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableNavigationControls);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableNavigationControls )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableNavigationControls);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableStopButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableStopButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableStopButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableStopButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnablePrintButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enablePrintButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnablePrintButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enablePrintButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableZoomControl )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableZoomControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableZoomControl )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableZoomControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableCloseButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableCloseButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableCloseButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableCloseButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableProgressControl )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableProgressControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableProgressControl )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableProgressControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableSearchControl )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSearchControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableSearchControl )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableSearchControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableRefreshButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableRefreshButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableRefreshButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableRefreshButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableDrillDown )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableDrillDown);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableDrillDown )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableDrillDown);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableAnimationCtrl )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableAnimationCtrl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableAnimationCtrl )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableAnimationCtrl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableSelectExpertButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSelectExpertButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableSelectExpertButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableSelectExpertButton);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ViewReport )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableToolbar )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableToolbar )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayBorder )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayBorder )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayTabs )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayTabs )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayBackgroundEdge )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayBackgroundEdge )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [hidden][helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SelectionFormula )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [hidden][helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_SelectionFormula )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TrackCursorInfo )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ ICRVTrackCursorInfo __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ActiveViewIndex )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ViewCount )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ActivateView )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddView )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT GroupPath);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CloseView )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetViewPath )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ short Index,
            /* [retval][out] */ VARIANT __RPC_FAR *Path);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *PrintReport )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Refresh )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SearchForText )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ BSTR Text);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowFirstPage )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowNextPage )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowPreviousPage )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowLastPage )( 
            ICrystalReportViewer2 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowNthPage )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ short PageNumber);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Zoom )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ short ZoomLevel);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCurrentPageNumber )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowGroup )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT GroupPath);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsBusy )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnablePopupMenu )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnablePopupMenu )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableExportButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableExportButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableSearchExpertButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableSearchExpertButton )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SearchByFormula )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            BSTR formula);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetViewName )( 
            ICrystalReportViewer2 __RPC_FAR * This,
            /* [out] */ BSTR __RPC_FAR *pTabName,
            /* [retval][out] */ BSTR __RPC_FAR *pDocName);
        
        END_INTERFACE
    } ICrystalReportViewer2Vtbl;

    interface ICrystalReportViewer2
    {
        CONST_VTBL struct ICrystalReportViewer2Vtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportViewer2_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportViewer2_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportViewer2_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportViewer2_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICrystalReportViewer2_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICrystalReportViewer2_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICrystalReportViewer2_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICrystalReportViewer2_get_ReportSource(This,pVal)	\
    (This)->lpVtbl -> get_ReportSource(This,pVal)

#define ICrystalReportViewer2_put_ReportSource(This,lpUnknown)	\
    (This)->lpVtbl -> put_ReportSource(This,lpUnknown)

#define ICrystalReportViewer2_get_DisplayGroupTree(This,displayGroupTree)	\
    (This)->lpVtbl -> get_DisplayGroupTree(This,displayGroupTree)

#define ICrystalReportViewer2_put_DisplayGroupTree(This,displayGroupTree)	\
    (This)->lpVtbl -> put_DisplayGroupTree(This,displayGroupTree)

#define ICrystalReportViewer2_get_DisplayToolbar(This,displayToolbar)	\
    (This)->lpVtbl -> get_DisplayToolbar(This,displayToolbar)

#define ICrystalReportViewer2_put_DisplayToolbar(This,displayToolbar)	\
    (This)->lpVtbl -> put_DisplayToolbar(This,displayToolbar)

#define ICrystalReportViewer2_get_EnableGroupTree(This,enableGroupTree)	\
    (This)->lpVtbl -> get_EnableGroupTree(This,enableGroupTree)

#define ICrystalReportViewer2_put_EnableGroupTree(This,enableGroupTree)	\
    (This)->lpVtbl -> put_EnableGroupTree(This,enableGroupTree)

#define ICrystalReportViewer2_get_EnableNavigationControls(This,enableNavigationControls)	\
    (This)->lpVtbl -> get_EnableNavigationControls(This,enableNavigationControls)

#define ICrystalReportViewer2_put_EnableNavigationControls(This,enableNavigationControls)	\
    (This)->lpVtbl -> put_EnableNavigationControls(This,enableNavigationControls)

#define ICrystalReportViewer2_get_EnableStopButton(This,enableStopButton)	\
    (This)->lpVtbl -> get_EnableStopButton(This,enableStopButton)

#define ICrystalReportViewer2_put_EnableStopButton(This,enableStopButton)	\
    (This)->lpVtbl -> put_EnableStopButton(This,enableStopButton)

#define ICrystalReportViewer2_get_EnablePrintButton(This,enablePrintButton)	\
    (This)->lpVtbl -> get_EnablePrintButton(This,enablePrintButton)

#define ICrystalReportViewer2_put_EnablePrintButton(This,enablePrintButton)	\
    (This)->lpVtbl -> put_EnablePrintButton(This,enablePrintButton)

#define ICrystalReportViewer2_get_EnableZoomControl(This,enableZoomControl)	\
    (This)->lpVtbl -> get_EnableZoomControl(This,enableZoomControl)

#define ICrystalReportViewer2_put_EnableZoomControl(This,enableZoomControl)	\
    (This)->lpVtbl -> put_EnableZoomControl(This,enableZoomControl)

#define ICrystalReportViewer2_get_EnableCloseButton(This,enableCloseButton)	\
    (This)->lpVtbl -> get_EnableCloseButton(This,enableCloseButton)

#define ICrystalReportViewer2_put_EnableCloseButton(This,enableCloseButton)	\
    (This)->lpVtbl -> put_EnableCloseButton(This,enableCloseButton)

#define ICrystalReportViewer2_get_EnableProgressControl(This,enableProgressControl)	\
    (This)->lpVtbl -> get_EnableProgressControl(This,enableProgressControl)

#define ICrystalReportViewer2_put_EnableProgressControl(This,enableProgressControl)	\
    (This)->lpVtbl -> put_EnableProgressControl(This,enableProgressControl)

#define ICrystalReportViewer2_get_EnableSearchControl(This,enableSearchControl)	\
    (This)->lpVtbl -> get_EnableSearchControl(This,enableSearchControl)

#define ICrystalReportViewer2_put_EnableSearchControl(This,enableSearchControl)	\
    (This)->lpVtbl -> put_EnableSearchControl(This,enableSearchControl)

#define ICrystalReportViewer2_get_EnableRefreshButton(This,enableRefreshButton)	\
    (This)->lpVtbl -> get_EnableRefreshButton(This,enableRefreshButton)

#define ICrystalReportViewer2_put_EnableRefreshButton(This,enableRefreshButton)	\
    (This)->lpVtbl -> put_EnableRefreshButton(This,enableRefreshButton)

#define ICrystalReportViewer2_get_EnableDrillDown(This,enableDrillDown)	\
    (This)->lpVtbl -> get_EnableDrillDown(This,enableDrillDown)

#define ICrystalReportViewer2_put_EnableDrillDown(This,enableDrillDown)	\
    (This)->lpVtbl -> put_EnableDrillDown(This,enableDrillDown)

#define ICrystalReportViewer2_get_EnableAnimationCtrl(This,enableAnimationCtrl)	\
    (This)->lpVtbl -> get_EnableAnimationCtrl(This,enableAnimationCtrl)

#define ICrystalReportViewer2_put_EnableAnimationCtrl(This,enableAnimationCtrl)	\
    (This)->lpVtbl -> put_EnableAnimationCtrl(This,enableAnimationCtrl)

#define ICrystalReportViewer2_get_EnableSelectExpertButton(This,enableSelectExpertButton)	\
    (This)->lpVtbl -> get_EnableSelectExpertButton(This,enableSelectExpertButton)

#define ICrystalReportViewer2_put_EnableSelectExpertButton(This,enableSelectExpertButton)	\
    (This)->lpVtbl -> put_EnableSelectExpertButton(This,enableSelectExpertButton)

#define ICrystalReportViewer2_ViewReport(This)	\
    (This)->lpVtbl -> ViewReport(This)

#define ICrystalReportViewer2_get_EnableToolbar(This,pVal)	\
    (This)->lpVtbl -> get_EnableToolbar(This,pVal)

#define ICrystalReportViewer2_put_EnableToolbar(This,newVal)	\
    (This)->lpVtbl -> put_EnableToolbar(This,newVal)

#define ICrystalReportViewer2_get_DisplayBorder(This,pVal)	\
    (This)->lpVtbl -> get_DisplayBorder(This,pVal)

#define ICrystalReportViewer2_put_DisplayBorder(This,newVal)	\
    (This)->lpVtbl -> put_DisplayBorder(This,newVal)

#define ICrystalReportViewer2_get_DisplayTabs(This,pVal)	\
    (This)->lpVtbl -> get_DisplayTabs(This,pVal)

#define ICrystalReportViewer2_put_DisplayTabs(This,newVal)	\
    (This)->lpVtbl -> put_DisplayTabs(This,newVal)

#define ICrystalReportViewer2_get_DisplayBackgroundEdge(This,pVal)	\
    (This)->lpVtbl -> get_DisplayBackgroundEdge(This,pVal)

#define ICrystalReportViewer2_put_DisplayBackgroundEdge(This,newVal)	\
    (This)->lpVtbl -> put_DisplayBackgroundEdge(This,newVal)

#define ICrystalReportViewer2_get_SelectionFormula(This,pVal)	\
    (This)->lpVtbl -> get_SelectionFormula(This,pVal)

#define ICrystalReportViewer2_put_SelectionFormula(This,newVal)	\
    (This)->lpVtbl -> put_SelectionFormula(This,newVal)

#define ICrystalReportViewer2_get_TrackCursorInfo(This,pVal)	\
    (This)->lpVtbl -> get_TrackCursorInfo(This,pVal)

#define ICrystalReportViewer2_get_ActiveViewIndex(This,pVal)	\
    (This)->lpVtbl -> get_ActiveViewIndex(This,pVal)

#define ICrystalReportViewer2_get_ViewCount(This,pVal)	\
    (This)->lpVtbl -> get_ViewCount(This,pVal)

#define ICrystalReportViewer2_ActivateView(This,Index)	\
    (This)->lpVtbl -> ActivateView(This,Index)

#define ICrystalReportViewer2_AddView(This,GroupPath)	\
    (This)->lpVtbl -> AddView(This,GroupPath)

#define ICrystalReportViewer2_CloseView(This,Index)	\
    (This)->lpVtbl -> CloseView(This,Index)

#define ICrystalReportViewer2_GetViewPath(This,Index,Path)	\
    (This)->lpVtbl -> GetViewPath(This,Index,Path)

#define ICrystalReportViewer2_PrintReport(This)	\
    (This)->lpVtbl -> PrintReport(This)

#define ICrystalReportViewer2_Refresh(This)	\
    (This)->lpVtbl -> Refresh(This)

#define ICrystalReportViewer2_SearchForText(This,Text)	\
    (This)->lpVtbl -> SearchForText(This,Text)

#define ICrystalReportViewer2_ShowFirstPage(This)	\
    (This)->lpVtbl -> ShowFirstPage(This)

#define ICrystalReportViewer2_ShowNextPage(This)	\
    (This)->lpVtbl -> ShowNextPage(This)

#define ICrystalReportViewer2_ShowPreviousPage(This)	\
    (This)->lpVtbl -> ShowPreviousPage(This)

#define ICrystalReportViewer2_ShowLastPage(This)	\
    (This)->lpVtbl -> ShowLastPage(This)

#define ICrystalReportViewer2_ShowNthPage(This,PageNumber)	\
    (This)->lpVtbl -> ShowNthPage(This,PageNumber)

#define ICrystalReportViewer2_Zoom(This,ZoomLevel)	\
    (This)->lpVtbl -> Zoom(This,ZoomLevel)

#define ICrystalReportViewer2_GetCurrentPageNumber(This,pVal)	\
    (This)->lpVtbl -> GetCurrentPageNumber(This,pVal)

#define ICrystalReportViewer2_ShowGroup(This,GroupPath)	\
    (This)->lpVtbl -> ShowGroup(This,GroupPath)

#define ICrystalReportViewer2_get_IsBusy(This,pVal)	\
    (This)->lpVtbl -> get_IsBusy(This,pVal)

#define ICrystalReportViewer2_get_EnablePopupMenu(This,pVal)	\
    (This)->lpVtbl -> get_EnablePopupMenu(This,pVal)

#define ICrystalReportViewer2_put_EnablePopupMenu(This,newVal)	\
    (This)->lpVtbl -> put_EnablePopupMenu(This,newVal)


#define ICrystalReportViewer2_get_EnableExportButton(This,pVal)	\
    (This)->lpVtbl -> get_EnableExportButton(This,pVal)

#define ICrystalReportViewer2_put_EnableExportButton(This,newVal)	\
    (This)->lpVtbl -> put_EnableExportButton(This,newVal)

#define ICrystalReportViewer2_get_EnableSearchExpertButton(This,pVal)	\
    (This)->lpVtbl -> get_EnableSearchExpertButton(This,pVal)

#define ICrystalReportViewer2_put_EnableSearchExpertButton(This,newVal)	\
    (This)->lpVtbl -> put_EnableSearchExpertButton(This,newVal)

#define ICrystalReportViewer2_SearchByFormula(This,formula)	\
    (This)->lpVtbl -> SearchByFormula(This,formula)

#define ICrystalReportViewer2_GetViewName(This,pTabName,pDocName)	\
    (This)->lpVtbl -> GetViewName(This,pTabName,pDocName)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer2_get_EnableExportButton_Proxy( 
    ICrystalReportViewer2 __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer2_get_EnableExportButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer2_put_EnableExportButton_Proxy( 
    ICrystalReportViewer2 __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB ICrystalReportViewer2_put_EnableExportButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer2_get_EnableSearchExpertButton_Proxy( 
    ICrystalReportViewer2 __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer2_get_EnableSearchExpertButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer2_put_EnableSearchExpertButton_Proxy( 
    ICrystalReportViewer2 __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB ICrystalReportViewer2_put_EnableSearchExpertButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer2_SearchByFormula_Proxy( 
    ICrystalReportViewer2 __RPC_FAR * This,
    BSTR formula);


void __RPC_STUB ICrystalReportViewer2_SearchByFormula_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer2_GetViewName_Proxy( 
    ICrystalReportViewer2 __RPC_FAR * This,
    /* [out] */ BSTR __RPC_FAR *pTabName,
    /* [retval][out] */ BSTR __RPC_FAR *pDocName);


void __RPC_STUB ICrystalReportViewer2_GetViewName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportViewer2_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportViewer3_INTERFACE_DEFINED__
#define __ICrystalReportViewer3_INTERFACE_DEFINED__

/* interface ICrystalReportViewer3 */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICrystalReportViewer3;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("911D0D71-17B5-11d3-BFB0-00A0C9DA4FA2")
    ICrystalReportViewer3 : public ICrystalReportViewer2
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableHelpButton( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableHelpButton( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportViewer3Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ReportSource )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ReportSource )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ LPUNKNOWN lpUnknown);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayGroupTree )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayGroupTree);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayGroupTree )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL displayGroupTree);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayToolbar )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *displayToolbar);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayToolbar )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL displayToolbar);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableGroupTree )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableGroupTree);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableGroupTree )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableGroupTree);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableNavigationControls )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableNavigationControls);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableNavigationControls )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableNavigationControls);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableStopButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableStopButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableStopButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableStopButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnablePrintButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enablePrintButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnablePrintButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enablePrintButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableZoomControl )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableZoomControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableZoomControl )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableZoomControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableCloseButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableCloseButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableCloseButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableCloseButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableProgressControl )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableProgressControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableProgressControl )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableProgressControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableSearchControl )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSearchControl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableSearchControl )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableSearchControl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableRefreshButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableRefreshButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableRefreshButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableRefreshButton);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableDrillDown )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableDrillDown);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableDrillDown )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableDrillDown);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableAnimationCtrl )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableAnimationCtrl);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableAnimationCtrl )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableAnimationCtrl);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableSelectExpertButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *enableSelectExpertButton);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableSelectExpertButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL enableSelectExpertButton);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ViewReport )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableToolbar )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableToolbar )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayBorder )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayBorder )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayTabs )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayTabs )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DisplayBackgroundEdge )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DisplayBackgroundEdge )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [hidden][helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SelectionFormula )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [hidden][helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_SelectionFormula )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TrackCursorInfo )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ ICRVTrackCursorInfo __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ActiveViewIndex )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ViewCount )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ActivateView )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddView )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT GroupPath);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CloseView )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT Index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetViewPath )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ short Index,
            /* [retval][out] */ VARIANT __RPC_FAR *Path);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *PrintReport )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Refresh )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SearchForText )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ BSTR Text);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowFirstPage )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowNextPage )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowPreviousPage )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowLastPage )( 
            ICrystalReportViewer3 __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowNthPage )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ short PageNumber);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Zoom )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ short ZoomLevel);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCurrentPageNumber )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ShowGroup )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT GroupPath);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsBusy )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnablePopupMenu )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnablePopupMenu )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableExportButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableExportButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableSearchExpertButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableSearchExpertButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SearchByFormula )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            BSTR formula);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetViewName )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [out] */ BSTR __RPC_FAR *pTabName,
            /* [retval][out] */ BSTR __RPC_FAR *pDocName);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableHelpButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableHelpButton )( 
            ICrystalReportViewer3 __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        END_INTERFACE
    } ICrystalReportViewer3Vtbl;

    interface ICrystalReportViewer3
    {
        CONST_VTBL struct ICrystalReportViewer3Vtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportViewer3_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportViewer3_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportViewer3_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportViewer3_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICrystalReportViewer3_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICrystalReportViewer3_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICrystalReportViewer3_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICrystalReportViewer3_get_ReportSource(This,pVal)	\
    (This)->lpVtbl -> get_ReportSource(This,pVal)

#define ICrystalReportViewer3_put_ReportSource(This,lpUnknown)	\
    (This)->lpVtbl -> put_ReportSource(This,lpUnknown)

#define ICrystalReportViewer3_get_DisplayGroupTree(This,displayGroupTree)	\
    (This)->lpVtbl -> get_DisplayGroupTree(This,displayGroupTree)

#define ICrystalReportViewer3_put_DisplayGroupTree(This,displayGroupTree)	\
    (This)->lpVtbl -> put_DisplayGroupTree(This,displayGroupTree)

#define ICrystalReportViewer3_get_DisplayToolbar(This,displayToolbar)	\
    (This)->lpVtbl -> get_DisplayToolbar(This,displayToolbar)

#define ICrystalReportViewer3_put_DisplayToolbar(This,displayToolbar)	\
    (This)->lpVtbl -> put_DisplayToolbar(This,displayToolbar)

#define ICrystalReportViewer3_get_EnableGroupTree(This,enableGroupTree)	\
    (This)->lpVtbl -> get_EnableGroupTree(This,enableGroupTree)

#define ICrystalReportViewer3_put_EnableGroupTree(This,enableGroupTree)	\
    (This)->lpVtbl -> put_EnableGroupTree(This,enableGroupTree)

#define ICrystalReportViewer3_get_EnableNavigationControls(This,enableNavigationControls)	\
    (This)->lpVtbl -> get_EnableNavigationControls(This,enableNavigationControls)

#define ICrystalReportViewer3_put_EnableNavigationControls(This,enableNavigationControls)	\
    (This)->lpVtbl -> put_EnableNavigationControls(This,enableNavigationControls)

#define ICrystalReportViewer3_get_EnableStopButton(This,enableStopButton)	\
    (This)->lpVtbl -> get_EnableStopButton(This,enableStopButton)

#define ICrystalReportViewer3_put_EnableStopButton(This,enableStopButton)	\
    (This)->lpVtbl -> put_EnableStopButton(This,enableStopButton)

#define ICrystalReportViewer3_get_EnablePrintButton(This,enablePrintButton)	\
    (This)->lpVtbl -> get_EnablePrintButton(This,enablePrintButton)

#define ICrystalReportViewer3_put_EnablePrintButton(This,enablePrintButton)	\
    (This)->lpVtbl -> put_EnablePrintButton(This,enablePrintButton)

#define ICrystalReportViewer3_get_EnableZoomControl(This,enableZoomControl)	\
    (This)->lpVtbl -> get_EnableZoomControl(This,enableZoomControl)

#define ICrystalReportViewer3_put_EnableZoomControl(This,enableZoomControl)	\
    (This)->lpVtbl -> put_EnableZoomControl(This,enableZoomControl)

#define ICrystalReportViewer3_get_EnableCloseButton(This,enableCloseButton)	\
    (This)->lpVtbl -> get_EnableCloseButton(This,enableCloseButton)

#define ICrystalReportViewer3_put_EnableCloseButton(This,enableCloseButton)	\
    (This)->lpVtbl -> put_EnableCloseButton(This,enableCloseButton)

#define ICrystalReportViewer3_get_EnableProgressControl(This,enableProgressControl)	\
    (This)->lpVtbl -> get_EnableProgressControl(This,enableProgressControl)

#define ICrystalReportViewer3_put_EnableProgressControl(This,enableProgressControl)	\
    (This)->lpVtbl -> put_EnableProgressControl(This,enableProgressControl)

#define ICrystalReportViewer3_get_EnableSearchControl(This,enableSearchControl)	\
    (This)->lpVtbl -> get_EnableSearchControl(This,enableSearchControl)

#define ICrystalReportViewer3_put_EnableSearchControl(This,enableSearchControl)	\
    (This)->lpVtbl -> put_EnableSearchControl(This,enableSearchControl)

#define ICrystalReportViewer3_get_EnableRefreshButton(This,enableRefreshButton)	\
    (This)->lpVtbl -> get_EnableRefreshButton(This,enableRefreshButton)

#define ICrystalReportViewer3_put_EnableRefreshButton(This,enableRefreshButton)	\
    (This)->lpVtbl -> put_EnableRefreshButton(This,enableRefreshButton)

#define ICrystalReportViewer3_get_EnableDrillDown(This,enableDrillDown)	\
    (This)->lpVtbl -> get_EnableDrillDown(This,enableDrillDown)

#define ICrystalReportViewer3_put_EnableDrillDown(This,enableDrillDown)	\
    (This)->lpVtbl -> put_EnableDrillDown(This,enableDrillDown)

#define ICrystalReportViewer3_get_EnableAnimationCtrl(This,enableAnimationCtrl)	\
    (This)->lpVtbl -> get_EnableAnimationCtrl(This,enableAnimationCtrl)

#define ICrystalReportViewer3_put_EnableAnimationCtrl(This,enableAnimationCtrl)	\
    (This)->lpVtbl -> put_EnableAnimationCtrl(This,enableAnimationCtrl)

#define ICrystalReportViewer3_get_EnableSelectExpertButton(This,enableSelectExpertButton)	\
    (This)->lpVtbl -> get_EnableSelectExpertButton(This,enableSelectExpertButton)

#define ICrystalReportViewer3_put_EnableSelectExpertButton(This,enableSelectExpertButton)	\
    (This)->lpVtbl -> put_EnableSelectExpertButton(This,enableSelectExpertButton)

#define ICrystalReportViewer3_ViewReport(This)	\
    (This)->lpVtbl -> ViewReport(This)

#define ICrystalReportViewer3_get_EnableToolbar(This,pVal)	\
    (This)->lpVtbl -> get_EnableToolbar(This,pVal)

#define ICrystalReportViewer3_put_EnableToolbar(This,newVal)	\
    (This)->lpVtbl -> put_EnableToolbar(This,newVal)

#define ICrystalReportViewer3_get_DisplayBorder(This,pVal)	\
    (This)->lpVtbl -> get_DisplayBorder(This,pVal)

#define ICrystalReportViewer3_put_DisplayBorder(This,newVal)	\
    (This)->lpVtbl -> put_DisplayBorder(This,newVal)

#define ICrystalReportViewer3_get_DisplayTabs(This,pVal)	\
    (This)->lpVtbl -> get_DisplayTabs(This,pVal)

#define ICrystalReportViewer3_put_DisplayTabs(This,newVal)	\
    (This)->lpVtbl -> put_DisplayTabs(This,newVal)

#define ICrystalReportViewer3_get_DisplayBackgroundEdge(This,pVal)	\
    (This)->lpVtbl -> get_DisplayBackgroundEdge(This,pVal)

#define ICrystalReportViewer3_put_DisplayBackgroundEdge(This,newVal)	\
    (This)->lpVtbl -> put_DisplayBackgroundEdge(This,newVal)

#define ICrystalReportViewer3_get_SelectionFormula(This,pVal)	\
    (This)->lpVtbl -> get_SelectionFormula(This,pVal)

#define ICrystalReportViewer3_put_SelectionFormula(This,newVal)	\
    (This)->lpVtbl -> put_SelectionFormula(This,newVal)

#define ICrystalReportViewer3_get_TrackCursorInfo(This,pVal)	\
    (This)->lpVtbl -> get_TrackCursorInfo(This,pVal)

#define ICrystalReportViewer3_get_ActiveViewIndex(This,pVal)	\
    (This)->lpVtbl -> get_ActiveViewIndex(This,pVal)

#define ICrystalReportViewer3_get_ViewCount(This,pVal)	\
    (This)->lpVtbl -> get_ViewCount(This,pVal)

#define ICrystalReportViewer3_ActivateView(This,Index)	\
    (This)->lpVtbl -> ActivateView(This,Index)

#define ICrystalReportViewer3_AddView(This,GroupPath)	\
    (This)->lpVtbl -> AddView(This,GroupPath)

#define ICrystalReportViewer3_CloseView(This,Index)	\
    (This)->lpVtbl -> CloseView(This,Index)

#define ICrystalReportViewer3_GetViewPath(This,Index,Path)	\
    (This)->lpVtbl -> GetViewPath(This,Index,Path)

#define ICrystalReportViewer3_PrintReport(This)	\
    (This)->lpVtbl -> PrintReport(This)

#define ICrystalReportViewer3_Refresh(This)	\
    (This)->lpVtbl -> Refresh(This)

#define ICrystalReportViewer3_SearchForText(This,Text)	\
    (This)->lpVtbl -> SearchForText(This,Text)

#define ICrystalReportViewer3_ShowFirstPage(This)	\
    (This)->lpVtbl -> ShowFirstPage(This)

#define ICrystalReportViewer3_ShowNextPage(This)	\
    (This)->lpVtbl -> ShowNextPage(This)

#define ICrystalReportViewer3_ShowPreviousPage(This)	\
    (This)->lpVtbl -> ShowPreviousPage(This)

#define ICrystalReportViewer3_ShowLastPage(This)	\
    (This)->lpVtbl -> ShowLastPage(This)

#define ICrystalReportViewer3_ShowNthPage(This,PageNumber)	\
    (This)->lpVtbl -> ShowNthPage(This,PageNumber)

#define ICrystalReportViewer3_Zoom(This,ZoomLevel)	\
    (This)->lpVtbl -> Zoom(This,ZoomLevel)

#define ICrystalReportViewer3_GetCurrentPageNumber(This,pVal)	\
    (This)->lpVtbl -> GetCurrentPageNumber(This,pVal)

#define ICrystalReportViewer3_ShowGroup(This,GroupPath)	\
    (This)->lpVtbl -> ShowGroup(This,GroupPath)

#define ICrystalReportViewer3_get_IsBusy(This,pVal)	\
    (This)->lpVtbl -> get_IsBusy(This,pVal)

#define ICrystalReportViewer3_get_EnablePopupMenu(This,pVal)	\
    (This)->lpVtbl -> get_EnablePopupMenu(This,pVal)

#define ICrystalReportViewer3_put_EnablePopupMenu(This,newVal)	\
    (This)->lpVtbl -> put_EnablePopupMenu(This,newVal)


#define ICrystalReportViewer3_get_EnableExportButton(This,pVal)	\
    (This)->lpVtbl -> get_EnableExportButton(This,pVal)

#define ICrystalReportViewer3_put_EnableExportButton(This,newVal)	\
    (This)->lpVtbl -> put_EnableExportButton(This,newVal)

#define ICrystalReportViewer3_get_EnableSearchExpertButton(This,pVal)	\
    (This)->lpVtbl -> get_EnableSearchExpertButton(This,pVal)

#define ICrystalReportViewer3_put_EnableSearchExpertButton(This,newVal)	\
    (This)->lpVtbl -> put_EnableSearchExpertButton(This,newVal)

#define ICrystalReportViewer3_SearchByFormula(This,formula)	\
    (This)->lpVtbl -> SearchByFormula(This,formula)

#define ICrystalReportViewer3_GetViewName(This,pTabName,pDocName)	\
    (This)->lpVtbl -> GetViewName(This,pTabName,pDocName)


#define ICrystalReportViewer3_get_EnableHelpButton(This,pVal)	\
    (This)->lpVtbl -> get_EnableHelpButton(This,pVal)

#define ICrystalReportViewer3_put_EnableHelpButton(This,newVal)	\
    (This)->lpVtbl -> put_EnableHelpButton(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer3_get_EnableHelpButton_Proxy( 
    ICrystalReportViewer3 __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportViewer3_get_EnableHelpButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE ICrystalReportViewer3_put_EnableHelpButton_Proxy( 
    ICrystalReportViewer3 __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB ICrystalReportViewer3_put_EnableHelpButton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportViewer3_INTERFACE_DEFINED__ */


#ifndef __ICRVEventInfo_INTERFACE_DEFINED__
#define __ICRVEventInfo_INTERFACE_DEFINED__

/* interface ICRVEventInfo */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICRVEventInfo;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("75347085-7260-11D1-BE46-00A0C95A6A5C")
    ICRVEventInfo : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Text( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Index( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ParentIndex( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Type( 
            /* [retval][out] */ CRObjectType __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CanDrillDown( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFields( 
            /* [retval][out] */ VARIANT __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICRVEventInfoVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICRVEventInfo __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICRVEventInfo __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Text )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Index )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ParentIndex )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Type )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [retval][out] */ CRObjectType __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CanDrillDown )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFields )( 
            ICRVEventInfo __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal);
        
        END_INTERFACE
    } ICRVEventInfoVtbl;

    interface ICRVEventInfo
    {
        CONST_VTBL struct ICRVEventInfoVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICRVEventInfo_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICRVEventInfo_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICRVEventInfo_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICRVEventInfo_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICRVEventInfo_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICRVEventInfo_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICRVEventInfo_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICRVEventInfo_get_Text(This,pVal)	\
    (This)->lpVtbl -> get_Text(This,pVal)

#define ICRVEventInfo_get_Index(This,pVal)	\
    (This)->lpVtbl -> get_Index(This,pVal)

#define ICRVEventInfo_get_ParentIndex(This,pVal)	\
    (This)->lpVtbl -> get_ParentIndex(This,pVal)

#define ICRVEventInfo_get_Type(This,pVal)	\
    (This)->lpVtbl -> get_Type(This,pVal)

#define ICRVEventInfo_get_CanDrillDown(This,pVal)	\
    (This)->lpVtbl -> get_CanDrillDown(This,pVal)

#define ICRVEventInfo_GetFields(This,pVal)	\
    (This)->lpVtbl -> GetFields(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVEventInfo_get_Text_Proxy( 
    ICRVEventInfo __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB ICRVEventInfo_get_Text_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVEventInfo_get_Index_Proxy( 
    ICRVEventInfo __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICRVEventInfo_get_Index_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVEventInfo_get_ParentIndex_Proxy( 
    ICRVEventInfo __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICRVEventInfo_get_ParentIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVEventInfo_get_Type_Proxy( 
    ICRVEventInfo __RPC_FAR * This,
    /* [retval][out] */ CRObjectType __RPC_FAR *pVal);


void __RPC_STUB ICRVEventInfo_get_Type_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRVEventInfo_get_CanDrillDown_Proxy( 
    ICRVEventInfo __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICRVEventInfo_get_CanDrillDown_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICRVEventInfo_GetFields_Proxy( 
    ICRVEventInfo __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pVal);


void __RPC_STUB ICRVEventInfo_GetFields_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICRVEventInfo_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportSourceRouter_INTERFACE_DEFINED__
#define __ICrystalReportSourceRouter_INTERFACE_DEFINED__

/* interface ICrystalReportSourceRouter */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICrystalReportSourceRouter;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A0E5F37D-CA67-11D1-A817-00A0C92784CD")
    ICrystalReportSourceRouter : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddReport( 
            IUnknown __RPC_FAR *pUnknown) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportSourceRouterVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportSourceRouter __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportSourceRouter __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportSourceRouter __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICrystalReportSourceRouter __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICrystalReportSourceRouter __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICrystalReportSourceRouter __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICrystalReportSourceRouter __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddReport )( 
            ICrystalReportSourceRouter __RPC_FAR * This,
            IUnknown __RPC_FAR *pUnknown);
        
        END_INTERFACE
    } ICrystalReportSourceRouterVtbl;

    interface ICrystalReportSourceRouter
    {
        CONST_VTBL struct ICrystalReportSourceRouterVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportSourceRouter_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportSourceRouter_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportSourceRouter_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportSourceRouter_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICrystalReportSourceRouter_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICrystalReportSourceRouter_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICrystalReportSourceRouter_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICrystalReportSourceRouter_AddReport(This,pUnknown)	\
    (This)->lpVtbl -> AddReport(This,pUnknown)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICrystalReportSourceRouter_AddReport_Proxy( 
    ICrystalReportSourceRouter __RPC_FAR * This,
    IUnknown __RPC_FAR *pUnknown);


void __RPC_STUB ICrystalReportSourceRouter_AddReport_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportSourceRouter_INTERFACE_DEFINED__ */


#ifndef __ICRFields_INTERFACE_DEFINED__
#define __ICRFields_INTERFACE_DEFINED__

/* interface ICRFields */
/* [unique][helpstring][dual][uuid][public][object] */ 


EXTERN_C const IID IID_ICRFields;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("75C66E65-8949-11D2-BF6D-00A0C9DA4FA2")
    ICRFields : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ long index,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal) = 0;
        
        virtual /* [restricted][helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SelectedFieldIndex( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICRFieldsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICRFields __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICRFields __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICRFields __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICRFields __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICRFields __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICRFields __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICRFields __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            ICRFields __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Item )( 
            ICRFields __RPC_FAR * This,
            /* [in] */ long index,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal);
        
        /* [restricted][helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            ICRFields __RPC_FAR * This,
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SelectedFieldIndex )( 
            ICRFields __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        END_INTERFACE
    } ICRFieldsVtbl;

    interface ICRFields
    {
        CONST_VTBL struct ICRFieldsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICRFields_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICRFields_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICRFields_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICRFields_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICRFields_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICRFields_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICRFields_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICRFields_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define ICRFields_get_Item(This,index,pVal)	\
    (This)->lpVtbl -> get_Item(This,index,pVal)

#define ICRFields_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define ICRFields_get_SelectedFieldIndex(This,pVal)	\
    (This)->lpVtbl -> get_SelectedFieldIndex(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRFields_get_Count_Proxy( 
    ICRFields __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICRFields_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRFields_get_Item_Proxy( 
    ICRFields __RPC_FAR * This,
    /* [in] */ long index,
    /* [retval][out] */ VARIANT __RPC_FAR *pVal);


void __RPC_STUB ICRFields_get_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [restricted][helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRFields_get__NewEnum_Proxy( 
    ICRFields __RPC_FAR * This,
    /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB ICRFields_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRFields_get_SelectedFieldIndex_Proxy( 
    ICRFields __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICRFields_get_SelectedFieldIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICRFields_INTERFACE_DEFINED__ */


#ifndef __ICRField_INTERFACE_DEFINED__
#define __ICRField_INTERFACE_DEFINED__

/* interface ICRField */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICRField;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("75C66E67-8949-11D2-BF6D-00A0C9DA4FA2")
    ICRField : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Value( 
            /* [retval][out] */ VARIANT __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_FieldType( 
            /* [retval][out] */ CRFieldType __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IsRawData( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICRFieldVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICRField __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICRField __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICRField __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICRField __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICRField __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICRField __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICRField __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Value )( 
            ICRField __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_FieldType )( 
            ICRField __RPC_FAR * This,
            /* [retval][out] */ CRFieldType __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Name )( 
            ICRField __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IsRawData )( 
            ICRField __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        END_INTERFACE
    } ICRFieldVtbl;

    interface ICRField
    {
        CONST_VTBL struct ICRFieldVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICRField_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICRField_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICRField_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICRField_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICRField_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICRField_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICRField_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICRField_get_Value(This,pVal)	\
    (This)->lpVtbl -> get_Value(This,pVal)

#define ICRField_get_FieldType(This,pVal)	\
    (This)->lpVtbl -> get_FieldType(This,pVal)

#define ICRField_get_Name(This,pVal)	\
    (This)->lpVtbl -> get_Name(This,pVal)

#define ICRField_get_IsRawData(This,pVal)	\
    (This)->lpVtbl -> get_IsRawData(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRField_get_Value_Proxy( 
    ICRField __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pVal);


void __RPC_STUB ICRField_get_Value_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRField_get_FieldType_Proxy( 
    ICRField __RPC_FAR * This,
    /* [retval][out] */ CRFieldType __RPC_FAR *pVal);


void __RPC_STUB ICRField_get_FieldType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRField_get_Name_Proxy( 
    ICRField __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB ICRField_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE ICRField_get_IsRawData_Proxy( 
    ICRField __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICRField_get_IsRawData_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICRField_INTERFACE_DEFINED__ */


#ifndef __IEnumCRFields_INTERFACE_DEFINED__
#define __IEnumCRFields_INTERFACE_DEFINED__

/* interface IEnumCRFields */
/* [uuid][object] */ 


EXTERN_C const IID IID_IEnumCRFields;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("9E088781-8E06-11d2-BF71-00A0C9DA4FA2")
    IEnumCRFields : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE Next( 
            /* [in] */ ULONG celt,
            /* [length_is][size_is][out] */ ICRField __RPC_FAR *__RPC_FAR *rgptr,
            /* [out] */ ULONG __RPC_FAR *pcFetched) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Skip( 
            /* [in] */ ULONG celt) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Reset( void) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Clone( 
            /* [out] */ IEnumCRFields __RPC_FAR *__RPC_FAR *ppEnum) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IEnumCRFieldsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IEnumCRFields __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IEnumCRFields __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IEnumCRFields __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Next )( 
            IEnumCRFields __RPC_FAR * This,
            /* [in] */ ULONG celt,
            /* [length_is][size_is][out] */ ICRField __RPC_FAR *__RPC_FAR *rgptr,
            /* [out] */ ULONG __RPC_FAR *pcFetched);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Skip )( 
            IEnumCRFields __RPC_FAR * This,
            /* [in] */ ULONG celt);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Reset )( 
            IEnumCRFields __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clone )( 
            IEnumCRFields __RPC_FAR * This,
            /* [out] */ IEnumCRFields __RPC_FAR *__RPC_FAR *ppEnum);
        
        END_INTERFACE
    } IEnumCRFieldsVtbl;

    interface IEnumCRFields
    {
        CONST_VTBL struct IEnumCRFieldsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IEnumCRFields_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IEnumCRFields_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IEnumCRFields_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IEnumCRFields_Next(This,celt,rgptr,pcFetched)	\
    (This)->lpVtbl -> Next(This,celt,rgptr,pcFetched)

#define IEnumCRFields_Skip(This,celt)	\
    (This)->lpVtbl -> Skip(This,celt)

#define IEnumCRFields_Reset(This)	\
    (This)->lpVtbl -> Reset(This)

#define IEnumCRFields_Clone(This,ppEnum)	\
    (This)->lpVtbl -> Clone(This,ppEnum)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IEnumCRFields_Next_Proxy( 
    IEnumCRFields __RPC_FAR * This,
    /* [in] */ ULONG celt,
    /* [length_is][size_is][out] */ ICRField __RPC_FAR *__RPC_FAR *rgptr,
    /* [out] */ ULONG __RPC_FAR *pcFetched);


void __RPC_STUB IEnumCRFields_Next_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IEnumCRFields_Skip_Proxy( 
    IEnumCRFields __RPC_FAR * This,
    /* [in] */ ULONG celt);


void __RPC_STUB IEnumCRFields_Skip_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IEnumCRFields_Reset_Proxy( 
    IEnumCRFields __RPC_FAR * This);


void __RPC_STUB IEnumCRFields_Reset_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IEnumCRFields_Clone_Proxy( 
    IEnumCRFields __RPC_FAR * This,
    /* [out] */ IEnumCRFields __RPC_FAR *__RPC_FAR *ppEnum);


void __RPC_STUB IEnumCRFields_Clone_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IEnumCRFields_INTERFACE_DEFINED__ */


#ifndef __ICollectCRFields_INTERFACE_DEFINED__
#define __ICollectCRFields_INTERFACE_DEFINED__

/* interface ICollectCRFields */
/* [uuid][object] */ 


EXTERN_C const IID IID_ICollectCRFields;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("81CBB971-8E0A-11d2-BF71-00A0C9DA4FA2")
    ICollectCRFields : public IUnknown
    {
    public:
        virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE GetItem( 
            /* [in] */ long index,
            /* [retval][out] */ ICRField __RPC_FAR *__RPC_FAR *ppField) = 0;
        
        virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE GetCount( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE EnumCRFields( 
            /* [retval][out] */ IEnumCRFields __RPC_FAR *__RPC_FAR *ppEnum) = 0;
        
        virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE GetSelectedFieldIndex( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICollectCRFieldsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICollectCRFields __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICollectCRFields __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICollectCRFields __RPC_FAR * This);
        
        /* [helpstring] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetItem )( 
            ICollectCRFields __RPC_FAR * This,
            /* [in] */ long index,
            /* [retval][out] */ ICRField __RPC_FAR *__RPC_FAR *ppField);
        
        /* [helpstring] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetCount )( 
            ICollectCRFields __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *EnumCRFields )( 
            ICollectCRFields __RPC_FAR * This,
            /* [retval][out] */ IEnumCRFields __RPC_FAR *__RPC_FAR *ppEnum);
        
        /* [helpstring] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetSelectedFieldIndex )( 
            ICollectCRFields __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        END_INTERFACE
    } ICollectCRFieldsVtbl;

    interface ICollectCRFields
    {
        CONST_VTBL struct ICollectCRFieldsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICollectCRFields_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICollectCRFields_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICollectCRFields_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICollectCRFields_GetItem(This,index,ppField)	\
    (This)->lpVtbl -> GetItem(This,index,ppField)

#define ICollectCRFields_GetCount(This,pVal)	\
    (This)->lpVtbl -> GetCount(This,pVal)

#define ICollectCRFields_EnumCRFields(This,ppEnum)	\
    (This)->lpVtbl -> EnumCRFields(This,ppEnum)

#define ICollectCRFields_GetSelectedFieldIndex(This,pVal)	\
    (This)->lpVtbl -> GetSelectedFieldIndex(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring] */ HRESULT STDMETHODCALLTYPE ICollectCRFields_GetItem_Proxy( 
    ICollectCRFields __RPC_FAR * This,
    /* [in] */ long index,
    /* [retval][out] */ ICRField __RPC_FAR *__RPC_FAR *ppField);


void __RPC_STUB ICollectCRFields_GetItem_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT STDMETHODCALLTYPE ICollectCRFields_GetCount_Proxy( 
    ICollectCRFields __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollectCRFields_GetCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICollectCRFields_EnumCRFields_Proxy( 
    ICollectCRFields __RPC_FAR * This,
    /* [retval][out] */ IEnumCRFields __RPC_FAR *__RPC_FAR *ppEnum);


void __RPC_STUB ICollectCRFields_EnumCRFields_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring] */ HRESULT STDMETHODCALLTYPE ICollectCRFields_GetSelectedFieldIndex_Proxy( 
    ICollectCRFields __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB ICollectCRFields_GetSelectedFieldIndex_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICollectCRFields_INTERFACE_DEFINED__ */



#ifndef __CRVIEWERLib_LIBRARY_DEFINED__
#define __CRVIEWERLib_LIBRARY_DEFINED__

/* library CRVIEWERLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_CRVIEWERLib;

#ifndef ___ICRViewerEvents_DISPINTERFACE_DEFINED__
#define ___ICRViewerEvents_DISPINTERFACE_DEFINED__

/* dispinterface _ICRViewerEvents */
/* [helpstring][nonextensible][uuid] */ 


EXTERN_C const IID DIID__ICRViewerEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("CFDF4A60-6FFC-11d1-BE46-00A0C95A6A5C")
    _ICRViewerEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _ICRViewerEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _ICRViewerEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _ICRViewerEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _ICRViewerEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _ICRViewerEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _ICRViewerEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _ICRViewerEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _ICRViewerEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _ICRViewerEventsVtbl;

    interface _ICRViewerEvents
    {
        CONST_VTBL struct _ICRViewerEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _ICRViewerEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _ICRViewerEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _ICRViewerEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _ICRViewerEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _ICRViewerEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _ICRViewerEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _ICRViewerEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___ICRViewerEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_CRViewer;

#ifdef __cplusplus

class DECLSPEC_UUID("C4847596-972C-11D0-9567-00A0C9273C2A")
CRViewer;
#endif

EXTERN_C const CLSID CLSID_CRVTrackCursorInfo;

#ifdef __cplusplus

class DECLSPEC_UUID("13FA5947-561C-11D1-BE3F-00A0C95A6A5C")
CRVTrackCursorInfo;
#endif

EXTERN_C const CLSID CLSID_CRFields;

#ifdef __cplusplus

class DECLSPEC_UUID("75C66E66-8949-11D2-BF6D-00A0C9DA4FA2")
CRFields;
#endif

EXTERN_C const CLSID CLSID_CRField;

#ifdef __cplusplus

class DECLSPEC_UUID("75C66E68-8949-11D2-BF6D-00A0C9DA4FA2")
CRField;
#endif

EXTERN_C const CLSID CLSID_CRVEventInfo;

#ifdef __cplusplus

class DECLSPEC_UUID("75347086-7260-11D1-BE46-00A0C95A6A5C")
CRVEventInfo;
#endif

EXTERN_C const CLSID CLSID_ReportSourceRouter;

#ifdef __cplusplus

class DECLSPEC_UUID("A0E5F37E-CA67-11D1-A817-00A0C92784CD")
ReportSourceRouter;
#endif
#endif /* __CRVIEWERLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long __RPC_FAR *, unsigned long            , VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long __RPC_FAR *, VARIANT __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
