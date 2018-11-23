/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Oct 07 02:25:24 1999
 */
/* Compiler settings for crvb60r.idl:
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

#ifndef __craxdrt_h__
#define __craxdrt_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __ICrystalReportSource_FWD_DEFINED__
#define __ICrystalReportSource_FWD_DEFINED__
typedef interface ICrystalReportSource ICrystalReportSource;
#endif 	/* __ICrystalReportSource_FWD_DEFINED__ */


#ifndef __ICrystalReportSourceEx_FWD_DEFINED__
#define __ICrystalReportSourceEx_FWD_DEFINED__
typedef interface ICrystalReportSourceEx ICrystalReportSourceEx;
#endif 	/* __ICrystalReportSourceEx_FWD_DEFINED__ */


#ifndef __ICrystalReportProperties_FWD_DEFINED__
#define __ICrystalReportProperties_FWD_DEFINED__
typedef interface ICrystalReportProperties ICrystalReportProperties;
#endif 	/* __ICrystalReportProperties_FWD_DEFINED__ */


#ifndef __ICrystalReportSourceProperties_FWD_DEFINED__
#define __ICrystalReportSourceProperties_FWD_DEFINED__
typedef interface ICrystalReportSourceProperties ICrystalReportSourceProperties;
#endif 	/* __ICrystalReportSourceProperties_FWD_DEFINED__ */


#ifndef __ICrystalReportPrinterPort_FWD_DEFINED__
#define __ICrystalReportPrinterPort_FWD_DEFINED__
typedef interface ICrystalReportPrinterPort ICrystalReportPrinterPort;
#endif 	/* __ICrystalReportPrinterPort_FWD_DEFINED__ */


#ifndef __ICrystalReportExport_FWD_DEFINED__
#define __ICrystalReportExport_FWD_DEFINED__
typedef interface ICrystalReportExport ICrystalReportExport;
#endif 	/* __ICrystalReportExport_FWD_DEFINED__ */


#ifndef __ICrystalReportSourceEvents_FWD_DEFINED__
#define __ICrystalReportSourceEvents_FWD_DEFINED__
typedef interface ICrystalReportSourceEvents ICrystalReportSourceEvents;
#endif 	/* __ICrystalReportSourceEvents_FWD_DEFINED__ */


#ifndef __ICrystalReportExportEvents_FWD_DEFINED__
#define __ICrystalReportExportEvents_FWD_DEFINED__
typedef interface ICrystalReportExportEvents ICrystalReportExportEvents;
#endif 	/* __ICrystalReportExportEvents_FWD_DEFINED__ */


#ifndef __IOlapGridObject_FWD_DEFINED__
#define __IOlapGridObject_FWD_DEFINED__
typedef interface IOlapGridObject IOlapGridObject;
#endif 	/* __IOlapGridObject_FWD_DEFINED__ */


#ifndef __Report_FWD_DEFINED__
#define __Report_FWD_DEFINED__

#ifdef __cplusplus
typedef class Report Report;
#else
typedef struct Report Report;
#endif /* __cplusplus */

#endif 	/* __Report_FWD_DEFINED__ */


#ifndef __Areas_FWD_DEFINED__
#define __Areas_FWD_DEFINED__

#ifdef __cplusplus
typedef class Areas Areas;
#else
typedef struct Areas Areas;
#endif /* __cplusplus */

#endif 	/* __Areas_FWD_DEFINED__ */


#ifndef __Sections_FWD_DEFINED__
#define __Sections_FWD_DEFINED__

#ifdef __cplusplus
typedef class Sections Sections;
#else
typedef struct Sections Sections;
#endif /* __cplusplus */

#endif 	/* __Sections_FWD_DEFINED__ */


#ifndef __Area_FWD_DEFINED__
#define __Area_FWD_DEFINED__

#ifdef __cplusplus
typedef class Area Area;
#else
typedef struct Area Area;
#endif /* __cplusplus */

#endif 	/* __Area_FWD_DEFINED__ */


#ifndef __Section_FWD_DEFINED__
#define __Section_FWD_DEFINED__

#ifdef __cplusplus
typedef class Section Section;
#else
typedef struct Section Section;
#endif /* __cplusplus */

#endif 	/* __Section_FWD_DEFINED__ */


#ifndef __ReportObjects_FWD_DEFINED__
#define __ReportObjects_FWD_DEFINED__

#ifdef __cplusplus
typedef class ReportObjects ReportObjects;
#else
typedef struct ReportObjects ReportObjects;
#endif /* __cplusplus */

#endif 	/* __ReportObjects_FWD_DEFINED__ */


#ifndef __FieldObject_FWD_DEFINED__
#define __FieldObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class FieldObject FieldObject;
#else
typedef struct FieldObject FieldObject;
#endif /* __cplusplus */

#endif 	/* __FieldObject_FWD_DEFINED__ */


#ifndef __TextObject_FWD_DEFINED__
#define __TextObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class TextObject TextObject;
#else
typedef struct TextObject TextObject;
#endif /* __cplusplus */

#endif 	/* __TextObject_FWD_DEFINED__ */


#ifndef __SubreportObject_FWD_DEFINED__
#define __SubreportObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class SubreportObject SubreportObject;
#else
typedef struct SubreportObject SubreportObject;
#endif /* __cplusplus */

#endif 	/* __SubreportObject_FWD_DEFINED__ */


#ifndef __DatabaseFieldDefinition_FWD_DEFINED__
#define __DatabaseFieldDefinition_FWD_DEFINED__

#ifdef __cplusplus
typedef class DatabaseFieldDefinition DatabaseFieldDefinition;
#else
typedef struct DatabaseFieldDefinition DatabaseFieldDefinition;
#endif /* __cplusplus */

#endif 	/* __DatabaseFieldDefinition_FWD_DEFINED__ */


#ifndef __FormulaFieldDefinition_FWD_DEFINED__
#define __FormulaFieldDefinition_FWD_DEFINED__

#ifdef __cplusplus
typedef class FormulaFieldDefinition FormulaFieldDefinition;
#else
typedef struct FormulaFieldDefinition FormulaFieldDefinition;
#endif /* __cplusplus */

#endif 	/* __FormulaFieldDefinition_FWD_DEFINED__ */


#ifndef __ParameterFieldDefinition_FWD_DEFINED__
#define __ParameterFieldDefinition_FWD_DEFINED__

#ifdef __cplusplus
typedef class ParameterFieldDefinition ParameterFieldDefinition;
#else
typedef struct ParameterFieldDefinition ParameterFieldDefinition;
#endif /* __cplusplus */

#endif 	/* __ParameterFieldDefinition_FWD_DEFINED__ */


#ifndef __GroupNameFieldDefinition_FWD_DEFINED__
#define __GroupNameFieldDefinition_FWD_DEFINED__

#ifdef __cplusplus
typedef class GroupNameFieldDefinition GroupNameFieldDefinition;
#else
typedef struct GroupNameFieldDefinition GroupNameFieldDefinition;
#endif /* __cplusplus */

#endif 	/* __GroupNameFieldDefinition_FWD_DEFINED__ */


#ifndef __SpecialVarFieldDefinition_FWD_DEFINED__
#define __SpecialVarFieldDefinition_FWD_DEFINED__

#ifdef __cplusplus
typedef class SpecialVarFieldDefinition SpecialVarFieldDefinition;
#else
typedef struct SpecialVarFieldDefinition SpecialVarFieldDefinition;
#endif /* __cplusplus */

#endif 	/* __SpecialVarFieldDefinition_FWD_DEFINED__ */


#ifndef __SummaryFieldDefinition_FWD_DEFINED__
#define __SummaryFieldDefinition_FWD_DEFINED__

#ifdef __cplusplus
typedef class SummaryFieldDefinition SummaryFieldDefinition;
#else
typedef struct SummaryFieldDefinition SummaryFieldDefinition;
#endif /* __cplusplus */

#endif 	/* __SummaryFieldDefinition_FWD_DEFINED__ */


#ifndef __RunningTotalFieldDefinition_FWD_DEFINED__
#define __RunningTotalFieldDefinition_FWD_DEFINED__

#ifdef __cplusplus
typedef class RunningTotalFieldDefinition RunningTotalFieldDefinition;
#else
typedef struct RunningTotalFieldDefinition RunningTotalFieldDefinition;
#endif /* __cplusplus */

#endif 	/* __RunningTotalFieldDefinition_FWD_DEFINED__ */


#ifndef __SQLExpressionFieldDefinition_FWD_DEFINED__
#define __SQLExpressionFieldDefinition_FWD_DEFINED__

#ifdef __cplusplus
typedef class SQLExpressionFieldDefinition SQLExpressionFieldDefinition;
#else
typedef struct SQLExpressionFieldDefinition SQLExpressionFieldDefinition;
#endif /* __cplusplus */

#endif 	/* __SQLExpressionFieldDefinition_FWD_DEFINED__ */


#ifndef __Database_FWD_DEFINED__
#define __Database_FWD_DEFINED__

#ifdef __cplusplus
typedef class Database Database;
#else
typedef struct Database Database;
#endif /* __cplusplus */

#endif 	/* __Database_FWD_DEFINED__ */


#ifndef __DatabaseTables_FWD_DEFINED__
#define __DatabaseTables_FWD_DEFINED__

#ifdef __cplusplus
typedef class DatabaseTables DatabaseTables;
#else
typedef struct DatabaseTables DatabaseTables;
#endif /* __cplusplus */

#endif 	/* __DatabaseTables_FWD_DEFINED__ */


#ifndef __DatabaseTable_FWD_DEFINED__
#define __DatabaseTable_FWD_DEFINED__

#ifdef __cplusplus
typedef class DatabaseTable DatabaseTable;
#else
typedef struct DatabaseTable DatabaseTable;
#endif /* __cplusplus */

#endif 	/* __DatabaseTable_FWD_DEFINED__ */


#ifndef __DatabaseFieldDefinitions_FWD_DEFINED__
#define __DatabaseFieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class DatabaseFieldDefinitions DatabaseFieldDefinitions;
#else
typedef struct DatabaseFieldDefinitions DatabaseFieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __DatabaseFieldDefinitions_FWD_DEFINED__ */


#ifndef __FormulaFieldDefinitions_FWD_DEFINED__
#define __FormulaFieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class FormulaFieldDefinitions FormulaFieldDefinitions;
#else
typedef struct FormulaFieldDefinitions FormulaFieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __FormulaFieldDefinitions_FWD_DEFINED__ */


#ifndef __ParameterFieldDefinitions_FWD_DEFINED__
#define __ParameterFieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class ParameterFieldDefinitions ParameterFieldDefinitions;
#else
typedef struct ParameterFieldDefinitions ParameterFieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __ParameterFieldDefinitions_FWD_DEFINED__ */


#ifndef __GroupNameFieldDefinitions_FWD_DEFINED__
#define __GroupNameFieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class GroupNameFieldDefinitions GroupNameFieldDefinitions;
#else
typedef struct GroupNameFieldDefinitions GroupNameFieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __GroupNameFieldDefinitions_FWD_DEFINED__ */


#ifndef __SummaryFieldDefinitions_FWD_DEFINED__
#define __SummaryFieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class SummaryFieldDefinitions SummaryFieldDefinitions;
#else
typedef struct SummaryFieldDefinitions SummaryFieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __SummaryFieldDefinitions_FWD_DEFINED__ */


#ifndef __RunningTotalFieldDefinitions_FWD_DEFINED__
#define __RunningTotalFieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class RunningTotalFieldDefinitions RunningTotalFieldDefinitions;
#else
typedef struct RunningTotalFieldDefinitions RunningTotalFieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __RunningTotalFieldDefinitions_FWD_DEFINED__ */


#ifndef __SQLExpressionFieldDefinitions_FWD_DEFINED__
#define __SQLExpressionFieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class SQLExpressionFieldDefinitions SQLExpressionFieldDefinitions;
#else
typedef struct SQLExpressionFieldDefinitions SQLExpressionFieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __SQLExpressionFieldDefinitions_FWD_DEFINED__ */


#ifndef __GraphObject_FWD_DEFINED__
#define __GraphObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class GraphObject GraphObject;
#else
typedef struct GraphObject GraphObject;
#endif /* __cplusplus */

#endif 	/* __GraphObject_FWD_DEFINED__ */


#ifndef __MapObject_FWD_DEFINED__
#define __MapObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class MapObject MapObject;
#else
typedef struct MapObject MapObject;
#endif /* __cplusplus */

#endif 	/* __MapObject_FWD_DEFINED__ */


#ifndef __OleObject_FWD_DEFINED__
#define __OleObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class OleObject OleObject;
#else
typedef struct OleObject OleObject;
#endif /* __cplusplus */

#endif 	/* __OleObject_FWD_DEFINED__ */


#ifndef __BlobFieldObject_FWD_DEFINED__
#define __BlobFieldObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class BlobFieldObject BlobFieldObject;
#else
typedef struct BlobFieldObject BlobFieldObject;
#endif /* __cplusplus */

#endif 	/* __BlobFieldObject_FWD_DEFINED__ */


#ifndef __LineObject_FWD_DEFINED__
#define __LineObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class LineObject LineObject;
#else
typedef struct LineObject LineObject;
#endif /* __cplusplus */

#endif 	/* __LineObject_FWD_DEFINED__ */


#ifndef __BoxObject_FWD_DEFINED__
#define __BoxObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class BoxObject BoxObject;
#else
typedef struct BoxObject BoxObject;
#endif /* __cplusplus */

#endif 	/* __BoxObject_FWD_DEFINED__ */


#ifndef __OlapGridObject_FWD_DEFINED__
#define __OlapGridObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class OlapGridObject OlapGridObject;
#else
typedef struct OlapGridObject OlapGridObject;
#endif /* __cplusplus */

#endif 	/* __OlapGridObject_FWD_DEFINED__ */


#ifndef __CrossTabObject_FWD_DEFINED__
#define __CrossTabObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class CrossTabObject CrossTabObject;
#else
typedef struct CrossTabObject CrossTabObject;
#endif /* __cplusplus */

#endif 	/* __CrossTabObject_FWD_DEFINED__ */


#ifndef __PageEngine_FWD_DEFINED__
#define __PageEngine_FWD_DEFINED__

#ifdef __cplusplus
typedef class PageEngine PageEngine;
#else
typedef struct PageEngine PageEngine;
#endif /* __cplusplus */

#endif 	/* __PageEngine_FWD_DEFINED__ */


#ifndef __PageGenerator_FWD_DEFINED__
#define __PageGenerator_FWD_DEFINED__

#ifdef __cplusplus
typedef class PageGenerator PageGenerator;
#else
typedef struct PageGenerator PageGenerator;
#endif /* __cplusplus */

#endif 	/* __PageGenerator_FWD_DEFINED__ */


#ifndef __Pages_FWD_DEFINED__
#define __Pages_FWD_DEFINED__

#ifdef __cplusplus
typedef class Pages Pages;
#else
typedef struct Pages Pages;
#endif /* __cplusplus */

#endif 	/* __Pages_FWD_DEFINED__ */


#ifndef __Page_FWD_DEFINED__
#define __Page_FWD_DEFINED__

#ifdef __cplusplus
typedef class Page Page;
#else
typedef struct Page Page;
#endif /* __cplusplus */

#endif 	/* __Page_FWD_DEFINED__ */


#ifndef __ExportOptions_FWD_DEFINED__
#define __ExportOptions_FWD_DEFINED__

#ifdef __cplusplus
typedef class ExportOptions ExportOptions;
#else
typedef struct ExportOptions ExportOptions;
#endif /* __cplusplus */

#endif 	/* __ExportOptions_FWD_DEFINED__ */


#ifndef __Application_FWD_DEFINED__
#define __Application_FWD_DEFINED__

#ifdef __cplusplus
typedef class Application Application;
#else
typedef struct Application Application;
#endif /* __cplusplus */

#endif 	/* __Application_FWD_DEFINED__ */


#ifndef __FormattingInfo_FWD_DEFINED__
#define __FormattingInfo_FWD_DEFINED__

#ifdef __cplusplus
typedef class FormattingInfo FormattingInfo;
#else
typedef struct FormattingInfo FormattingInfo;
#endif /* __cplusplus */

#endif 	/* __FormattingInfo_FWD_DEFINED__ */


#ifndef __SortFields_FWD_DEFINED__
#define __SortFields_FWD_DEFINED__

#ifdef __cplusplus
typedef class SortFields SortFields;
#else
typedef struct SortFields SortFields;
#endif /* __cplusplus */

#endif 	/* __SortFields_FWD_DEFINED__ */


#ifndef __SortField_FWD_DEFINED__
#define __SortField_FWD_DEFINED__

#ifdef __cplusplus
typedef class SortField SortField;
#else
typedef struct SortField SortField;
#endif /* __cplusplus */

#endif 	/* __SortField_FWD_DEFINED__ */


#ifndef __PrintingStatus_FWD_DEFINED__
#define __PrintingStatus_FWD_DEFINED__

#ifdef __cplusplus
typedef class PrintingStatus PrintingStatus;
#else
typedef struct PrintingStatus PrintingStatus;
#endif /* __cplusplus */

#endif 	/* __PrintingStatus_FWD_DEFINED__ */


#ifndef __SubreportLink_FWD_DEFINED__
#define __SubreportLink_FWD_DEFINED__

#ifdef __cplusplus
typedef class SubreportLink SubreportLink;
#else
typedef struct SubreportLink SubreportLink;
#endif /* __cplusplus */

#endif 	/* __SubreportLink_FWD_DEFINED__ */


#ifndef __SubreportLinks_FWD_DEFINED__
#define __SubreportLinks_FWD_DEFINED__

#ifdef __cplusplus
typedef class SubreportLinks SubreportLinks;
#else
typedef struct SubreportLinks SubreportLinks;
#endif /* __cplusplus */

#endif 	/* __SubreportLinks_FWD_DEFINED__ */


#ifndef __CrossTabGroups_FWD_DEFINED__
#define __CrossTabGroups_FWD_DEFINED__

#ifdef __cplusplus
typedef class CrossTabGroups CrossTabGroups;
#else
typedef struct CrossTabGroups CrossTabGroups;
#endif /* __cplusplus */

#endif 	/* __CrossTabGroups_FWD_DEFINED__ */


#ifndef __CrossTabGroup_FWD_DEFINED__
#define __CrossTabGroup_FWD_DEFINED__

#ifdef __cplusplus
typedef class CrossTabGroup CrossTabGroup;
#else
typedef struct CrossTabGroup CrossTabGroup;
#endif /* __cplusplus */

#endif 	/* __CrossTabGroup_FWD_DEFINED__ */


#ifndef __FieldDefinitions_FWD_DEFINED__
#define __FieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class FieldDefinitions FieldDefinitions;
#else
typedef struct FieldDefinitions FieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __FieldDefinitions_FWD_DEFINED__ */


#ifndef __ObjectSummaryFieldDefinitions_FWD_DEFINED__
#define __ObjectSummaryFieldDefinitions_FWD_DEFINED__

#ifdef __cplusplus
typedef class ObjectSummaryFieldDefinitions ObjectSummaryFieldDefinitions;
#else
typedef struct ObjectSummaryFieldDefinitions ObjectSummaryFieldDefinitions;
#endif /* __cplusplus */

#endif 	/* __ObjectSummaryFieldDefinitions_FWD_DEFINED__ */


#ifndef __TableLink_FWD_DEFINED__
#define __TableLink_FWD_DEFINED__

#ifdef __cplusplus
typedef class TableLink TableLink;
#else
typedef struct TableLink TableLink;
#endif /* __cplusplus */

#endif 	/* __TableLink_FWD_DEFINED__ */


#ifndef __TableLinks_FWD_DEFINED__
#define __TableLinks_FWD_DEFINED__

#ifdef __cplusplus
typedef class TableLinks TableLinks;
#else
typedef struct TableLinks TableLinks;
#endif /* __cplusplus */

#endif 	/* __TableLinks_FWD_DEFINED__ */


#ifndef __FieldMappingData_FWD_DEFINED__
#define __FieldMappingData_FWD_DEFINED__

#ifdef __cplusplus
typedef class FieldMappingData FieldMappingData;
#else
typedef struct FieldMappingData FieldMappingData;
#endif /* __cplusplus */

#endif 	/* __FieldMappingData_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "craxdi.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

/* interface __MIDL_itf_crvb60r_0000 */
/* [local] */ 

typedef /* [hidden] */ 
enum CRReportSourceDataType
    {	DT_UNKNOWN	= 0,
	DT_FLAT	= 1,
	DT_NUMBER	= 2,
	DT_GROUP	= 3,
	DT_ERROR	= 4,
	DT_SEARCH	= 5,
	DT_FILE	= 6,
	DT_PROMPT	= 7,
	DT_EXPORT_FORMAT	= DT_PROMPT + 1
    }	CRReportSourceDataType;



extern RPC_IF_HANDLE __MIDL_itf_crvb60r_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_crvb60r_0000_v0_0_s_ifspec;

#ifndef __ICrystalReportSource_INTERFACE_DEFINED__
#define __ICrystalReportSource_INTERFACE_DEFINED__

/* interface ICrystalReportSource */
/* [object][nonextensible][hidden][oleautomation][uuid] */ 


EXTERN_C const IID IID_ICrystalReportSource;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("3DCC8FB4-C434-11d1-A817-00A0C92784CD")
    ICrystalReportSource : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE GetPage( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ long lPageNumber,
            /* [in] */ short drillDownLevel) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetTotaller( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ BSTR bstrRootGroup,
            /* [in] */ long lStartFrom,
            /* [in] */ short nLevelsPastRoot,
            /* [in] */ VARIANT vtMaxNodeCount) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetLastPageNumber( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ short drillDownLevel) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE FindGroup( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ BSTR bstrGroupPath,
            /* [in] */ short drillDownLevel) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE FindText( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ long lFromPage,
            /* [in] */ long lFromInstance,
            /* [in] */ BSTR bstrText,
            /* [in] */ CRSearchDirection nMode,
            /* [in] */ short drillDownLevel) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE DrillGraph( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ long lPageNumber,
            /* [in] */ long xOffset,
            /* [in] */ long yOffset,
            /* [in] */ short drillDownLevel) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Refresh( 
            /* [in] */ long lCookie) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Cancel( 
            /* [in] */ long lCookie) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportSourceVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportSource __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportSource __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPage )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ long lPageNumber,
            /* [in] */ short drillDownLevel);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTotaller )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ BSTR bstrRootGroup,
            /* [in] */ long lStartFrom,
            /* [in] */ short nLevelsPastRoot,
            /* [in] */ VARIANT vtMaxNodeCount);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetLastPageNumber )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ short drillDownLevel);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FindGroup )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ BSTR bstrGroupPath,
            /* [in] */ short drillDownLevel);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FindText )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ long lFromPage,
            /* [in] */ long lFromInstance,
            /* [in] */ BSTR bstrText,
            /* [in] */ CRSearchDirection nMode,
            /* [in] */ short drillDownLevel);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DrillGraph )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrBranch,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ long lPageNumber,
            /* [in] */ long xOffset,
            /* [in] */ long yOffset,
            /* [in] */ short drillDownLevel);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Refresh )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ long lCookie);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Cancel )( 
            ICrystalReportSource __RPC_FAR * This,
            /* [in] */ long lCookie);
        
        END_INTERFACE
    } ICrystalReportSourceVtbl;

    interface ICrystalReportSource
    {
        CONST_VTBL struct ICrystalReportSourceVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportSource_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportSource_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportSource_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportSource_GetPage(This,lCookie,bstrBranch,bstrFormula,lPageNumber,drillDownLevel)	\
    (This)->lpVtbl -> GetPage(This,lCookie,bstrBranch,bstrFormula,lPageNumber,drillDownLevel)

#define ICrystalReportSource_GetTotaller(This,lCookie,bstrBranch,bstrFormula,bstrRootGroup,lStartFrom,nLevelsPastRoot,vtMaxNodeCount)	\
    (This)->lpVtbl -> GetTotaller(This,lCookie,bstrBranch,bstrFormula,bstrRootGroup,lStartFrom,nLevelsPastRoot,vtMaxNodeCount)

#define ICrystalReportSource_GetLastPageNumber(This,lCookie,bstrBranch,bstrFormula,drillDownLevel)	\
    (This)->lpVtbl -> GetLastPageNumber(This,lCookie,bstrBranch,bstrFormula,drillDownLevel)

#define ICrystalReportSource_FindGroup(This,lCookie,bstrBranch,bstrFormula,bstrGroupPath,drillDownLevel)	\
    (This)->lpVtbl -> FindGroup(This,lCookie,bstrBranch,bstrFormula,bstrGroupPath,drillDownLevel)

#define ICrystalReportSource_FindText(This,lCookie,bstrBranch,bstrFormula,lFromPage,lFromInstance,bstrText,nMode,drillDownLevel)	\
    (This)->lpVtbl -> FindText(This,lCookie,bstrBranch,bstrFormula,lFromPage,lFromInstance,bstrText,nMode,drillDownLevel)

#define ICrystalReportSource_DrillGraph(This,lCookie,bstrBranch,bstrFormula,lPageNumber,xOffset,yOffset,drillDownLevel)	\
    (This)->lpVtbl -> DrillGraph(This,lCookie,bstrBranch,bstrFormula,lPageNumber,xOffset,yOffset,drillDownLevel)

#define ICrystalReportSource_Refresh(This,lCookie)	\
    (This)->lpVtbl -> Refresh(This,lCookie)

#define ICrystalReportSource_Cancel(This,lCookie)	\
    (This)->lpVtbl -> Cancel(This,lCookie)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICrystalReportSource_GetPage_Proxy( 
    ICrystalReportSource __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrBranch,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ long lPageNumber,
    /* [in] */ short drillDownLevel);


void __RPC_STUB ICrystalReportSource_GetPage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSource_GetTotaller_Proxy( 
    ICrystalReportSource __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrBranch,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ BSTR bstrRootGroup,
    /* [in] */ long lStartFrom,
    /* [in] */ short nLevelsPastRoot,
    /* [in] */ VARIANT vtMaxNodeCount);


void __RPC_STUB ICrystalReportSource_GetTotaller_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSource_GetLastPageNumber_Proxy( 
    ICrystalReportSource __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrBranch,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ short drillDownLevel);


void __RPC_STUB ICrystalReportSource_GetLastPageNumber_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSource_FindGroup_Proxy( 
    ICrystalReportSource __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrBranch,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ BSTR bstrGroupPath,
    /* [in] */ short drillDownLevel);


void __RPC_STUB ICrystalReportSource_FindGroup_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSource_FindText_Proxy( 
    ICrystalReportSource __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrBranch,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ long lFromPage,
    /* [in] */ long lFromInstance,
    /* [in] */ BSTR bstrText,
    /* [in] */ CRSearchDirection nMode,
    /* [in] */ short drillDownLevel);


void __RPC_STUB ICrystalReportSource_FindText_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSource_DrillGraph_Proxy( 
    ICrystalReportSource __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrBranch,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ long lPageNumber,
    /* [in] */ long xOffset,
    /* [in] */ long yOffset,
    /* [in] */ short drillDownLevel);


void __RPC_STUB ICrystalReportSource_DrillGraph_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSource_Refresh_Proxy( 
    ICrystalReportSource __RPC_FAR * This,
    /* [in] */ long lCookie);


void __RPC_STUB ICrystalReportSource_Refresh_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSource_Cancel_Proxy( 
    ICrystalReportSource __RPC_FAR * This,
    /* [in] */ long lCookie);


void __RPC_STUB ICrystalReportSource_Cancel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportSource_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportSourceEx_INTERFACE_DEFINED__
#define __ICrystalReportSourceEx_INTERFACE_DEFINED__

/* interface ICrystalReportSourceEx */
/* [object][nonextensible][hidden][oleautomation][uuid] */ 


EXTERN_C const IID IID_ICrystalReportSourceEx;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("3DCC8FB6-C434-11d1-A817-00A0C92784CD")
    ICrystalReportSourceEx : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE GetPage( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lPageNumber,
            /* [in] */ VARIANT vtReserved) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetTotaller( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lStartFrom,
            /* [in] */ short nLevelsPastRoot,
            /* [in] */ VARIANT vtMaxNodeCount) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetLastPageNumber( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ VARIANT vtReserved) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE FindGroup( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ BSTR bstrGroupPath,
            /* [in] */ VARIANT vtReserved) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE FindText( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lFromPage,
            /* [in] */ long lFromInstance,
            /* [in] */ BSTR bstrText,
            /* [in] */ CRSearchDirection nMode,
            /* [in] */ VARIANT vtReserved) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE DrillGraph( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lPageNumber,
            /* [in] */ long xOffset,
            /* [in] */ long yOffset,
            /* [in] */ BSTR bstrReserved,
            /* [in] */ VARIANT vtReserved,
            /* [in] */ VARIANT vtReserved2) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE DrillMap( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lPageNumber,
            /* [in] */ long xOffset,
            /* [in] */ long yOffset,
            /* [in] */ BSTR bstrReserved,
            /* [in] */ VARIANT vtReserved,
            /* [in] */ VARIANT vtReserved2) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Search( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lPageN,
            /* [in] */ long lSectionInstN,
            /* [in] */ BSTR bstrSearchFormula,
            /* [in] */ BSTR bstrReserved,
            /* [in] */ VARIANT vtReserved) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Export( 
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ VARIANT exportFormat,
            /* [in] */ VARIANT vtReserved) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetExportFormats( 
            /* [in] */ long lCookie) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Refresh( 
            /* [in] */ long lCookie,
            /* [in] */ VARIANT vtPromptingInfo) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Cancel( 
            /* [in] */ long lCookie) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportSourceExVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportSourceEx __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportSourceEx __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPage )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lPageNumber,
            /* [in] */ VARIANT vtReserved);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTotaller )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lStartFrom,
            /* [in] */ short nLevelsPastRoot,
            /* [in] */ VARIANT vtMaxNodeCount);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetLastPageNumber )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ VARIANT vtReserved);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FindGroup )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ BSTR bstrGroupPath,
            /* [in] */ VARIANT vtReserved);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FindText )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lFromPage,
            /* [in] */ long lFromInstance,
            /* [in] */ BSTR bstrText,
            /* [in] */ CRSearchDirection nMode,
            /* [in] */ VARIANT vtReserved);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DrillGraph )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lPageNumber,
            /* [in] */ long xOffset,
            /* [in] */ long yOffset,
            /* [in] */ BSTR bstrReserved,
            /* [in] */ VARIANT vtReserved,
            /* [in] */ VARIANT vtReserved2);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DrillMap )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lPageNumber,
            /* [in] */ long xOffset,
            /* [in] */ long yOffset,
            /* [in] */ BSTR bstrReserved,
            /* [in] */ VARIANT vtReserved,
            /* [in] */ VARIANT vtReserved2);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Search )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ long lPageN,
            /* [in] */ long lSectionInstN,
            /* [in] */ BSTR bstrSearchFormula,
            /* [in] */ BSTR bstrReserved,
            /* [in] */ VARIANT vtReserved);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Export )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ BSTR bstrViewContext,
            /* [in] */ BSTR bstrSubreportContext,
            /* [in] */ BSTR bstrFormula,
            /* [in] */ VARIANT vtPromptingInfo,
            /* [in] */ VARIANT exportFormat,
            /* [in] */ VARIANT vtReserved);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetExportFormats )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Refresh )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie,
            /* [in] */ VARIANT vtPromptingInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Cancel )( 
            ICrystalReportSourceEx __RPC_FAR * This,
            /* [in] */ long lCookie);
        
        END_INTERFACE
    } ICrystalReportSourceExVtbl;

    interface ICrystalReportSourceEx
    {
        CONST_VTBL struct ICrystalReportSourceExVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportSourceEx_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportSourceEx_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportSourceEx_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportSourceEx_GetPage(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lPageNumber,vtReserved)	\
    (This)->lpVtbl -> GetPage(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lPageNumber,vtReserved)

#define ICrystalReportSourceEx_GetTotaller(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lStartFrom,nLevelsPastRoot,vtMaxNodeCount)	\
    (This)->lpVtbl -> GetTotaller(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lStartFrom,nLevelsPastRoot,vtMaxNodeCount)

#define ICrystalReportSourceEx_GetLastPageNumber(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,vtReserved)	\
    (This)->lpVtbl -> GetLastPageNumber(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,vtReserved)

#define ICrystalReportSourceEx_FindGroup(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,bstrGroupPath,vtReserved)	\
    (This)->lpVtbl -> FindGroup(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,bstrGroupPath,vtReserved)

#define ICrystalReportSourceEx_FindText(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lFromPage,lFromInstance,bstrText,nMode,vtReserved)	\
    (This)->lpVtbl -> FindText(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lFromPage,lFromInstance,bstrText,nMode,vtReserved)

#define ICrystalReportSourceEx_DrillGraph(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lPageNumber,xOffset,yOffset,bstrReserved,vtReserved,vtReserved2)	\
    (This)->lpVtbl -> DrillGraph(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lPageNumber,xOffset,yOffset,bstrReserved,vtReserved,vtReserved2)

#define ICrystalReportSourceEx_DrillMap(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lPageNumber,xOffset,yOffset,bstrReserved,vtReserved,vtReserved2)	\
    (This)->lpVtbl -> DrillMap(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lPageNumber,xOffset,yOffset,bstrReserved,vtReserved,vtReserved2)

#define ICrystalReportSourceEx_Search(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lPageN,lSectionInstN,bstrSearchFormula,bstrReserved,vtReserved)	\
    (This)->lpVtbl -> Search(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,lPageN,lSectionInstN,bstrSearchFormula,bstrReserved,vtReserved)

#define ICrystalReportSourceEx_Export(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,exportFormat,vtReserved)	\
    (This)->lpVtbl -> Export(This,lCookie,bstrViewContext,bstrSubreportContext,bstrFormula,vtPromptingInfo,exportFormat,vtReserved)

#define ICrystalReportSourceEx_GetExportFormats(This,lCookie)	\
    (This)->lpVtbl -> GetExportFormats(This,lCookie)

#define ICrystalReportSourceEx_Refresh(This,lCookie,vtPromptingInfo)	\
    (This)->lpVtbl -> Refresh(This,lCookie,vtPromptingInfo)

#define ICrystalReportSourceEx_Cancel(This,lCookie)	\
    (This)->lpVtbl -> Cancel(This,lCookie)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_GetPage_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ long lPageNumber,
    /* [in] */ VARIANT vtReserved);


void __RPC_STUB ICrystalReportSourceEx_GetPage_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_GetTotaller_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ long lStartFrom,
    /* [in] */ short nLevelsPastRoot,
    /* [in] */ VARIANT vtMaxNodeCount);


void __RPC_STUB ICrystalReportSourceEx_GetTotaller_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_GetLastPageNumber_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ VARIANT vtReserved);


void __RPC_STUB ICrystalReportSourceEx_GetLastPageNumber_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_FindGroup_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ BSTR bstrGroupPath,
    /* [in] */ VARIANT vtReserved);


void __RPC_STUB ICrystalReportSourceEx_FindGroup_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_FindText_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ long lFromPage,
    /* [in] */ long lFromInstance,
    /* [in] */ BSTR bstrText,
    /* [in] */ CRSearchDirection nMode,
    /* [in] */ VARIANT vtReserved);


void __RPC_STUB ICrystalReportSourceEx_FindText_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_DrillGraph_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ long lPageNumber,
    /* [in] */ long xOffset,
    /* [in] */ long yOffset,
    /* [in] */ BSTR bstrReserved,
    /* [in] */ VARIANT vtReserved,
    /* [in] */ VARIANT vtReserved2);


void __RPC_STUB ICrystalReportSourceEx_DrillGraph_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_DrillMap_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ long lPageNumber,
    /* [in] */ long xOffset,
    /* [in] */ long yOffset,
    /* [in] */ BSTR bstrReserved,
    /* [in] */ VARIANT vtReserved,
    /* [in] */ VARIANT vtReserved2);


void __RPC_STUB ICrystalReportSourceEx_DrillMap_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_Search_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ long lPageN,
    /* [in] */ long lSectionInstN,
    /* [in] */ BSTR bstrSearchFormula,
    /* [in] */ BSTR bstrReserved,
    /* [in] */ VARIANT vtReserved);


void __RPC_STUB ICrystalReportSourceEx_Search_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_Export_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ BSTR bstrViewContext,
    /* [in] */ BSTR bstrSubreportContext,
    /* [in] */ BSTR bstrFormula,
    /* [in] */ VARIANT vtPromptingInfo,
    /* [in] */ VARIANT exportFormat,
    /* [in] */ VARIANT vtReserved);


void __RPC_STUB ICrystalReportSourceEx_Export_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_GetExportFormats_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie);


void __RPC_STUB ICrystalReportSourceEx_GetExportFormats_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_Refresh_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie,
    /* [in] */ VARIANT vtPromptingInfo);


void __RPC_STUB ICrystalReportSourceEx_Refresh_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEx_Cancel_Proxy( 
    ICrystalReportSourceEx __RPC_FAR * This,
    /* [in] */ long lCookie);


void __RPC_STUB ICrystalReportSourceEx_Cancel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportSourceEx_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportProperties_INTERFACE_DEFINED__
#define __ICrystalReportProperties_INTERFACE_DEFINED__

/* interface ICrystalReportProperties */
/* [object][nonextensible][hidden][oleautomation][uuid] */ 


EXTERN_C const IID IID_ICrystalReportProperties;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("3DCC8FB5-C434-11d1-A817-00A0C92784CD")
    ICrystalReportProperties : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE GetTitle( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportPropertiesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportProperties __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportProperties __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportProperties __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTitle )( 
            ICrystalReportProperties __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        END_INTERFACE
    } ICrystalReportPropertiesVtbl;

    interface ICrystalReportProperties
    {
        CONST_VTBL struct ICrystalReportPropertiesVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportProperties_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportProperties_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportProperties_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportProperties_GetTitle(This,pVal)	\
    (This)->lpVtbl -> GetTitle(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICrystalReportProperties_GetTitle_Proxy( 
    ICrystalReportProperties __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportProperties_GetTitle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportProperties_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportSourceProperties_INTERFACE_DEFINED__
#define __ICrystalReportSourceProperties_INTERFACE_DEFINED__

/* interface ICrystalReportSourceProperties */
/* [object][nonextensible][hidden][oleautomation][uuid] */ 


EXTERN_C const IID IID_ICrystalReportSourceProperties;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("6876D971-F0F2-11d1-BEDF-00A0C95A6A5C")
    ICrystalReportSourceProperties : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE GetImageType( 
            /* [retval][out] */ CRImageType __RPC_FAR *pVal) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SupportsSelectionFormula( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportSourcePropertiesVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportSourceProperties __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportSourceProperties __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportSourceProperties __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetImageType )( 
            ICrystalReportSourceProperties __RPC_FAR * This,
            /* [retval][out] */ CRImageType __RPC_FAR *pVal);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SupportsSelectionFormula )( 
            ICrystalReportSourceProperties __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        END_INTERFACE
    } ICrystalReportSourcePropertiesVtbl;

    interface ICrystalReportSourceProperties
    {
        CONST_VTBL struct ICrystalReportSourcePropertiesVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportSourceProperties_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportSourceProperties_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportSourceProperties_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportSourceProperties_GetImageType(This,pVal)	\
    (This)->lpVtbl -> GetImageType(This,pVal)

#define ICrystalReportSourceProperties_SupportsSelectionFormula(This,pVal)	\
    (This)->lpVtbl -> SupportsSelectionFormula(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICrystalReportSourceProperties_GetImageType_Proxy( 
    ICrystalReportSourceProperties __RPC_FAR * This,
    /* [retval][out] */ CRImageType __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportSourceProperties_GetImageType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceProperties_SupportsSelectionFormula_Proxy( 
    ICrystalReportSourceProperties __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB ICrystalReportSourceProperties_SupportsSelectionFormula_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportSourceProperties_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportPrinterPort_INTERFACE_DEFINED__
#define __ICrystalReportPrinterPort_INTERFACE_DEFINED__

/* interface ICrystalReportPrinterPort */
/* [object][nonextensible][hidden][oleautomation][uuid] */ 


EXTERN_C const IID IID_ICrystalReportPrinterPort;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("b4742013-45a6-11d1-abec-00a0c9274b91")
    ICrystalReportPrinterPort : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE Print( 
            /* [in] */ BSTR pBstrBranch,
            /* [in] */ short drillDownLevel) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetPrinterName( 
            /* [retval][out] */ BSTR __RPC_FAR *ppPrinterName) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetDriverName( 
            /* [retval][out] */ BSTR __RPC_FAR *ppDriverName) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetPortName( 
            /* [retval][out] */ BSTR __RPC_FAR *ppPortName) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetPaperOrientation( 
            /* [retval][out] */ CRPaperOrientation __RPC_FAR *pPaperOrientation) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetPaperSize( 
            /* [retval][out] */ CRPaperSize __RPC_FAR *pPaperSize) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportPrinterPortVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportPrinterPort __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportPrinterPort __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportPrinterPort __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Print )( 
            ICrystalReportPrinterPort __RPC_FAR * This,
            /* [in] */ BSTR pBstrBranch,
            /* [in] */ short drillDownLevel);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPrinterName )( 
            ICrystalReportPrinterPort __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *ppPrinterName);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetDriverName )( 
            ICrystalReportPrinterPort __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *ppDriverName);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPortName )( 
            ICrystalReportPrinterPort __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *ppPortName);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPaperOrientation )( 
            ICrystalReportPrinterPort __RPC_FAR * This,
            /* [retval][out] */ CRPaperOrientation __RPC_FAR *pPaperOrientation);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPaperSize )( 
            ICrystalReportPrinterPort __RPC_FAR * This,
            /* [retval][out] */ CRPaperSize __RPC_FAR *pPaperSize);
        
        END_INTERFACE
    } ICrystalReportPrinterPortVtbl;

    interface ICrystalReportPrinterPort
    {
        CONST_VTBL struct ICrystalReportPrinterPortVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportPrinterPort_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportPrinterPort_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportPrinterPort_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportPrinterPort_Print(This,pBstrBranch,drillDownLevel)	\
    (This)->lpVtbl -> Print(This,pBstrBranch,drillDownLevel)

#define ICrystalReportPrinterPort_GetPrinterName(This,ppPrinterName)	\
    (This)->lpVtbl -> GetPrinterName(This,ppPrinterName)

#define ICrystalReportPrinterPort_GetDriverName(This,ppDriverName)	\
    (This)->lpVtbl -> GetDriverName(This,ppDriverName)

#define ICrystalReportPrinterPort_GetPortName(This,ppPortName)	\
    (This)->lpVtbl -> GetPortName(This,ppPortName)

#define ICrystalReportPrinterPort_GetPaperOrientation(This,pPaperOrientation)	\
    (This)->lpVtbl -> GetPaperOrientation(This,pPaperOrientation)

#define ICrystalReportPrinterPort_GetPaperSize(This,pPaperSize)	\
    (This)->lpVtbl -> GetPaperSize(This,pPaperSize)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICrystalReportPrinterPort_Print_Proxy( 
    ICrystalReportPrinterPort __RPC_FAR * This,
    /* [in] */ BSTR pBstrBranch,
    /* [in] */ short drillDownLevel);


void __RPC_STUB ICrystalReportPrinterPort_Print_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportPrinterPort_GetPrinterName_Proxy( 
    ICrystalReportPrinterPort __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *ppPrinterName);


void __RPC_STUB ICrystalReportPrinterPort_GetPrinterName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportPrinterPort_GetDriverName_Proxy( 
    ICrystalReportPrinterPort __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *ppDriverName);


void __RPC_STUB ICrystalReportPrinterPort_GetDriverName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportPrinterPort_GetPortName_Proxy( 
    ICrystalReportPrinterPort __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *ppPortName);


void __RPC_STUB ICrystalReportPrinterPort_GetPortName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportPrinterPort_GetPaperOrientation_Proxy( 
    ICrystalReportPrinterPort __RPC_FAR * This,
    /* [retval][out] */ CRPaperOrientation __RPC_FAR *pPaperOrientation);


void __RPC_STUB ICrystalReportPrinterPort_GetPaperOrientation_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportPrinterPort_GetPaperSize_Proxy( 
    ICrystalReportPrinterPort __RPC_FAR * This,
    /* [retval][out] */ CRPaperSize __RPC_FAR *pPaperSize);


void __RPC_STUB ICrystalReportPrinterPort_GetPaperSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportPrinterPort_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportExport_INTERFACE_DEFINED__
#define __ICrystalReportExport_INTERFACE_DEFINED__

/* interface ICrystalReportExport */
/* [object][nonextensible][hidden][oleautomation][uuid] */ 


EXTERN_C const IID IID_ICrystalReportExport;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("BD10A9C0-07CC-11D2-BEFF-00A0C95A6A5C")
    ICrystalReportExport : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE SetReportSource( 
            /* [in] */ IUnknown __RPC_FAR *pNewVal) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Export( 
            /* [in] */ OLE_HANDLE hWnd,
            /* [in] */ BSTR pBstrBranch,
            /* [in] */ BSTR pBstrFormula) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetFileName( 
            /* [retval][out] */ BSTR __RPC_FAR *ppVal) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetFileName( 
            /* [in] */ BSTR pNewVal) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetExportType( 
            /* [retval][out] */ BSTR __RPC_FAR *ppVal) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetExportType( 
            /* [in] */ BSTR pNewVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportExportVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportExport __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportExport __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportExport __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetReportSource )( 
            ICrystalReportExport __RPC_FAR * This,
            /* [in] */ IUnknown __RPC_FAR *pNewVal);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Export )( 
            ICrystalReportExport __RPC_FAR * This,
            /* [in] */ OLE_HANDLE hWnd,
            /* [in] */ BSTR pBstrBranch,
            /* [in] */ BSTR pBstrFormula);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFileName )( 
            ICrystalReportExport __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *ppVal);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetFileName )( 
            ICrystalReportExport __RPC_FAR * This,
            /* [in] */ BSTR pNewVal);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetExportType )( 
            ICrystalReportExport __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *ppVal);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetExportType )( 
            ICrystalReportExport __RPC_FAR * This,
            /* [in] */ BSTR pNewVal);
        
        END_INTERFACE
    } ICrystalReportExportVtbl;

    interface ICrystalReportExport
    {
        CONST_VTBL struct ICrystalReportExportVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportExport_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportExport_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportExport_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportExport_SetReportSource(This,pNewVal)	\
    (This)->lpVtbl -> SetReportSource(This,pNewVal)

#define ICrystalReportExport_Export(This,hWnd,pBstrBranch,pBstrFormula)	\
    (This)->lpVtbl -> Export(This,hWnd,pBstrBranch,pBstrFormula)

#define ICrystalReportExport_GetFileName(This,ppVal)	\
    (This)->lpVtbl -> GetFileName(This,ppVal)

#define ICrystalReportExport_SetFileName(This,pNewVal)	\
    (This)->lpVtbl -> SetFileName(This,pNewVal)

#define ICrystalReportExport_GetExportType(This,ppVal)	\
    (This)->lpVtbl -> GetExportType(This,ppVal)

#define ICrystalReportExport_SetExportType(This,pNewVal)	\
    (This)->lpVtbl -> SetExportType(This,pNewVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICrystalReportExport_SetReportSource_Proxy( 
    ICrystalReportExport __RPC_FAR * This,
    /* [in] */ IUnknown __RPC_FAR *pNewVal);


void __RPC_STUB ICrystalReportExport_SetReportSource_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportExport_Export_Proxy( 
    ICrystalReportExport __RPC_FAR * This,
    /* [in] */ OLE_HANDLE hWnd,
    /* [in] */ BSTR pBstrBranch,
    /* [in] */ BSTR pBstrFormula);


void __RPC_STUB ICrystalReportExport_Export_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportExport_GetFileName_Proxy( 
    ICrystalReportExport __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *ppVal);


void __RPC_STUB ICrystalReportExport_GetFileName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportExport_SetFileName_Proxy( 
    ICrystalReportExport __RPC_FAR * This,
    /* [in] */ BSTR pNewVal);


void __RPC_STUB ICrystalReportExport_SetFileName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportExport_GetExportType_Proxy( 
    ICrystalReportExport __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *ppVal);


void __RPC_STUB ICrystalReportExport_GetExportType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportExport_SetExportType_Proxy( 
    ICrystalReportExport __RPC_FAR * This,
    /* [in] */ BSTR pNewVal);


void __RPC_STUB ICrystalReportExport_SetExportType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportExport_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportSourceEvents_INTERFACE_DEFINED__
#define __ICrystalReportSourceEvents_INTERFACE_DEFINED__

/* interface ICrystalReportSourceEvents */
/* [object][nonextensible][hidden][oleautomation][uuid] */ 


EXTERN_C const IID IID_ICrystalReportSourceEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("3DCC8FB3-C434-11d1-A817-00A0C92784CD")
    ICrystalReportSourceEvents : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE OnStartBinding( 
            /* [in] */ long dwReserved) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OnStopBinding( 
            /* [in] */ long hrStatus,
            /* [in] */ BSTR bstrStatusText) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OnDataAvailable( 
            /* [in] */ VARIANT vtDataDescription,
            /* [in] */ VARIANT vtData,
            /* [in] */ VARIANT vtParam) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OnProgress( 
            /* [in] */ long ulProgress,
            /* [in] */ long ulProgressMax,
            /* [in] */ long ulStatusCode,
            /* [in] */ BSTR szStatusText) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportSourceEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportSourceEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportSourceEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportSourceEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OnStartBinding )( 
            ICrystalReportSourceEvents __RPC_FAR * This,
            /* [in] */ long dwReserved);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OnStopBinding )( 
            ICrystalReportSourceEvents __RPC_FAR * This,
            /* [in] */ long hrStatus,
            /* [in] */ BSTR bstrStatusText);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OnDataAvailable )( 
            ICrystalReportSourceEvents __RPC_FAR * This,
            /* [in] */ VARIANT vtDataDescription,
            /* [in] */ VARIANT vtData,
            /* [in] */ VARIANT vtParam);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OnProgress )( 
            ICrystalReportSourceEvents __RPC_FAR * This,
            /* [in] */ long ulProgress,
            /* [in] */ long ulProgressMax,
            /* [in] */ long ulStatusCode,
            /* [in] */ BSTR szStatusText);
        
        END_INTERFACE
    } ICrystalReportSourceEventsVtbl;

    interface ICrystalReportSourceEvents
    {
        CONST_VTBL struct ICrystalReportSourceEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportSourceEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportSourceEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportSourceEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportSourceEvents_OnStartBinding(This,dwReserved)	\
    (This)->lpVtbl -> OnStartBinding(This,dwReserved)

#define ICrystalReportSourceEvents_OnStopBinding(This,hrStatus,bstrStatusText)	\
    (This)->lpVtbl -> OnStopBinding(This,hrStatus,bstrStatusText)

#define ICrystalReportSourceEvents_OnDataAvailable(This,vtDataDescription,vtData,vtParam)	\
    (This)->lpVtbl -> OnDataAvailable(This,vtDataDescription,vtData,vtParam)

#define ICrystalReportSourceEvents_OnProgress(This,ulProgress,ulProgressMax,ulStatusCode,szStatusText)	\
    (This)->lpVtbl -> OnProgress(This,ulProgress,ulProgressMax,ulStatusCode,szStatusText)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICrystalReportSourceEvents_OnStartBinding_Proxy( 
    ICrystalReportSourceEvents __RPC_FAR * This,
    /* [in] */ long dwReserved);


void __RPC_STUB ICrystalReportSourceEvents_OnStartBinding_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEvents_OnStopBinding_Proxy( 
    ICrystalReportSourceEvents __RPC_FAR * This,
    /* [in] */ long hrStatus,
    /* [in] */ BSTR bstrStatusText);


void __RPC_STUB ICrystalReportSourceEvents_OnStopBinding_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEvents_OnDataAvailable_Proxy( 
    ICrystalReportSourceEvents __RPC_FAR * This,
    /* [in] */ VARIANT vtDataDescription,
    /* [in] */ VARIANT vtData,
    /* [in] */ VARIANT vtParam);


void __RPC_STUB ICrystalReportSourceEvents_OnDataAvailable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportSourceEvents_OnProgress_Proxy( 
    ICrystalReportSourceEvents __RPC_FAR * This,
    /* [in] */ long ulProgress,
    /* [in] */ long ulProgressMax,
    /* [in] */ long ulStatusCode,
    /* [in] */ BSTR szStatusText);


void __RPC_STUB ICrystalReportSourceEvents_OnProgress_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportSourceEvents_INTERFACE_DEFINED__ */


#ifndef __ICrystalReportExportEvents_INTERFACE_DEFINED__
#define __ICrystalReportExportEvents_INTERFACE_DEFINED__

/* interface ICrystalReportExportEvents */
/* [object][nonextensible][hidden][oleautomation][uuid] */ 


EXTERN_C const IID IID_ICrystalReportExportEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("4D773761-0AD4-11d2-BF01-00A0C95A6A5C")
    ICrystalReportExportEvents : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE ExportCancelled( void) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ExportFailed( void) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ExportComplete( 
            /* [in] */ BSTR pFileName,
            /* [in] */ BSTR pFileType) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICrystalReportExportEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICrystalReportExportEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICrystalReportExportEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICrystalReportExportEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ExportCancelled )( 
            ICrystalReportExportEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ExportFailed )( 
            ICrystalReportExportEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ExportComplete )( 
            ICrystalReportExportEvents __RPC_FAR * This,
            /* [in] */ BSTR pFileName,
            /* [in] */ BSTR pFileType);
        
        END_INTERFACE
    } ICrystalReportExportEventsVtbl;

    interface ICrystalReportExportEvents
    {
        CONST_VTBL struct ICrystalReportExportEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICrystalReportExportEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICrystalReportExportEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICrystalReportExportEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICrystalReportExportEvents_ExportCancelled(This)	\
    (This)->lpVtbl -> ExportCancelled(This)

#define ICrystalReportExportEvents_ExportFailed(This)	\
    (This)->lpVtbl -> ExportFailed(This)

#define ICrystalReportExportEvents_ExportComplete(This,pFileName,pFileType)	\
    (This)->lpVtbl -> ExportComplete(This,pFileName,pFileType)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE ICrystalReportExportEvents_ExportCancelled_Proxy( 
    ICrystalReportExportEvents __RPC_FAR * This);


void __RPC_STUB ICrystalReportExportEvents_ExportCancelled_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportExportEvents_ExportFailed_Proxy( 
    ICrystalReportExportEvents __RPC_FAR * This);


void __RPC_STUB ICrystalReportExportEvents_ExportFailed_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE ICrystalReportExportEvents_ExportComplete_Proxy( 
    ICrystalReportExportEvents __RPC_FAR * This,
    /* [in] */ BSTR pFileName,
    /* [in] */ BSTR pFileType);


void __RPC_STUB ICrystalReportExportEvents_ExportComplete_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICrystalReportExportEvents_INTERFACE_DEFINED__ */


#ifndef __IOlapGridObject_INTERFACE_DEFINED__
#define __IOlapGridObject_INTERFACE_DEFINED__

/* interface IOlapGridObject */
/* [object][nonextensible][hidden][dual][oleautomation][uuid] */ 


EXTERN_C const IID IID_IOlapGridObject;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0bac5dd2-44c9-11d1-abec-00a0c9274b91")
    IOlapGridObject : public IDispatch
    {
    public:
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR __RPC_FAR *ppName) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR pName) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Kind( 
            /* [retval][out] */ CRObjectKind __RPC_FAR *pObjectKind) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Left( 
            /* [retval][out] */ long __RPC_FAR *pLeft) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_Left( 
            /* [in] */ long left) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Top( 
            /* [retval][out] */ long __RPC_FAR *pTop) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_Top( 
            /* [in] */ long top) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Width( 
            /* [retval][out] */ long __RPC_FAR *pWidth) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Height( 
            /* [retval][out] */ long __RPC_FAR *pHeight) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_LeftLineStyle( 
            /* [retval][out] */ CRLineStyle __RPC_FAR *pLeftLineStyle) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_LeftLineStyle( 
            /* [in] */ CRLineStyle leftLineStyle) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_RightLineStyle( 
            /* [retval][out] */ CRLineStyle __RPC_FAR *pRightLineStyle) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_RightLineStyle( 
            /* [in] */ CRLineStyle rightLineStyle) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_TopLineStyle( 
            /* [retval][out] */ CRLineStyle __RPC_FAR *pTopLineStyle) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_TopLineStyle( 
            /* [in] */ CRLineStyle topLineStyle) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_BottomLineStyle( 
            /* [retval][out] */ CRLineStyle __RPC_FAR *pBottomLineStyle) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_BottomLineStyle( 
            /* [in] */ CRLineStyle bottomLineStyle) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_HasDropShadow( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_HasDropShadow( 
            /* [in] */ VARIANT_BOOL bl) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_BackColor( 
            /* [retval][out] */ OLE_COLOR __RPC_FAR *pBackColor) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_BackColor( 
            /* [in] */ OLE_COLOR backColor) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_BorderColor( 
            /* [retval][out] */ OLE_COLOR __RPC_FAR *pBorderColor) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_BorderColor( 
            /* [in] */ OLE_COLOR borderColor) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ ISection __RPC_FAR *__RPC_FAR *ppParent) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_Suppress( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_Suppress( 
            /* [in] */ VARIANT_BOOL bl) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_CloseAtPageBreak( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_CloseAtPageBreak( 
            /* [in] */ VARIANT_BOOL bl) = 0;
        
        virtual /* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE get_KeepTogether( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool) = 0;
        
        virtual /* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE put_KeepTogether( 
            /* [in] */ VARIANT_BOOL bl) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOlapGridObjectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOlapGridObject __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOlapGridObject __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOlapGridObject __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Name )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *ppName);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Name )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ BSTR pName);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Kind )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ CRObjectKind __RPC_FAR *pObjectKind);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Left )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pLeft);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Left )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ long left);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Top )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pTop);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Top )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ long top);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Width )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pWidth);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Height )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pHeight);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LeftLineStyle )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ CRLineStyle __RPC_FAR *pLeftLineStyle);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_LeftLineStyle )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ CRLineStyle leftLineStyle);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RightLineStyle )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ CRLineStyle __RPC_FAR *pRightLineStyle);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RightLineStyle )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ CRLineStyle rightLineStyle);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TopLineStyle )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ CRLineStyle __RPC_FAR *pTopLineStyle);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TopLineStyle )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ CRLineStyle topLineStyle);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_BottomLineStyle )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ CRLineStyle __RPC_FAR *pBottomLineStyle);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_BottomLineStyle )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ CRLineStyle bottomLineStyle);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_HasDropShadow )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_HasDropShadow )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL bl);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_BackColor )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ OLE_COLOR __RPC_FAR *pBackColor);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_BackColor )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ OLE_COLOR backColor);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_BorderColor )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ OLE_COLOR __RPC_FAR *pBorderColor);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_BorderColor )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ OLE_COLOR borderColor);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Parent )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ ISection __RPC_FAR *__RPC_FAR *ppParent);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Suppress )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Suppress )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL bl);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CloseAtPageBreak )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CloseAtPageBreak )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL bl);
        
        /* [helpstring][propget][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_KeepTogether )( 
            IOlapGridObject __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool);
        
        /* [helpstring][propput][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_KeepTogether )( 
            IOlapGridObject __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL bl);
        
        END_INTERFACE
    } IOlapGridObjectVtbl;

    interface IOlapGridObject
    {
        CONST_VTBL struct IOlapGridObjectVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOlapGridObject_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOlapGridObject_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOlapGridObject_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOlapGridObject_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOlapGridObject_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOlapGridObject_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOlapGridObject_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOlapGridObject_get_Name(This,ppName)	\
    (This)->lpVtbl -> get_Name(This,ppName)

#define IOlapGridObject_put_Name(This,pName)	\
    (This)->lpVtbl -> put_Name(This,pName)

#define IOlapGridObject_get_Kind(This,pObjectKind)	\
    (This)->lpVtbl -> get_Kind(This,pObjectKind)

#define IOlapGridObject_get_Left(This,pLeft)	\
    (This)->lpVtbl -> get_Left(This,pLeft)

#define IOlapGridObject_put_Left(This,left)	\
    (This)->lpVtbl -> put_Left(This,left)

#define IOlapGridObject_get_Top(This,pTop)	\
    (This)->lpVtbl -> get_Top(This,pTop)

#define IOlapGridObject_put_Top(This,top)	\
    (This)->lpVtbl -> put_Top(This,top)

#define IOlapGridObject_get_Width(This,pWidth)	\
    (This)->lpVtbl -> get_Width(This,pWidth)

#define IOlapGridObject_get_Height(This,pHeight)	\
    (This)->lpVtbl -> get_Height(This,pHeight)

#define IOlapGridObject_get_LeftLineStyle(This,pLeftLineStyle)	\
    (This)->lpVtbl -> get_LeftLineStyle(This,pLeftLineStyle)

#define IOlapGridObject_put_LeftLineStyle(This,leftLineStyle)	\
    (This)->lpVtbl -> put_LeftLineStyle(This,leftLineStyle)

#define IOlapGridObject_get_RightLineStyle(This,pRightLineStyle)	\
    (This)->lpVtbl -> get_RightLineStyle(This,pRightLineStyle)

#define IOlapGridObject_put_RightLineStyle(This,rightLineStyle)	\
    (This)->lpVtbl -> put_RightLineStyle(This,rightLineStyle)

#define IOlapGridObject_get_TopLineStyle(This,pTopLineStyle)	\
    (This)->lpVtbl -> get_TopLineStyle(This,pTopLineStyle)

#define IOlapGridObject_put_TopLineStyle(This,topLineStyle)	\
    (This)->lpVtbl -> put_TopLineStyle(This,topLineStyle)

#define IOlapGridObject_get_BottomLineStyle(This,pBottomLineStyle)	\
    (This)->lpVtbl -> get_BottomLineStyle(This,pBottomLineStyle)

#define IOlapGridObject_put_BottomLineStyle(This,bottomLineStyle)	\
    (This)->lpVtbl -> put_BottomLineStyle(This,bottomLineStyle)

#define IOlapGridObject_get_HasDropShadow(This,pBool)	\
    (This)->lpVtbl -> get_HasDropShadow(This,pBool)

#define IOlapGridObject_put_HasDropShadow(This,bl)	\
    (This)->lpVtbl -> put_HasDropShadow(This,bl)

#define IOlapGridObject_get_BackColor(This,pBackColor)	\
    (This)->lpVtbl -> get_BackColor(This,pBackColor)

#define IOlapGridObject_put_BackColor(This,backColor)	\
    (This)->lpVtbl -> put_BackColor(This,backColor)

#define IOlapGridObject_get_BorderColor(This,pBorderColor)	\
    (This)->lpVtbl -> get_BorderColor(This,pBorderColor)

#define IOlapGridObject_put_BorderColor(This,borderColor)	\
    (This)->lpVtbl -> put_BorderColor(This,borderColor)

#define IOlapGridObject_get_Parent(This,ppParent)	\
    (This)->lpVtbl -> get_Parent(This,ppParent)

#define IOlapGridObject_get_Suppress(This,pBool)	\
    (This)->lpVtbl -> get_Suppress(This,pBool)

#define IOlapGridObject_put_Suppress(This,bl)	\
    (This)->lpVtbl -> put_Suppress(This,bl)

#define IOlapGridObject_get_CloseAtPageBreak(This,pBool)	\
    (This)->lpVtbl -> get_CloseAtPageBreak(This,pBool)

#define IOlapGridObject_put_CloseAtPageBreak(This,bl)	\
    (This)->lpVtbl -> put_CloseAtPageBreak(This,bl)

#define IOlapGridObject_get_KeepTogether(This,pBool)	\
    (This)->lpVtbl -> get_KeepTogether(This,pBool)

#define IOlapGridObject_put_KeepTogether(This,bl)	\
    (This)->lpVtbl -> put_KeepTogether(This,bl)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_Name_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *ppName);


void __RPC_STUB IOlapGridObject_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_Name_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ BSTR pName);


void __RPC_STUB IOlapGridObject_put_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_Kind_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ CRObjectKind __RPC_FAR *pObjectKind);


void __RPC_STUB IOlapGridObject_get_Kind_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_Left_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pLeft);


void __RPC_STUB IOlapGridObject_get_Left_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_Left_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ long left);


void __RPC_STUB IOlapGridObject_put_Left_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_Top_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pTop);


void __RPC_STUB IOlapGridObject_get_Top_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_Top_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ long top);


void __RPC_STUB IOlapGridObject_put_Top_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_Width_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pWidth);


void __RPC_STUB IOlapGridObject_get_Width_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_Height_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pHeight);


void __RPC_STUB IOlapGridObject_get_Height_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_LeftLineStyle_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ CRLineStyle __RPC_FAR *pLeftLineStyle);


void __RPC_STUB IOlapGridObject_get_LeftLineStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_LeftLineStyle_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ CRLineStyle leftLineStyle);


void __RPC_STUB IOlapGridObject_put_LeftLineStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_RightLineStyle_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ CRLineStyle __RPC_FAR *pRightLineStyle);


void __RPC_STUB IOlapGridObject_get_RightLineStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_RightLineStyle_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ CRLineStyle rightLineStyle);


void __RPC_STUB IOlapGridObject_put_RightLineStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_TopLineStyle_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ CRLineStyle __RPC_FAR *pTopLineStyle);


void __RPC_STUB IOlapGridObject_get_TopLineStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_TopLineStyle_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ CRLineStyle topLineStyle);


void __RPC_STUB IOlapGridObject_put_TopLineStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_BottomLineStyle_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ CRLineStyle __RPC_FAR *pBottomLineStyle);


void __RPC_STUB IOlapGridObject_get_BottomLineStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_BottomLineStyle_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ CRLineStyle bottomLineStyle);


void __RPC_STUB IOlapGridObject_put_BottomLineStyle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_HasDropShadow_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool);


void __RPC_STUB IOlapGridObject_get_HasDropShadow_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_HasDropShadow_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL bl);


void __RPC_STUB IOlapGridObject_put_HasDropShadow_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_BackColor_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ OLE_COLOR __RPC_FAR *pBackColor);


void __RPC_STUB IOlapGridObject_get_BackColor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_BackColor_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ OLE_COLOR backColor);


void __RPC_STUB IOlapGridObject_put_BackColor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_BorderColor_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ OLE_COLOR __RPC_FAR *pBorderColor);


void __RPC_STUB IOlapGridObject_get_BorderColor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_BorderColor_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ OLE_COLOR borderColor);


void __RPC_STUB IOlapGridObject_put_BorderColor_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_Parent_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ ISection __RPC_FAR *__RPC_FAR *ppParent);


void __RPC_STUB IOlapGridObject_get_Parent_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_Suppress_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool);


void __RPC_STUB IOlapGridObject_get_Suppress_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_Suppress_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL bl);


void __RPC_STUB IOlapGridObject_put_Suppress_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_CloseAtPageBreak_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool);


void __RPC_STUB IOlapGridObject_get_CloseAtPageBreak_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_CloseAtPageBreak_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL bl);


void __RPC_STUB IOlapGridObject_put_CloseAtPageBreak_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propget][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_get_KeepTogether_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pBool);


void __RPC_STUB IOlapGridObject_get_KeepTogether_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][propput][id] */ HRESULT STDMETHODCALLTYPE IOlapGridObject_put_KeepTogether_Proxy( 
    IOlapGridObject __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL bl);


void __RPC_STUB IOlapGridObject_put_KeepTogether_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOlapGridObject_INTERFACE_DEFINED__ */



#ifndef __CRAXDRT_LIBRARY_DEFINED__
#define __CRAXDRT_LIBRARY_DEFINED__

/* library CRAXDRT */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_CRAXDRT;

EXTERN_C const CLSID CLSID_Report;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741e10-45a6-11d1-abec-00a0c9274b91")
Report;
#endif

EXTERN_C const CLSID CLSID_Areas;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741e20-45a6-11d1-abec-00a0c9274b91")
Areas;
#endif

EXTERN_C const CLSID CLSID_Sections;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741e30-45a6-11d1-abec-00a0c9274b91")
Sections;
#endif

EXTERN_C const CLSID CLSID_Area;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741e40-45a6-11d1-abec-00a0c9274b91")
Area;
#endif

EXTERN_C const CLSID CLSID_Section;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741e50-45a6-11d1-abec-00a0c9274b91")
Section;
#endif

EXTERN_C const CLSID CLSID_ReportObjects;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741e60-45a6-11d1-abec-00a0c9274b91")
ReportObjects;
#endif

EXTERN_C const CLSID CLSID_FieldObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741e70-45a6-11d1-abec-00a0c9274b91")
FieldObject;
#endif

EXTERN_C const CLSID CLSID_TextObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741e90-45a6-11d1-abec-00a0c9274b91")
TextObject;
#endif

EXTERN_C const CLSID CLSID_SubreportObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741ea0-45a6-11d1-abec-00a0c9274b91")
SubreportObject;
#endif

EXTERN_C const CLSID CLSID_DatabaseFieldDefinition;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741ec0-45a6-11d1-abec-00a0c9274b91")
DatabaseFieldDefinition;
#endif

EXTERN_C const CLSID CLSID_FormulaFieldDefinition;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5b90-44c9-11d1-abec-00a0c9274b91")
FormulaFieldDefinition;
#endif

EXTERN_C const CLSID CLSID_ParameterFieldDefinition;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5bb0-44c9-11d1-abec-00a0c9274b91")
ParameterFieldDefinition;
#endif

EXTERN_C const CLSID CLSID_GroupNameFieldDefinition;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5bd0-44c9-11d1-abec-00a0c9274b91")
GroupNameFieldDefinition;
#endif

EXTERN_C const CLSID CLSID_SpecialVarFieldDefinition;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5be0-44c9-11d1-abec-00a0c9274b91")
SpecialVarFieldDefinition;
#endif

EXTERN_C const CLSID CLSID_SummaryFieldDefinition;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5bf0-44c9-11d1-abec-00a0c9274b91")
SummaryFieldDefinition;
#endif

EXTERN_C const CLSID CLSID_RunningTotalFieldDefinition;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742070-45a6-11d1-abec-00a0c9274b91")
RunningTotalFieldDefinition;
#endif

EXTERN_C const CLSID CLSID_SQLExpressionFieldDefinition;

#ifdef __cplusplus

class DECLSPEC_UUID("b47420a0-45a6-11d1-abec-00a0c9274b91")
SQLExpressionFieldDefinition;
#endif

EXTERN_C const CLSID CLSID_Database;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741ee0-45a6-11d1-abec-00a0c9274b91")
Database;
#endif

EXTERN_C const CLSID CLSID_DatabaseTables;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741ef0-45a6-11d1-abec-00a0c9274b91")
DatabaseTables;
#endif

EXTERN_C const CLSID CLSID_DatabaseTable;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f00-45a6-11d1-abec-00a0c9274b91")
DatabaseTable;
#endif

EXTERN_C const CLSID CLSID_DatabaseFieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f10-45a6-11d1-abec-00a0c9274b91")
DatabaseFieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_FormulaFieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5b80-44c9-11d1-abec-00a0c9274b91")
FormulaFieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_ParameterFieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5ba0-44c9-11d1-abec-00a0c9274b91")
ParameterFieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_GroupNameFieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5bc0-44c9-11d1-abec-00a0c9274b91")
GroupNameFieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_SummaryFieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5c00-44c9-11d1-abec-00a0c9274b91")
SummaryFieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_RunningTotalFieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742080-45a6-11d1-abec-00a0c9274b91")
RunningTotalFieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_SQLExpressionFieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742090-45a6-11d1-abec-00a0c9274b91")
SQLExpressionFieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_GraphObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f20-45a6-11d1-abec-00a0c9274b91")
GraphObject;
#endif

EXTERN_C const CLSID CLSID_MapObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742060-45a6-11d1-abec-00a0c9274b91")
MapObject;
#endif

EXTERN_C const CLSID CLSID_OleObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f30-45a6-11d1-abec-00a0c9274b91")
OleObject;
#endif

EXTERN_C const CLSID CLSID_BlobFieldObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f40-45a6-11d1-abec-00a0c9274b91")
BlobFieldObject;
#endif

EXTERN_C const CLSID CLSID_LineObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f50-45a6-11d1-abec-00a0c9274b91")
LineObject;
#endif

EXTERN_C const CLSID CLSID_BoxObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f60-45a6-11d1-abec-00a0c9274b91")
BoxObject;
#endif

EXTERN_C const CLSID CLSID_OlapGridObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b47420b0-45a6-11d1-abec-00a0c9274b91")
OlapGridObject;
#endif

EXTERN_C const CLSID CLSID_CrossTabObject;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f70-45a6-11d1-abec-00a0c9274b91")
CrossTabObject;
#endif

EXTERN_C const CLSID CLSID_PageEngine;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f80-45a6-11d1-abec-00a0c9274b91")
PageEngine;
#endif

EXTERN_C const CLSID CLSID_PageGenerator;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741f90-45a6-11d1-abec-00a0c9274b91")
PageGenerator;
#endif

EXTERN_C const CLSID CLSID_Pages;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741fa0-45a6-11d1-abec-00a0c9274b91")
Pages;
#endif

EXTERN_C const CLSID CLSID_Page;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741fb0-45a6-11d1-abec-00a0c9274b91")
Page;
#endif

EXTERN_C const CLSID CLSID_ExportOptions;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741fc0-45a6-11d1-abec-00a0c9274b91")
ExportOptions;
#endif

EXTERN_C const CLSID CLSID_Application;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741fd0-45a6-11d1-abec-00a0c9274b91")
Application;
#endif

EXTERN_C const CLSID CLSID_FormattingInfo;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741fe0-45a6-11d1-abec-00a0c9274b91")
FormattingInfo;
#endif

EXTERN_C const CLSID CLSID_SortFields;

#ifdef __cplusplus

class DECLSPEC_UUID("b4741ff0-45a6-11d1-abec-00a0c9274b91")
SortFields;
#endif

EXTERN_C const CLSID CLSID_SortField;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742000-45a6-11d1-abec-00a0c9274b91")
SortField;
#endif

EXTERN_C const CLSID CLSID_PrintingStatus;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742030-45a6-11d1-abec-00a0c9274b91")
PrintingStatus;
#endif

EXTERN_C const CLSID CLSID_SubreportLink;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742130-45a6-11d1-abec-00a0c9274b91")
SubreportLink;
#endif

EXTERN_C const CLSID CLSID_SubreportLinks;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742120-45a6-11d1-abec-00a0c9274b91")
SubreportLinks;
#endif

EXTERN_C const CLSID CLSID_CrossTabGroups;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742100-45a6-11d1-abec-00a0c9274b91")
CrossTabGroups;
#endif

EXTERN_C const CLSID CLSID_CrossTabGroup;

#ifdef __cplusplus

class DECLSPEC_UUID("b4742110-45a6-11d1-abec-00a0c9274b91")
CrossTabGroup;
#endif

EXTERN_C const CLSID CLSID_FieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("b47420e0-45a6-11d1-abec-00a0c9274b91")
FieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_ObjectSummaryFieldDefinitions;

#ifdef __cplusplus

class DECLSPEC_UUID("b47420f0-45a6-11d1-abec-00a0c9274b91")
ObjectSummaryFieldDefinitions;
#endif

EXTERN_C const CLSID CLSID_TableLink;

#ifdef __cplusplus

class DECLSPEC_UUID("b47420d0-45a6-11d1-abec-00a0c9274b91")
TableLink;
#endif

EXTERN_C const CLSID CLSID_TableLinks;

#ifdef __cplusplus

class DECLSPEC_UUID("b47420c0-45a6-11d1-abec-00a0c9274b91")
TableLinks;
#endif

EXTERN_C const CLSID CLSID_FieldMappingData;

#ifdef __cplusplus

class DECLSPEC_UUID("0bac5e60-44c9-11d1-abec-00a0c9274b91")
FieldMappingData;
#endif
#endif /* __CRAXDRT_LIBRARY_DEFINED__ */

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
