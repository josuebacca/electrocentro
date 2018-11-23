Attribute VB_Name = "Module1"
'
'               Visual Basic Declarations of CRPE32.DLL
'               =====================================
'
'       File:         GLOBAL32.BAS
'
'       Author:       Seagate Software Information Management Group, Inc.
'       Date:         15 Apr 92
'
'       Purpose:      This file presents the API to the Crystal Reports
'                     Print Engine DLL (Professional).
'
'       Language:     Visual Basic for Windows
'
'       Copyright (c) 1992 - 1999 Seagate Software Information Management Group, Inc.
'
'       Revisions:
'
'          CCS  15 Apr 92  -  Original Development
'          KYL  12 Jul 92  -  Modified Existing Declarations
'                             Added Missing Declarations
'          KYL  27 Aug 92  -  Converted to CRPE32.DLL
'          CRD  08 Feb 93  -  Added new calls for 2.0 and Global declares for samples
'          CRD  25 Feb 93  -  Added new calls for 2.0 Pro
'          RBC  23 Apr 93  -  Added more new calls, rearranged to match CRPE.H
'          DVA  22 Dec 93  -  Added new calls for 3.0
'          TW   15 Mar 94  -  3.0 call reorganization
'          RS   28 Aug 95  -  32-bit update
'          JEA  12 May 96  -  Revised for 5.0
'                             Added the following Error Codes
'                               PE_ERR_BADSECTIONHEIGHT     = 560
'                               PE_ERR_BADVALUETYPE         = 561
'                               PE_ERR_INVALIDSUBREPORTNAME = 562
'                               PE_ERR_FIELDEXIST           = 563
'                               PE_ERR_NOPARENTWINDOW       = 564
'                               PE_ERR_INVALIDZOOMFACTOR    = 565
'                             Added the Following Constants
'                               PE_WORD_LEN = 2
'                               PE_PF_NAME_LEN = 256
'                               PE_PF_PROMPT_LEN = 256
'                               PE_PF_VALUE_LEN = 256
'                               PE_PF_NUMBER = 0
'                               PE_PF_CURRENCY = 1
'                               PE_PF_BOOLEAN = 2
'                               PE_PF_DATE = 3
'                               PE_PF_STRING = 4
'                               PE_SIZEOF_VARINFO_TYPE
'                               PE_SUBREPORT_NAME_LEN = 128
'                               PE_SIZEOF_SUBREPORT_INFO
'                               PE_PARAMETER_NAME_LEN = 128
'                               PE_SIZEOF_PARAMETER_INFO
'                               PE_SECT_PAGE_HEADER = 2
'                               PE_SECT_PAGE_FOOTER = 7
'                               PE_SECT_REPORT_HEADER = 1
'                               PE_SECT_REPORT_FOOTER = 8
'                               PE_SECT_GROUP_HEADER = 3
'                               PE_SECT_GROUP_FOOTER = 5
'                               PE_SECT_DETAIL = 4
'                             Added the following Structures
'                               PEParameterFieldInfo
'                               PESubreportInfo
'                               PEParameterInfo
'                             Added the following Declarations
'                               PEGetNParameterFields
'                               PEGetNthParameterField
'                               PESetNthParameterField
'                               PEGetNSubreportsInSection
'                               PEGetNthSubreportInSection
'                               PEGetSubreportInfo
'                               PEOpenSubreport
'                               PECloseSubreport
'                               PESetDialogParentWindow
'                               PEEnableProgressDialog
'                               PEGetNPages
'                               PEShowNthPage
'                               PEGetNSections
'                               PEGetSectionCode
'                               PEGetNthParamInfo
'                             Added the following Functions
'                               PE_SECTION_CODE
'                               PE_SECTION_TYPE
'                               PE_GROUP_N
'                               PE_SECTION_N
'          JEA 13 May 96  -  Changed PELogOnServerWithPrivateInfo 2nd parameter from
'                            PrivateInfo As Any to ByVal PrivateInfo As Long
'          JEA 02 Jun 96  -  Added Following Constants
'                              PE_UNCHANGED_COLOR = -2
'                            Added Following Structure Members
'                              PESectionOptions.underlaySection
'                              PESectionOptions.backgroundColor
'          JEA 27 Jun 96  -  Added the Following Constants
'                              PE_SIZEOF_EXPORT_OPTIONS
'                              PE_SIZEOF_GRAPH_DATA_INFO
'                              PE_SIZEOF_GRAPH_TEXT_INFO
'                              PE_SIZEOF_GRAPH_OPTIONS
'                              PE_SIZEOF_JOB_INFO
'                              PE_SIZEOF_PRINT_FILE_OPTIONS
'                              PE_SIZEOF_CHAR_SEP_FILE_OPTIONS
'                              PE_SIZEOF_PRINT_OPTIONS
'                              PE_SIZEOF_SECTION_OPTIONS
'          JEA 08 Jul 96  -  Added reserve members to these structures for alignment:
'                              Type PEPrintFileOptions
'                                  StructSize As Integer   ' initialize to # of bytes in PEPrintFileOptions
'
'                                  UseReportNumberFmt As Integer
'                                  reserved1 As Integer     ' reserved - do not set
'                                  UseReportDateFormat As Integer
'                                  reserved2 As Integer     ' reserved - do not set
'                              End Type
'
'                              Type PECharSepFileOptions
'                                  StructSize As Integer   ' initialize to PE_SIZEOF_CHAR_SEP_FILE_OPTIONS
'
'                                  UseReportNumberFmt As Integer
'                                  reserved1 As Integer     ' reserved - do not set
'                                  UseReportDateFormat As Integer
'                                  reserved2 As Integer     ' reserved - do not set
'
'                                  StringDelimiter As String * 1
'                                  FieldDelimiter As String * PE_FIELDDELIMLEN
'                              End Type
'         JEA 09 Jul 96  -  Changed Jobinfo Structure declaration
'                           from
'                           PEJobInfo
'                               PrintEnded As Integer
'                           PEJobInfo
'                               PrintEnded As Long
'
'         JEA 11 Jul 96  -  Removed declarations for the following calls
'                            PESetLineHeight
'                            PEGetNLinesInSection
'                            PEGetLineHeight
'                            For old code:
'                               PESetLineHeight now calls PESetMinimumSectionHeight
'                               PEGetLineHeight now calls PEGetMinimumSectionHeight
'                               PEGetNLineInSection always returns 1.
'
'                            Changed the DEVMODE argument of the follow calls to As Any
'                               PESelectPrinter
'                               PEGetSelectedPrinter
'          JEA 02 Aug 96  -  Added the following error code:
'                               Global Const PE_ERR_PAGESIZEOVERFLOW = 567
'                               Global Const PE_ERR_LOWSYSTEMRESOURCES = 568
'                         -  Added the following format formula name constants:
'                               Global Const SECTION_VISIBILITY = 58
'                               Global Const NEW_PAGE_BEFORE = 60
'                               Global Const NEW_PAGE_AFTER = 61
'                               Global Const KEEP_SECTION_TOGETHER = 62
'                               Global Const SUPPRESS_BLANK_SECTION = 63
'                               Global Const RESET_PAGE_N_AFTER = 64
'                               Global Const PRINT_AT_BOTTOM_OF_PAGE = 65
'                               Global Const UNDERLAY_SECTION = 66
'                               Global Const SECTION_BACK_COLOUR = 67
'                         -  Added the following function
'Declare Function PESetSectionFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal formulaName%, ByVal formulaString$) As Integer
'                         -  Changed structure member from
'                              PESectionOptions.SuppressBlankLines
'                              to PESectionOptions.SuppressBlankSection
'          JEA 07 Aug 96  -  Added Function PEVBGetVersion
'          JEA 09 Aug 96  -  Added Function To re-direct PEGetJobStatus
'                            Changed PEJobInfo.numRecordsRead,
'                                    PEJobInfo.numRecordsSelected,
'                                    PEJobInfo.numRecordsPrinted back to single Longs.
'          JEA 23 Aug 96  -  Added
'                            showArea As Integer
'                            freeFormPlacement As Integer
'                            members to PESectionOptions structure
'          JEA 28 Aug 96  -  Changed the Following Declarations:
'                            Changed
'                            Declare Function PEHasSavedData Lib "crpe32.dll" (ByVal printJob%, HasSavedData%) As Integer
'                            to
'                            Declare Function PEHasSavedData Lib "crpe32.dll" (ByVal printJob%, HasSavedData As Long) As Integer
'                            Changed
'                            Declare Function PEPrintControlsShowing Lib "crpe32.dll" (ByVal printJob%, ControlsShowing%) As Integer
'                            to
'                            Declare Function PEPrintControlsShowing Lib "crpe32.dll" (ByVal printJob%, ControlsShowing As Long) As Integer
'          JEA 07 OCT 96  -  Changed
'                            PEGraphOptions.ShowDataValue
'                            PEGraphOptions.ShowGridLine
'                            PEGraphOptions.VerticalBars
'                            PEGraphOptions.ShowLegend to Longs to match the CRPE32.DLL which
'                            expects 4 Byte BOOLS.
'                            Changed
'                            Global Const PE_SIZEOF_GRAPH_OPTIONS = 5 * PE_WORD_LEN + 2 * 8 + PE_GRAPH_TEXT_LEN
'                            to
'                            Global Const PE_SIZEOF_GRAPH_OPTIONS = PE_WORD_LEN + 2 * 8 + 4 * 4 + PE_GRAPH_TEXT_LEN
'                            to accomodate new longs.
'          CGH 10 OCT 96  -  Changed PEPrintReport window size parameters to Longs to accomadate CW_USEDEFAULT
'          JEA 15 Oct 96  -  Added the following declarations
'                              PESetAreaFormat
'                              PESetAreaFormatFormula
'                              PEGetAreaFormat
'                            Added function to create area codes:
'                              Function PE_AREA_CODE(sectionType%, groupN%) As Integer
'          JEA 04 Nov 96  -  Added constant
'                              PE_SIZEOF_PARAMETER_FIELD_INFO
'          JEA 09 JAN 97  -  Added new naming convention (PE_FFN) for format formula name constants.
'                         -  Added new format formula name constant PE_FFN_SHOW_AREA = 59
'                         -  Added the following constants
'                              Global Const PE_VI_STRING_LEN = 256
'                              Global Const PE_BYTE_LEN = 1
'                              Global Const PE_LONG_LEN = 4
'                              Global Const PE_DOUBLE_LEN = 8
'                              Global Const PE_VI_STRING_LEN = 256
'                              Global Const PE_SIZEOF_VALUE_INFO = 1 * PE_BYTE_LEN + _
'                                                                 15 * PE_WORD_LEN + _
'                                                                  2 * PE_LONG_LEN + _
'                                                                  2 * PE_DOUBLE_LEN + _
'                                                                  1 * PE_VI_STRING_LEN
'                         -  Added the following Type
'                              Type PEValueInfo
'                         -  Added the following declarations
'                              Declare Function PEConvertPFInfoToVInfo...
'                              Declare Function PEConvertVInfoToPFInfo...
'                         -  Added the following Error Codes
'                             Global Const PE_ERR_BADGROUPNUMBER = 570
'                             Global Const PE_ERR_INVALIDNEGATIVEVALUE = 572
'                             Global Const PE_ERR_INVALIDMEMORYPOINTER = 573
'                             Global Const PE_ERR_INVALIDPARAMETERNUMBER = 594
'                             Global Const PE_ERR_SQLSERVERNOTOPENED = 599

'          JEA 14 JAN 97  -  Changed
'                              PE_SECTION_TYPE = ((sectionCode) / 1000)
'                                to
'                              PE_SECTION_TYPE = ((sectionCode) \ 1000)
'                                to truncate instead of round.
'                            Changed
'                              PE_SECTION_N = (((sectionCode / 25) Mod 40))
'                                to
'                              PE_SECTION_N = (((sectionCode \ 25) Mod 40))
'                                to truncate instead of round.
'          JEA 31 JAN 97  - Added constant
'                              Global Const PE_PF_REPORT_NAME_LEN = 128
'                         - Added Structure Members
'                              PEParameterFieldInfo.ReportName
'                              PEParameterFieldInfo.needsCurrentValue
'          JEA 29 APR 97  - Added Constant
'                              PE_SIZEOF_WINDOW_OPTIONS
'                              PE_FFN_AREASECTION_VISIBILITY
'                           Added Type
'                              PEWindowOptions
'                           Added Functions
'                              PEGetSectionFormatFormula
'                              PEGetAreaFormatFormula
'                              PEGetWindowOptions
'                              PESetWindowOptions
'          JEA 20 MAY 97  - Added Constants
'                              PE_VI_NUMBER
'                              PE_VI_CURRENCY
'                              PE_VI_BOOLEAN
'                              PE_VI_DATE
'                              PE_VI_STRING
'                              PE_VI_DATETIME
'                              PE_VI_TIME
'                              PE_VI_INTEGER
'                              PE_VI_COLOR
'                              PE_VI_CHAR
'                              PE_VI_LONG
'                              PE_VI_NOVALUE
'                           Added Structure Members
'                              PEWindowOptions.hasPrintSetupButton
'                              PEWindowOptions.hasRefreshButton
'                              PEValueInfo.ignored
'                              PEValueInfo.viLong
'                           Updated Constants
'                              PE_SIZEOF_WINDOW_OPTIONS
'                              PE_SIZEOF_VALUE_INFO
'          JEA 14 JUL 97  - Added Constants
'                              PE_GO_TBN_ALL_GROUPS_UNSORTED
'                              PE_GO_TBN_ALL_GROUPS_SORTED
'                              PE_GO_TBN_TOP_N_GROUPS
'                              PE_GO_TBN_BOTTOM_N_GROUPS
'                              PE_SF_ORIGINAL
'                              PE_SF_SPECIFIED
'                              PE_SIZEOF_GROUP_OPTIONS
'                              PE_SI_APPLICATION_NAME_LEN
'                              PE_SI_TITLE_LEN
'                              PE_SI_SUBJECT_LEN
'                              PE_SI_AUTHOR_LEN
'                              PE_SI_KEYWORDS_LEN
'                              PE_SI_COMMENTS_LEN
'                              PE_SI_REPORT_TEMPLATE_LEN
'                              PE_SIZEOF_REPORT_SUMMARY_INFO
'                              PE_PT_LONGVARCHAR
'                              PE_PT_BINARY
'                              PE_PT_VARBINARY
'                              PE_PT_LONGVARBINARY
'                              PE_PT_BIGINT
'                              PE_PT_TINYINT
'                              PE_PT_BIT
'                              PE_PT_CHAR
'                              PE_PT_NUMERIC
'                              PE_PT_DECIMAL
'                              PE_PT_INTEGER
'                              PE_PT_SMALLINT
'                              PE_PT_FLOAT
'                              PE_PT_REAL
'                              PE_PT_DOUBLE
'                              PE_PT_DATE
'                              PE_PT_TIME
'                              PE_PT_TIMESTAMP
'                              PE_PT_VARCHAR
'                           Added Types
'                              PEGroupOptions
'                              PEReportSummaryInfo
'                           Added Functions
'                              PEGetGroupOptions
'                              PESetGroupOptions
'                              PEGetReportSummaryInfo
'                              PESetReportSummaryInfo
'          ABC 25 Aug 98  - Added Constants
'                              PE_ERR_INVALIDOBJECTFORMATNAME
'                              PE_ERR_INVALIDOBJECTTYPE
'                              PE_ERR_INVALIDGRAPHDATATYPE
'                              PE_ERR_INVALIDSUBREPORTLINKNUMBER
'                              PE_ERR_SUBREPORTLINKEXIST
'                              PE_ERR_BADROWCOLVALUE
'                              PE_ERR_INVALIDSUMMARYNUMBER
'                              PE_ERR_INVALIDGRAPHDATAFIELDNUMBER
'                              PE_ERR_INVALIDSUBREPORTNUMBER
'                              PE_ERR_INVALIDFIELDSCOPE
'                              PE_ERR_FIELDINUSE
'                              PE_ERR_INVALIDPAGEMARGINS
'                              PE_ERR_REPORTONSECUREQUERY
'                              PE_ERR_CANNOTOPENSECUREQUERY
'                              PE_ERR_INVALIDSECTIONNUMBER
'                              PE_ERR_TABLENAMEEXIST
'                              PE_ERR_INVALIDCURSOR
'                              PE_ERR_FIRSTPASSNOTFINISHED
'                              PE_ERR_CREATEDATASOURCE
'                              PE_ERR_CREATEDRILLDOWNPARAMETERS
'                              PE_ERR_CHECKFORDATASOURCECHANGES
'                              PE_ERR_STARTBACKGROUNDPROCESSING
'                              PE_ERR_SQLSERVERINUSE
'                              PE_ERR_GROUPSORTFIELDNOTSET
'                              PE_ERR_CANNOTSETSAVESUMMARIES
'                              PE_ERR_LOADOLAPDATABASEMANAGER
'                              PE_ERR_OPENOLAPCUBE
'                              PE_ERR_READOLAPCUBEDATA
'                              PE_ERR_CANNOTSAVEQUERY
'                              PE_ERR_CANNOTREADQUERYDATA
'                              PE_ERR_MAINREPORTFIELDLINKED
'                              PE_ERR_INVALIDMAPPINGTYPEVALUE
'                              PE_ERR_HITTESTFAILED
'                              PE_ERR_BADSQLEXPRESSIONNAME
'                              PE_ERR_BADSQLEXPRESSIONNUMBER
'                              PE_ERR_BADSQLEXPRESSIONTEXT
'                              PE_ERR_INVALIDDEFAULTVALUEINDEX
'                              PE_ERR_NOMINMAXVALUE
'                              PE_ERR_INCONSISTANTTYPES
'                              PE_ERR_CANNOTLINKTABLES
'                              PE_ERR_CREATEROUTER
'                              PE_ERR_INVALIDFIELDINDEX
'                              PE_ERR_INVALIDGRAPHTITLETYPE
'                              PE_ERR_INVALIDGRAPHTITLEFONTTYPE
'                              PE_ERR_OTHERERROR
'                              PE_ERR_INTERNALERROR
'                              PE_RPTOPT_CVTDATETIMETOSTR
'                              PE_RPTOPT_CVTDATETIMETODATE
'                              PE_RPTOPT_KEEPDATETIMETYPE
'                              PE_SIZEOF_REPORT_OPTIONS
'                              PE_OE_SINGLE_THREADED
'                              PE_OE_MULTI_THREADED
'                              PE_SIZEOF_ENGINE_OPTIONS
'                              PE_PF_EDITMASK_LEN
'                              PE_PF_DATETIME
'                              PE_PF_TIME
'                              PE_PO_REPORT
'                              PE_PO_STOREDPROC
'                              PE_PO_QUERY
'                              PE_SIZEOF_PARAMETER_VALUE_INFO
'                              PE_DT_SQL_STORED_PROCEDURE
'                              PE_FILE_PATH_LEN
'                              PE_FM_AUTO_FLD_MAP
'                              PE_FM_CRPE_PROMPT_FLD_MAP
'                              PE_RI_INCLUDEUPPERBOUND
'                              PE_RI_INCLUDELOWERBOUND
'                              PE_RI_NOUPPERBOUND
'                              PE_RI_NOLOWERBOUND
'                           Added Types
'                              PEReportOptions
'                              PEParameterValueInfo
'                           Modified Types
'                              PESubreportInfo
'                                  Added Members
'                                      NLinks
'                                      IsOnDemand
'                              PEParameterFieldInfo
'                                  Added Members
'                                      MandatoryPrompt
'                                      MinSize
'                                      MaxSize
'                                      EditMask
'                                      isHidden
'                              PEPrintOptions
'                                  Added Member
'                                      outputFileName
'                           Removed Type
'                              PEParameterInfo
'                           Added Functions
'                              PEGetReportOptions
'                              PESetReportOptions
'                              PESetSectionHeight
'                              PEGetSectionHeight
'                              PEGetNSQLExpressions
'                              PEGetNthSQLExpression
'                              PEGetSQLExpression
'                              PESetSQLExpression
'                              PECheckSQLExpression
'                              PEGetNParameterDefaultValues
'                              PEGetNthParameterDefaultValue
'                              PESetNthParameterDefaultValue
'                              PEAddParameterDefaultValue
'                              PEDeleteNthParameterDefaultValue
'                              PEGetParameterMinMaxValue
'                              PESetParameterMinMaxValue
'                              PEGetParameterValueInfo
'                              PESetParameterValueInfo
'                              PEGetNParameterCurrentValues
'                              PEGetNthParameterCurrentValue
'                              PEAddParameterCurrentValue
'                              PEGetNParameterCurrentRanges
'                              PEGetNthParameterCurrentRange
'                              PEAddParameterCurrentRange
'                              PEGetNthParameterType
'                              PEClearParameterCurrentValuesAndRanges
'                              PEVerifyDatabase
'                              PESetFieldMappingType
'                              PEGetFieldMappingType
'                              PEGetAllowPromptDialog
'                              PESetAllowPromptDialog
'                           Removed Functions
'                              PESetMinimumSectionHeight
'                              PEGetMinimumSectionHeight
'                              PEGetNParams
'                              PEGetNthParam
'                              PEGetNthParamInfo
'                              PESetNthParam
'                           For old code:
'                              PESetMinimumSectionHeight now calls PESetSectionHeight
'                              PEGetMinimumSectionHeight now calls PEGetSectionHeight
'                           Modified PEClosePrintJob Sub to a Function to fix an error in the BAS file
' JWC Feb 23, 1999:
' Added Consts:
'       PE_ERR_PARAMTYPEDIFFERENT
'       PE_ERR_INCONSISTANTRANGETYPES
'       PE_ERR_RANGEORDISCRETE
'       PE_ERR_NOTMAINREPORT
'       PE_ERR_INVALIDCURRENTVALUEINDEX
'       PE_ERR_LINKEDPARAMVALUE
'       PE_ERR_INVALIDSORTMETHODINDEX
'       PE_ERR_INVALIDGRAPHSUBTYPE
'       PE_ERR_BADGRAPHOPTIONINFO
'       PE_ERR_BADGRAPHAXISINFO
'
'       PE_GC_BYSECOND
'       PE_GC_BYMINUTE
'       PE_GC_BYHOUR
'       PE_GC_BYAMPM
'
'       PE_GC_TYPETIME
'
'       PE_GT_BARCHART
'       PE_GT_LINECHART
'       PE_GT_AREACHART
'       PE_GT_PIECHART
'       PE_GT_DOUGHNUTCHART
'       PE_GT_THREEDRISERCHART
'       PE_GT_THREEDSURFACECHART
'       PE_GT_SCATTERCHART
'       PE_GT_RADARCHART
'       PE_GT_BUBBLECHART
'       PE_GT_STOCKCHART
'       PE_GT_USERDEFINEDCHART
'       PE_GT_UNKNOWNTYPECHART
'
'       PE_GST_SIDEBYSIDEBARCHART
'       PE_GST_STACKEDBARCHART
'       PE_GST_PERCENTBARCHART
'       PE_GST_FAKED3DSIDEBYSIDEBARCHART
'       PE_GST_FAKED3DSTACKEDBARCHART
'       PE_GST_FAKED3DPERCENTBARCHART
'       PE_GST_REGULARLINECHART
'       PE_GST_STACKEDLINECHART
'       PE_GST_PERCENTAGELINECHART
'       PE_GST_LINECHARTWITHMARKERS
'       PE_GST_STACKEDLINECHARTWITHMARKERS
'       PE_GST_PERCENTAGELINECHARTWITHMARKERS
'
'       PE_GST_ABSOLUTEAREACHART
'       PE_GST_STACKEDAREACHART
'       PE_GST_PERCENTAREACHART
'       PE_GST_FAKED3DABSOLUTEAREACHART
'       PE_GST_FAKED3DSTACKEDAREACHART
'       PE_GST_FAKED3DPERCENTAREACHART
'
'       PE_GST_REGULARPIECHART
'       PE_GST_FAKED3DREGULARPIECHART
'       PE_GST_MULTIPLEPIECHART
'       PE_GST_MULTIPLEPROPORTIONALPIECHART
'
'       PE_GST_REGULARDOUGHNUTCHART
'       PE_GST_MULTIPLEDOUGHNUTCHART
'       PE_GST_MULTIPLEPROPORTIONALDOUGHNUTCHART
'
'       PE_GST_THREEDREGULARCHART
'       PE_GST_THREEDPYRAMIDCHART
'       PE_GST_THREEDOCTAGONCHART
'       PE_GST_THREEDCUTCORNERSCHART
'
'       PE_GST_THREEDSURFACEREGULARCHART
'       PE_GST_THREEDSURFACEWITHSIDESCHART
'       PE_GST_THREEDSURFACEHONEYCOMBCHART
'
'       PE_GST_XYSCATTERCHART
'       PE_GST_XYSCATTERDUALAXISCHART
'       PE_GST_XYSCATTERWITHLABELSCHART
'       PE_GST_XYSCATTERDUALAXISWITHLABELSCHART
'
'       PE_GST_REGULARRADARCHART
'       PE_GST_STACKEDRADARCHART
'       PE_GST_RADARDUALAXISCHART
'
'       PE_GST_REGULARBUBBLECHART
'       PE_GST_DUALAXISBUBBLECHART
'
'       PE_GST_HIGHLOWCHART
'       PE_GST_HIGHLOWDUALAXISCHART
'       PE_GST_HIGHLOWOPENCHART
'       PE_GST_HIGHLOWOPENDUALAXISCHART
'       PE_GST_HIGHLOWOPENCLOSECHART
'       PE_GST_HIGHLOWOPENCLOSEDUALAXISCHART
'       PE_GST_UNKNOWNSUBTYPECHART
'
'       PE_GTT_TITLE
'       PE_GTT_SUBTITLE
'       PE_GTT_FOOTNOTE
'       PE_GTT_SERIESTITLE
'       PE_GTT_GROUPSTITLE
'       PE_GTT_XAXISTITLE
'       PE_GTT_YAXISTITLE
'       PE_GTT_ZAXISTITLE
'
'       PE_GTF_TITLEFONT
'       PE_GTF_SUBTITLEFONT
'       PE_GTF_FOOTNOTEFONT
'       PE_GTF_GROUPSTITLEFONT
'       PE_GTF_DATATITLEFONT
'       PE_GTF_LEGENDFONT
'       PE_GTF_GROUPLABELSFONT
'       PE_GTF_DATALABELSFONT
'       PE_FACE_NAME_LEN
'
'       PE_GLP_PLACEUPPERRIGHT
'       PE_GLP_PLACEBOTTOMCENTER
'       PE_GLP_PLACETOPCENTER
'       PE_GLP_PLACERIGHT
'       PE_GLP_PLACELEFT
'
'       PE_GBS_MINIMUMBARSIZE
'       PE_GBS_SMALLBARSIZE
'       PE_GBS_AVERAGEBARSIZE
'       PE_GBS_LARGEBARSIZE
'       PE_GBS_MAXIMUMBARSIZE
'
'       PE_GPS_MINIMUMPIESIZE
'       PE_GPS_SMALLPIESIZE
'       PE_GPS_AVERAGEPIESIZE
'       PE_GPS_LARGEPIESIZE
'       PE_GPS_MAXIMUMPIESIZE
'
'       PE_GDPS_NODETACHMENT
'       PE_GDPS_SMALLESTSLICE
'       PE_GDPS_LARGESTSLICE
'
'       PE_GMS_SMALLMARKERS
'       PE_GMS_MEDIUMSMALLMARKERS
'       PE_GMS_MEDIUMMARKERS
'       PE_GMS_MEDIUMLARGEMARKERS
'       PE_GMS_LARGEMARKERS
'
'       PE_GMSP_RECTANGLESHAPE
'       PE_GMSP_CIRCLESHAPE
'       PE_GMSP_DIAMONDSHAPE
'       PE_GMSP_TRIANGLESHAPE
'
'       PE_GCR_COLORCHART
'       PE_GCR_BLACKANDWHITECHART
'
'       PE_GDP_NONE
'       PE_GDP_SHOWLABEL
'       PE_GDP_SHOWVALUE
'
'       PE_GNF_NODECIMAL
'       PE_GNF_ONEDECIMAL
'       PE_GNF_TWODECIMAL
'       PE_GNF_CURRENCYNODECIMAL
'       PE_GNF_CURRENCYTWODECIMAL
'       PE_GNF_PERCENTNODECIMAL
'       PE_GNF_PERCENTONEDECIMAL
'       PE_GNF_PERCENTTWODECIMAL
'
'       PE_GVA_STANDARDVIEW
'       PE_GVA_TALLVIEW
'       PE_GVA_TOPVIEW
'       PE_GVA_DISTORTEDVIEW
'       PE_GVA_SHORTVIEW
'       PE_GVA_GROUPEYEVIEW
'       PE_GVA_GROUPEMPHASISVIEW
'       PE_GVA_FEWSERIESVIEW
'       PE_GVA_FEWGROUPSVIEW
'       PE_GVA_DISTORTEDSTDVIEW
'       PE_GVA_THICKGROUPSVIEW
'       PE_GVA_SHORTERVIEW
'       PE_GVA_THICKSERIESVIEW
'       PE_GVA_THICKSTDVIEW
'       PE_GVA_BIRDSEYEVIEW
'       PE_GVA_MAXVIEW
'
'       PE_OR_NO_SORT
'       PE_OR_ALPHANUMERIC_ASCENDING
'       PE_OR_ALPHANUMERIC_DESCENDING
'       PE_OR_NUMERIC_ASCENDING
'       PE_OR_NUMERIC_DESCENDING
'
'       PE_DR_HASRANGE
'       PE_DR_HASDISCRETE
'       PE_DR_HASDISCRETEANDRANGE
'
'       PE_CONNECTION_BUFFER_LEN
'
'       PE_TCD_OKAY
'       PE_TCD_DATABASENOTFOUND
'       PE_TCD_SERVERNOTFOUND
'       PE_TCD_SERVERNOTOPENED
'       PE_TCD_ALIASCHANGED
'       PE_TCD_INDEXESCHANGED
'       PE_TCD_DRIVERCHANGED
'       PE_TCD_DICTIONARYCHANGED
'       PE_TCD_FILETYPECHANGED
'       PE_TCD_RECORDSIZECHANGED
'       PE_TCD_ACCESSCHANGED
'       PE_TCD_PARAMETERSCHANGED
'       PE_TCD_LOCATIONCHANGED
'       PE_TCD_DATABASEOTHER
'       PE_TCD_NUMFIELDSCHANGED
'       PE_TCD_FIELDOTHER
'       PE_TCD_FIELDNAMECHANGED
'       PE_TCD_FIELDDESCCHANGED
'       PE_TCD_FIELDTYPECHANGED
'       PE_TCD_FIELDSIZECHANGED
'       PE_TCD_NATIVEFIELDTYPECHANGED
'       PE_TCD_NATIVEFIELDOFFSETCHANGED
'       PE_TCD_NATIVEFIELDSIZECHANGED
'       PE_TCD_FIELDDECPLACESCHANGED
'
'       PE_SIZEOF_GRAPH_TYPE_INFO
'       PE_SIZEOF_FONT_COLOR_INFO
'       PE_SIZEOF_GRAPH_OPTION_INFO
'       PE_SIZEOF_GRAPH_AXIS_INFO
'       PE_SIZEOF_PICK_LIST_OPTION
'       PE_SIZEOF_TABLE_DIFFERENCE_INFO
'
' Modified Const values:
'       PE_SIZEOF_REPORT_OPTIONS
'       PE_SIZEOF_TABLE_LOCATION
'       PE_SIZEOF_WINDOW_OPTIONS
'
' Modified Types:
'       PEReportOptions
'               Added member:
'               doAsyncQuery
'       PETableLocation
'               Added members:
'               SubLocation
'               ConnectBuffer
'       PEWindowOptions
'               Added members:
'               showToolbarTips
'               showDocumentTips
'
' Added Type(s):
'       PEGraphTypeInfo
'       PEFontColorInfo
'       PEGraphOptionInfo
'       PEGraphAxisInfo
'       PEParameterPickListOption
'       PETableDifferenceInfo
'
' Added Declarations:
'       PEGetGraphTypeInfo()
'       PESetGraphTypeInfo()
'       PEGetGraphTextInfo()
'       PESetGraphTextInfo()
'       PEGetGraphFontInfo()
'       PESetGraphFontInfo()
'       PEGetGraphOptionInfo()
'       PESetGraphOptionInfo()
'
'       PEGetNthParameterValueDescription()
'       PESetNthParameterValueDescription()
'       PEGetParameterPickListOption()
'       PESetParameterPickListOption()
'
'       PECheckNthTableDifferences()
'
' Moved Declarations to obs32.bas:
'       PEGetGraphType()
'       PEGetGraphData()
'       PEGetGraphText()
'       PEGetGraphOptions()
'       PESetGraphType()
'       PESetGraphData()
'       PESetGraphText()
'       PESetGraphOptions()
' Moved Consts to obs32.bas:
'       PE_SIDE_BY_SIDE_BAR_GRAPH
'       PE_STACKED_BAR_GRAPH
'       PE_PERCENT_BAR_GRAPH
'       PE_FAKED_3D_SIDE_BY_SIDE_BAR_GRAPH
'       PE_FAKED_3D_STACKED_BAR_GRAPH
'       PE_FAKED_3D_PERCENT_BAR_GRAPH
'       PE_PIE_GRAPH
'       PE_MULTIPLE_PIE_GRAPH
'       PE_PROPORTIONAL_MULTI_PIE_GRAPH
'       PE_LINE_GRAPH
'       PE_AREA_GRAPH
'       PE_THREED_BAR_GRAPH
'       PE_USER_DEFINED_GRAPH
'       PE_UNKNOWN_TYPE_GRAPH
'
'       PE_GRAPH_ROWS_ONLY
'       PE_GRAPH_COLS_ONLY
'       PE_GRAPH_MIXED_ROW_COL
'       PE_GRAPH_MIXED_COL_ROW
'       PE_GRAPH_UNKNOWN_DIRECTION
'
'       PE_GRAPH_DATA_NULL_SELECTION
'       PE_GRAPH_TEXT_LEN
'
'       PE_SIZEOF_GRAPH_DATA_INFO
'       PE_SIZEOF_GRAPH_TEXT_INFO
'       PE_SIZEOF_GRAPH_OPTIONS
'
'       PE_SIZEOF_PRINT_FILE_OPTIONS
'       PE_SIZEOF_CHAR_SEP_FILE_OPTIONS
'
' Moved types to obs32.bas
'       PEGraphDataInfo
'       PEGraphTextInfo
'       PEGraphOptions
'
'       PEPrintFileOptions
'       PECharSepFileOptions
'
' BCK - Added consts (For 7.0 MR1)
'//axis division method
'       PE_ADM_AUTOMATIC                0
'       PE_ADM_MANUAL                   1
'
'       PE_ERR_INVALIDPARAMETERRANGEINFO    672 // Invalid PE_RI_* combination.
'
'       PE_GRAPH_TEXT_LEN
'
' SM - Updated for 8.0 Nov 11, 1999
'
' Added error constants:
' PE_ERR_NOTEXTERNALSUBREPORT
' PE_ERR_INVALIDPARAMETERVALUE
' PE_ERR_INVALIDFORMULASYNTAXTYPE
' PE_ERR_INVALIDCROPVALUE
' PE_ERR_INVALIDCOLLATIONVALUE
' PE_ERR_STARTPAGEGREATERSTOPPAGE
'
' Added PEReportOptions constants for promptMode member
' PE_RPTOPT_PROMPT_NONE
' PE_RPTOPT_PROMPT_NORMAL
' PE_RPTOPT_PROMPT_ALWAYS

' Modified type PEReportOptions
' - Added members:
'       promptMode
'       selectDistinctRecords
'
' Updated struct size const PE_SIZEOF_REPORT_OPTIONS
'
' Modified type PEGroupOptions
' - Added members
'       hierarchicalSorting
'       instanceIDField [PE_FIELD_NAME_LEN]
'       parentIDField [PE_FIELD_NAME_LEN]
'       groupIndent
'
' Updated struct size const PE_SIZEOF_GROUP_OPTIONS
'
' Added function
' - PEGetNSectionsInArea()
'
' Modified type PESectionOptions
' - Added member
'       reserveMinimumPageFooter
'
' Updated struct size const PE_SIZEOF_SECTION_OPTIONS
'
' Added constants:
' PE_SRI_UNDEFINED
' PE_SRI_ONOPENJOB
' PE_SRI_ONFUNCTIONCALL
'
' Modified type PESubreportInfo
' Added members:
'       external
'       reimportOption
'
' Updated struct size const PE_SIZEOF_SUBREPORT_INFO
'
' Modified type PEFontColorInfo
' Added member:
'       twipSize
'
' Updated struct size const PE_SIZEOF_FONT_COLOR_INFO
'
' Added functions:
'       PESetGraphTextDefaultOption
'       PEGetGraphTextDefaultOption
'
' Added constants:
' PE_GLL_PERCENTAGE
' PE_GLL_AMOUNT
' PE_GLL_CUSTOM
'
' Modified type PEGraphOptionInfo
' Added member:
'       legendLayout
'
' Updated struct size const PE_SIZEOF_GRAPH_OPTION_INFO
'
' Modified type PEGraphAxisInfo
' Added members:
'       dataAxisYAutoScale
'       dataAxisY2AutoScale
'       seriesAxisAutoScale
'
' Updated struct size const PE_SIZEOF_GRAPH_AXIS_INFO
'
' Modified type PEWindowOptions
' Added member:
'       hasLaunchButton
'
' Updated struct size const PE_SIZEOF_WINDOW_OPTIONS
'
' Added function:
'       PEFreeDevMode
'
' Modified type PEReportSummaryInfo
' Added member:
'       savePreviewPicture
'
' Updated struct size const PE_SIZEOF_REPORT_SUMMARY_INFO
'
' Added consts:
' PE_FS_SIZE
' PE_FST_CRYSTAL
' PE_FST_BASIC
'
' Added type PEFormulaSyntax
' Added struct size const PE_SIZEOF_FORMULA_SYNTAX
' Added functions:
'       PESetFormulaSyntax
'       PEGetFormulaSyntax
'
'

Global Const PE_UNCHANGED_COLOR = -2

Global Const PE_ERR_NOERROR = 0

Global Const PE_ERR_NOTENOUGHMEMORY = 500
Global Const PE_ERR_INVALIDJOBNO = 501
Global Const PE_ERR_INVALIDHANDLE = 502
Global Const PE_ERR_STRINGTOOLONG = 503
Global Const PE_ERR_NOSUCHREPORT = 504
Global Const PE_ERR_NODESTINATION = 505
Global Const PE_ERR_BADFILENUMBER = 506
Global Const PE_ERR_BADFILENAME = 507
Global Const PE_ERR_BADFIELDNUMBER = 508
Global Const PE_ERR_BADFIELDNAME = 509
Global Const PE_ERR_BADFORMULANAME = 510
Global Const PE_ERR_BADSORTDIRECTION = 511
Global Const PE_ERR_ENGINENOTOPEN = 512
Global Const PE_ERR_INVALIDPRINTER = 513
Global Const PE_ERR_PRINTFILEEXISTS = 514
Global Const PE_ERR_BADFORMULATEXT = 515
Global Const PE_ERR_BADGROUPSECTION = 516
Global Const PE_ERR_ENGINEBUSY = 517
Global Const PE_ERR_BADSECTION = 518
Global Const PE_ERR_NOPRINTWINDOW = 519
Global Const PE_ERR_JOBALREADYSTARTED = 520
Global Const PE_ERR_BADSUMMARYFIELD = 521
Global Const PE_ERR_NOTENOUGHSYSRES = 522
Global Const PE_ERR_BADGROUPCONDITION = 523
Global Const PE_ERR_JOBBUSY = 524
Global Const PE_ERR_BADREPORTFILE = 525
Global Const PE_ERR_NODEFAULTPRINTER = 526
Global Const PE_ERR_SQLSERVERERROR = 527
Global Const PE_ERR_BADLINENUMBER = 528
Global Const PE_ERR_DISKFULL = 529
Global Const PE_ERR_FILEERROR = 530
Global Const PE_ERR_INCORRECTPASSWORD = 531
Global Const PE_ERR_BADDATABASEDLL = 532
Global Const PE_ERR_BADDATABASEFILE = 533
Global Const PE_ERR_ERRORINDATABASEDLL = 534
Global Const PE_ERR_DATABASESESSION = 535
Global Const PE_ERR_DATABASELOGON = 536
Global Const PE_ERR_DATABASELOCATION = 537
Global Const PE_ERR_BADSTRUCTSIZE = 538
Global Const PE_ERR_BADDATE = 539
Global Const PE_ERR_BADEXPORTDLL = 540
Global Const PE_ERR_ERRORINEXPORTDLL = 541
Global Const PE_ERR_PREVATFIRSTPAGE = 542
Global Const PE_ERR_NEXTATLASTPAGE = 543
Global Const PE_ERR_CANNOTACCESSREPORT = 544
Global Const PE_ERR_USERCANCELLED = 545
Global Const PE_ERR_OLE2NOTLOADED = 546
Global Const PE_ERR_BADCROSSTABGROUP = 547
Global Const PE_ERR_NOCTSUMMARIZEDFIELD = 548
Global Const PE_ERR_DESTINATIONNOTEXPORT = 549
Global Const PE_ERR_INVALIDPAGENUMBER = 550
Global Const PE_ERR_NOTSTOREDPROCEDURE = 552
Global Const PE_ERR_INVALIDPARAMETER = 553
Global Const PE_ERR_GRAPHNOTFOUND = 554
Global Const PE_ERR_INVALIDGRAPHTYPE = 555
Global Const PE_ERR_INVALIDGRAPHDATA = 556
Global Const PE_ERR_CANNOTMOVEGRAPH = 557
Global Const PE_ERR_INVALIDGRAPHTEXT = 558
Global Const PE_ERR_INVALIDGRAPHOPT = 559

'New Error Codes For 5.0
Global Const PE_ERR_BADSECTIONHEIGHT = 560
Global Const PE_ERR_BADVALUETYPE = 561
Global Const PE_ERR_INVALIDSUBREPORTNAME = 562
Global Const PE_ERR_NOPARENTWINDOW = 564     'dialog parent window
Global Const PE_ERR_INVALIDZOOMFACTOR = 565  'zoom factor
Global Const PE_ERR_PAGESIZEOVERFLOW = 567
Global Const PE_ERR_LOWSYSTEMRESOURCES = 568
Global Const PE_ERR_BADGROUPNUMBER = 570
Global Const PE_ERR_INVALIDNEGATIVEVALUE = 572
Global Const PE_ERR_INVALIDMEMORYPOINTER = 573
Global Const PE_ERR_INVALIDPARAMETERNUMBER = 594
Global Const PE_ERR_SQLSERVERNOTOPENED = 599

Global Const PE_ERR_NOTIMPLEMENTED = 999

'new constants for 7.0
Global Const PE_ERR_INVALIDOBJECTFORMATNAME = 571
Global Const PE_ERR_INVALIDOBJECTTYPE = 574
Global Const PE_ERR_INVALIDGRAPHDATATYPE = 577
Global Const PE_ERR_INVALIDSUBREPORTLINKNUMBER = 582
Global Const PE_ERR_SUBREPORTLINKEXIST = 583
Global Const PE_ERR_BADROWCOLVALUE = 584
Global Const PE_ERR_INVALIDSUMMARYNUMBER = 585
Global Const PE_ERR_INVALIDGRAPHDATAFIELDNUMBER = 586
Global Const PE_ERR_INVALIDSUBREPORTNUMBER = 587
Global Const PE_ERR_INVALIDFIELDSCOPE = 588
Global Const PE_ERR_FIELDINUSE = 590
Global Const PE_ERR_INVALIDPAGEMARGINS = 595
Global Const PE_ERR_REPORTONSECUREQUERY = 596
Global Const PE_ERR_CANNOTOPENSECUREQUERY = 597
Global Const PE_ERR_INVALIDSECTIONNUMBER = 598
Global Const PE_ERR_TABLENAMEEXIST = 606
Global Const PE_ERR_INVALIDCURSOR = 607
Global Const PE_ERR_FIRSTPASSNOTFINISHED = 608
Global Const PE_ERR_CREATEDATASOURCE = 609
Global Const PE_ERR_CREATEDRILLDOWNPARAMETERS = 610
Global Const PE_ERR_CHECKFORDATASOURCECHANGES = 613
Global Const PE_ERR_STARTBACKGROUNDPROCESSING = 614
Global Const PE_ERR_SQLSERVERINUSE = 619
Global Const PE_ERR_GROUPSORTFIELDNOTSET = 620
Global Const PE_ERR_CANNOTSETSAVESUMMARIES = 621
Global Const PE_ERR_LOADOLAPDATABASEMANAGER = 622
Global Const PE_ERR_OPENOLAPCUBE = 623
Global Const PE_ERR_READOLAPCUBEDATA = 624
Global Const PE_ERR_CANNOTSAVEQUERY = 626
Global Const PE_ERR_CANNOTREADQUERYDATA = 627
Global Const PE_ERR_MAINREPORTFIELDLINKED = 629
Global Const PE_ERR_INVALIDMAPPINGTYPEVALUE = 630
Global Const PE_ERR_HITTESTFAILED = 636
Global Const PE_ERR_BADSQLEXPRESSIONNAME = 637        ' no SQL expression by the specified *name* exists in this report.
Global Const PE_ERR_BADSQLEXPRESSIONNUMBER = 638      ' no SQL expression by the specified *number* exists in this report.
Global Const PE_ERR_BADSQLEXPRESSIONTEXT = 639        ' not a valid SQL expression
Global Const PE_ERR_INVALIDDEFAULTVALUEINDEX = 641    ' invalid index for default value of a parameter.
Global Const PE_ERR_NOMINMAXVALUE = 642               ' the specified PE_PF_* type does not have min/max values.
Global Const PE_ERR_INCONSISTANTTYPES = 643           ' if both min and max values are specified in PESetParameterMinMaxValue,
                                                      ' the value types for the min and max must be the same.
Global Const PE_ERR_CANNOTLINKTABLES = 645
Global Const PE_ERR_CREATEROUTER = 646
Global Const PE_ERR_INVALIDFIELDINDEX = 647
Global Const PE_ERR_INVALIDGRAPHTITLETYPE = 648
Global Const PE_ERR_INVALIDGRAPHTITLEFONTTYPE = 649

'end new for 7.0 section

' new consts for 7.0, MR1
Global Const PE_ERR_PARAMTYPEDIFFERENT = 650 ' the type used in a add/set value API for a
                                             ' parameter differs with it's existing type.
Global Const PE_ERR_INCONSISTANTRANGETYPES = 651 ' the value type for both start & end range
                                                 ' values must be the same.
Global Const PE_ERR_RANGEORDISCRETE = 652       ' an operation was attempted on a discrete parameter that is
                                                ' only legal for range parameters or vice versa.

Global Const PE_ERR_NOTMAINREPORT = 654         ' an operation was attempted that is disallowed for subreports.

Global Const PE_ERR_INVALIDCURRENTVALUEINDEX = 655 ' invalid index for current value of a parameter.
Global Const PE_ERR_LINKEDPARAMVALUE = 656 ' operation illegal on linked parameter.

Global Const PE_ERR_INVALIDPARAMETERRANGEINFO = 672 ' Invalid PE_RI_* combination.

Global Const PE_ERR_INVALIDSORTMETHODINDEX = 674 ' Invalid sort method index.

Global Const PE_ERR_INVALIDGRAPHSUBTYPE = 675   ' Invalid PE_GST_* or
                                                ' PE_GST_* does not match PE_GT_* or
                                                ' PE_GST_* current graph type.
Global Const PE_ERR_BADGRAPHOPTIONINFO = 676    ' one of them members of PEGraphOptionInfo is out of range.
Global Const PE_ERR_BADGRAPHAXISINFO = 677      ' one of them members of PEGraphAxisInfo is out of range.

Global Const PE_ERR_NOTEXTERNALSUBREPORT = 680  ' the subreport is not imported.
Global Const PE_ERR_INVALIDPARAMETERVALUE = 687
Global Const PE_ERR_INVALIDFORMULASYNTAXTYPE = 688 ' specified formula syntax not in PE_FST_*
Global Const PE_ERR_INVALIDCROPVALUE = 689
Global Const PE_ERR_INVALIDCOLLATIONVALUE = 690
Global Const PE_ERR_STARTPAGEGREATERSTOPPAGE = 691
Global Const PE_ERR_INVALIDEXPORTFORMAT = 692

Global Const PE_ERR_OTHERERROR = 997
Global Const PE_ERR_INTERNALERROR = 998               ' programming error

'Constants using to calculate structure size constants
Global Const PE_BYTE_LEN = 1
Global Const PE_WORD_LEN = 2
Global Const PE_LONG_LEN = 4
Global Const PE_DOUBLE_LEN = 8

' Open, print and close report (used when no changes needed to report)
' --------------------------------------------------------------------

Declare Function PEPrintReport Lib "crpe32.dll" (ByVal RptName$, ByVal Printer%, ByVal Window%, ByVal Title$, ByVal Lft&, ByVal Top&, ByVal Wdth&, ByVal height&, ByVal style As Long, ByVal PWindow As Long) As Integer


' Open and close print engine
' ---------------------------

Declare Function PEOpenEngine Lib "crpe32.dll" () As Integer
Declare Sub PECloseEngine Lib "crpe32.dll" ()
Declare Function PECanCloseEngine Lib "crpe32.dll" () As Integer


' Get version info
' ----------------

Global Const PE_GV_DLL = 100      ' values for version parameter of PEGetVersion
Global Const PE_GV_ENGINE = 200

Declare Function PEGetVersion Lib "crpe32.dll" (ByVal version%) As Integer


' Open and close print job (i.e. report)
' --------------------------------------

Declare Function PEOpenPrintJob Lib "crpe32.dll" (ByVal RptName$) As Integer

Declare Function PEClosePrintJob Lib "crpe32.dll" (ByVal printJob%) As Integer


' Start and cancel print job (i.e. print the report, usually after changing report)
' ---------------------------------------------------------------------------------

Declare Function PEStartPrintJob Lib "crpe32.dll" (ByVal printJob%, ByVal WaitOrNot%) As Integer

Declare Sub PECancelPrintJob Lib "crpe32.dll" (ByVal printJob%)


' report options
'----------------
Global Const PE_RPTOPT_CVTDATETIMETOSTR = 0
Global Const PE_RPTOPT_CVTDATETIMETODATE = 1
Global Const PE_RPTOPT_KEEPDATETIMETYPE = 2

' Following are the valid values for promptMode
Global Const PE_RPTOPT_PROMPT_NONE = 0
Global Const PE_RPTOPT_PROMPT_NORMAL = 1
Global Const PE_RPTOPT_PROMPT_ALWAYS = 2

Type PEReportOptions
    StructSize As Integer                       'initialize to PE_SIZEOF_REPORT_OPTIONS
    saveDataWithReport As Integer               'BOOL value, except use PE_UNCHANGED for no change
    saveSummariesWithReport As Integer          'BOOL value, except use PE_UNCHANGED for no change
    useIndexForSpeed As Integer                 'BOOL value, except use PE_UNCHANGED for no change
    translateDOSStrings As Integer              'BOOL value, except use PE_UNCHANGED for no change
    translateDosMemos As Integer                'BOOL value, except use PE_UNCHANGED for no change
    convertDateTimeType As Integer              'a PE_RPTOPT_ value, except use PE_UNCHANGED for no change
    convertNullFieldToDefault As Integer        'BOOL value, except use PE_UNCHANGED for no change
    morePrintEngineErrorMessages As Integer     'BOOL value, except use PE_UNCHANGED for no change
    caseInsensitiveSQLData As Integer           'BOOL value, except use PE_UNCHANGED for no change
    verifyOnEveryPrint As Integer               'BOOL value, except use PE_UNCHAGED for no change
    zoomMode As Integer                         'a PE_ZOOM_ constant, except use PE_UNCHANGED for no change
    hasGroupTree As Integer                     'BOOL value, except use PE_UNCHANGED for no change
    dontGenerateDataForHiddenObjects As Integer ' BOOL value, except use PE_UNCHANGED for no change
    performGroupingOnServer As Integer          ' BOOL value, except use PE_UNCHANGED for no change
    doAsyncQuery As Integer                     ' BOOL value, except use PE_UNCHANGED for no change
    promptMode As Integer                       ' PE_RPTOPT_PROMPT_NONE, PE_RPTOPT_PROMPT_NORMAL, PE_RPTOPT_PROMPT_ALWAYS, use PE_UNCHANGED for no change
    SelectDistinctRecords As Integer            ' BOOL value, except use PE_UNCHANGED for no change
End Type

Global Const PE_SIZEOF_REPORT_OPTIONS = 18 * PE_WORD_LEN

Declare Function PEGetReportOptions Lib "crpe32.dll" (ByVal printJob%, reportOptions As PEReportOptions) As Integer

Declare Function PESetReportOptions Lib "crpe32.dll" (ByVal printJob%, reportOptions As PEReportOptions) As Integer


' Print job status
' ----------------
 
Declare Function PEIsPrintJobFinished Lib "crpe32.dll" (ByVal printJob%) As Integer

' To work around the problem of 4 - Byte alignment the PEGetJobStatus
' call has been re-declared locally. When your application calls PEGetJobStatus
' it is calling a function in this file which in turn calls CRPE32.DLL.
Global Const PE_JOBNOTSTARTED = 1
Global Const PE_JOBINPROGRESS = 2
Global Const PE_JOBCOMPLETED = 3
Global Const PE_JOBFAILED = 4
Global Const PE_JOBCANCELLED = 5
Global Const PE_JOBHALTED = 6   ' too many records or too much time

Type PEJobInfo
    StructSize As Integer  ' initialize to PE_SIZEOF_JOB_INFO

    NumRecordsRead As Long
    NumRecordsSelected As Long
    NumRecordsPrinted As Long

    DisplayPageN As Integer
    LatestPageN As Integer
    StartPageN As Integer

    PrintEnded As Long
End Type

Type SplitPEJobInfo
    StructSize As Integer  ' initialize to PE_SIZEOF_JOB_INFO

    NumRecordsRead1 As Integer
    NumRecordsRead2 As Integer
    NumRecordsSelected1 As Integer
    NumRecordsSelected2 As Integer
    NumRecordsPrinted1 As Integer
    NumRecordsPrinted2 As Integer

    DisplayPageN As Integer
    LatestPageN As Integer
    StartPageN As Integer

    PrintEnded As Long
End Type

Global Const PE_SIZEOF_JOB_INFO = 10 * PE_WORD_LEN + 4

Declare Function RealPEGetJobStatus Lib "crpe32.dll" Alias "PEGetJobStatus" (ByVal printJob%, JobInfo As SplitPEJobInfo) As Integer

' Controlling dialogs
' -------------------

Declare Function PESetDialogParentWindow Lib "crpe32.dll" (ByVal printJob%, ByVal parentWindow As Long) As Integer

Declare Function PEEnableProgressDialog Lib "crpe32.dll" (ByVal printJob%, ByVal enable%) As Integer

' Controlling Paremeter Field Prompting Dialog
' --------------------------------------------

' Set boolean to indicate whether CRPE is allowed to prompt for parameter values
' during printing.

Declare Function PEGetAllowPromptDialog Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PESetAllowPromptDialog Lib "crpe32.dll" (ByVal printJob%, ByVal showPromptDialog%) As Integer


' Print job error codes and messages
' ----------------------------------

Declare Function PEGetErrorCode Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetErrorText Lib "crpe32.dll" (ByVal printJob%, TextHandle As Long, TextLength%) As Integer

Declare Function PEGetHandleString Lib "crpe32.dll" (ByVal TextHandle As Long, ByVal Buffer$, ByVal BufferLength%) As Integer


' Setting the print date
' ----------------------

Declare Function PEGetPrintDate Lib "crpe32.dll" (ByVal printJob%, Date_Year%, Date_Month%, Date_Day%) As Integer

Declare Function PESetPrintDate Lib "crpe32.dll" (ByVal printJob%, ByVal Date_Year%, ByVal Date_Month%, ByVal Date_Day%) As Integer

' Encoding and Decoding Section Codes
' -----------------------------------

Global Const PE_ALLSECTIONS = 0

'Section types for use with PE_SECTION_CODE, PE_SECTION_TYPE, PE_GROUP_N and PE_SECTION_N functions
Global Const PE_SECT_PAGE_HEADER = 2
Global Const PE_SECT_PAGE_FOOTER = 7
Global Const PE_SECT_REPORT_HEADER = 1
Global Const PE_SECT_REPORT_FOOTER = 8
Global Const PE_SECT_GROUP_HEADER = 3
Global Const PE_SECT_GROUP_FOOTER = 5
Global Const PE_SECT_DETAIL = 4

'The old section constants with comment showing them in terms of the new:
'(Note that PE_GRANDTOTALSECTION and PE_SUMMARYSECTION both map
' to PE_SECT_REPORT_FOOTER.)

Global Const PE_HEADERSECTION = 2000  'PE_SECTION_CODE (PE_SECT_PAGE_HEADER,   0, 0)
Global Const PE_FOOTERSECTION = 7000  'PE_SECTION_CODE (PE_SECT_PAGE_FOOTER,   0, 0)
Global Const PE_TITLESECTION = 1000   'PE_SECTION_CODE (PE_SECT_REPORT_HEADER, 0, 0)
Global Const PE_SUMMARYSECTION = 8000 'PE_SECTION_CODE (PE_SECT_REPORT_FOOTER, 0, 0)
Global Const PE_GROUPHEADER = 3000    'PE_SECTION_CODE (PE_SECT_GROUP_HEADER,  0, 0)
Global Const PE_GROUPFOOTER = 5000    'PE_SECTION_CODE (PE_SECT_GROUP_FOOTER,  0, 0)
Global Const PE_DETAILSECTION = 4000  'PE_SECTION_CODE (PE_SECT_DETAIL,        0, 0)
Global Const PE_GRANDTOTALSECTION = PE_SUMMARYSECTION


' Controlling group conditions (i.e. group breaks)
' ------------------------------------------------
Global Const PE_SF_MAX_NAME_LENGTH = 50

Global Const PE_SF_DESCENDING = 0
Global Const PE_SF_ASCENDING = 1
Global Const PE_SF_ORIGINAL = 2 'only for group condition
Global Const PE_SF_SPECIFIED = 3 'only for group condition

' use PE_ANYCHANGE for all field types except Date
Global Const PE_GC_ANYCHANGE = 0

' use these constants for Date and DateTime fields
Global Const PE_GC_DAILY = 0
Global Const PE_GC_WEEKLY = 1
Global Const PE_GC_BIWEEKLY = 2
Global Const PE_GC_SEMIMONTHLY = 3
Global Const PE_GC_MONTHLY = 4
Global Const PE_GC_QUARTERLY = 5
Global Const PE_GC_SEMIANNUALLY = 6
Global Const PE_GC_ANNUALLY = 7

' use these constants for Time  and DateTime fields
Global Const PE_GC_BYSECOND = 8
Global Const PE_GC_BYMINUTE = 9
Global Const PE_GC_BYHOUR = 10
Global Const PE_GC_BYAMPM = 11

' use these constants for Boolean fields
Global Const PE_GC_TOYES = 1
Global Const PE_GC_TONO = 2
Global Const PE_GC_EVERYYES = 3
Global Const PE_GC_EVERYNO = 4
Global Const PE_GC_NEXTISYES = 5
Global Const PE_GC_NEXTISNO = 6

Declare Function PESetGroupCondition Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal ConditionField$, ByVal condition%, ByVal SortDirection%) As Integer

Declare Function PEGetNGroups Lib "crpe32.dll" (ByVal printJob%) As Integer

' for PEGetGroupCondition, Condition% encodes both
' the condition and the type of the condition field
Global Const PE_GC_CONDITIONMASK = &HFF
Global Const PE_GC_TYPEMASK = &HF00

Global Const PE_GC_TYPEOTHER = &H0
Global Const PE_GC_TYPEDATE = &H200
Global Const PE_GC_TYPEBOOLEAN = &H400
Global Const PE_GC_TYPETIME = &H800

Declare Function PEGetGroupCondition Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ConditionFieldHandle As Long, ConditionFieldLength%, condition%, SortDirection%) As Integer

Global Const PE_FIELD_NAME_LEN = 512

Global Const PE_GO_TBN_ALL_GROUPS_UNSORTED = 0
Global Const PE_GO_TBN_ALL_GROUPS_SORTED = 1
Global Const PE_GO_TBN_TOP_N_GROUPS = 2
Global Const PE_GO_TBN_BOTTOM_N_GROUPS = 3

Type PEGroupOptions
    StructSize As Integer
    'when setting, pass a PE_GC_ constant, or PE_UNCHANGED for no change.
    'when getting, use PE_GC_TYPEMASK and PE_GC_CONDITIONMASK to decode the condition.
    condition As Integer
    fieldName As String * PE_FIELD_NAME_LEN ' formula form, or empty for no change.
    SortDirection As Integer                ' a PE_SF_ const, or PE_UNCHANGED for no change.
    repeatGroupHeader As Integer            ' BOOL value, or PE_UNCHANGED for no change.
    keepGroupTogether As Integer            ' BOOL value, or PE_UNCHANGED for no change.
    topOrBottomNGroups As Integer           ' a PE_GO_TBN_ constant, or PE_UNCHANGED for no change.
    topOrBottomNSortFieldName As String * PE_FIELD_NAME_LEN ' formula form, or empty for no change.
    ntoporbottomgroups As Integer           ' the number of groups to keep, 0 for all, or PE_UNCHANGED for no change.
    discardOtherGroups As Integer           ' BOOL value, or PE_UNCHANGED for no change.
    ignored As Integer                     ' for 4 byte alignment ignored.
    hierarchicalSorting As Integer              ' Boolean or PE_UNCHANGED
    instanceIDField As String * PE_FIELD_NAME_LEN       ' for hierarchical grouping
    parentIDField As String * PE_FIELD_NAME_LEN         ' for hierarchical grouping
    groupIndent As Long                         ' twips
End Type

Global Const PE_SIZEOF_GROUP_OPTIONS = 10 * PE_WORD_LEN + _
                                       4 * PE_FIELD_NAME_LEN + _
                                       1 * PE_LONG_LEN

Declare Function PEGetGroupOptions Lib "crpe32.dll" (ByVal printJob%, ByVal groupN%, groupOptions As PEGroupOptions) As Integer
                                
Declare Function PESetGroupOptions Lib "crpe32.dll" (ByVal printJob%, ByVal groupN%, groupOptions As PEGroupOptions) As Integer

' Controlling formulas, selection formulas and group selection formulas
' ---------------------------------------------------------------------

Declare Function PEGetNFormulas Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthFormula Lib "crpe32.dll" (ByVal printJob%, ByVal FormulaN%, nameHandle As Long, nameLength%, TextHandle As Long, TextLength%) As Integer

Declare Function PEGetFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaName$, TextHandle As Long, TextLength%) As Integer

Declare Function PESetFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaName$, ByVal formulaString$) As Integer

Declare Function PECheckFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaName$) As Integer

Declare Function PEGetSelectionFormula Lib "crpe32.dll" (ByVal printJob%, TextHandle As Long, TextLength%) As Integer

Declare Function PESetSelectionFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaString$) As Integer

Declare Function PECheckSelectionFormula Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetGroupSelectionFormula Lib "crpe32.dll" (ByVal printJob%, TextHandle As Long, TextLength%) As Integer

Declare Function PESetGroupSelectionFormula Lib "crpe32.dll" (ByVal printJob%, ByVal formulaString$) As Integer

Declare Function PECheckGroupSelectionFormula Lib "crpe32.dll" (ByVal printJob%) As Integer


' SQL Expressions
'-----------------

Declare Function PEGetNSQLExpressions Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthSQLExpression Lib "crpe32.dll" (ByVal printJob%, ByVal expressionN%, nameHandle&, nameLength%, TextHandle&, TextLength%) As Integer

Declare Function PEGetSQLExpression Lib "crpe32.dll" (ByVal printJob%, ByVal expressionName$, TextHandle&, TextLength%) As Integer

Declare Function PESetSQLExpression Lib "crpe32.dll" (ByVal printJob%, ByVal expressionName$, ByVal expressionString$) As Integer

Declare Function PECheckSQLExpression Lib "crpe32.dll" (ByVal printJob%, ByVal expressionName$) As Integer

'********************************************************************************
' NOTE : Stored Procedures
'
' The previous Stored Procedure API Calls PEGetNParams, PEGetNthParam,
' PEGetNthParamInfo and PESetNthParam have been made obsolete.  Older
' applications that used these API Calls will still work as before, but for new
' development please use the new Parameter API calls below,
'
' The Stored Procedure Parameters have now been unified with the Parameter
' Fields.
'
' The replacements for these calls are as follows :
'      PEGetNParams        = PEGetNParameterFields
'      PEGetNthParam       = PEGetNthParameterField
'      PEGetNthParamInfo   = PEGetParameterValueInfo
'      PESetNthParam       = PESetNthParameterField
'
' NOTE : To tell if a Parameter Field is a Stored Procedure, use the
'        PEGetNthParameterType or PEGetNthParameterField API Calls
'
' If you wish to SET a parameter to NULL then set the CurrentValue to CRWNULL.
' The CRWNULL is of Type String and is independant of the datatype of the
' parameter.
'
'********************************************************************************

' Controlling Parameter Fields
' ----------------------------

Global Const PE_PF_REPORT_NAME_LEN = 128
Global Const PE_PF_NAME_LEN = 256
Global Const PE_PF_PROMPT_LEN = 256
Global Const PE_PF_VALUE_LEN = 256
Global Const PE_PF_EDITMASK_LEN = 256

Global Const PE_PF_NUMBER = 0
Global Const PE_PF_CURRENCY = 1
Global Const PE_PF_BOOLEAN = 2
Global Const PE_PF_DATE = 3
Global Const PE_PF_STRING = 4
Global Const PE_PF_DATETIME = 5
Global Const PE_PF_TIME = 6

Type PEParameterFieldInfo
    'Initialize to PE_SIZEOF_PARAMETER_FIELD_INFO.
    StructSize As Integer

    'PE_PF_ constant
    valueType As Integer

    'Indicate the default value is set in PEParameterFieldInfo.
    DefaultValueSet As Integer

    'Indicate the current value is set in PEParameterFieldInfo.
    CurrentValueSet As Integer

    'All strings are null-terminated.
    Name As String * PE_PF_NAME_LEN
    Prompt As String * PE_PF_PROMPT_LEN

    ' Could be Number, Date, DateTime, Time, Boolean, or String
    DefaultValue As String * PE_PF_VALUE_LEN
    currentValue As String * PE_PF_VALUE_LEN

    'name of report where the field belongs, only used in PEGetNthParameterField
    reportName As String * PE_PF_REPORT_NAME_LEN

    'returns false (0) if parameter is linked, not in use, or has current value set
    needsCurrentValue As Integer
    
    'for String values this will be TRUE if the string is limited on length, for
    'other types it will be TRUE if the parameter is limited by a range
    isLimited As Integer

    ' For string fields, these are the minimum/maximum length of the string.
    ' For numeric fields, they are the minimum/maximum numeric value.
    ' For other fields, use PEGetParameterMinMaxValue
    MinSize As Double
    MaxSize As Double

    'An edit mask that restricts what may be entered for string parameters.
    EditMask As String * PE_PF_EDITMASK_LEN

    'return true if it is essbase sub var
    isHidden As Integer
    
End Type
    
Global Const PE_SIZEOF_VARINFO_TYPE = 5 * PE_WORD_LEN + PE_PF_NAME_LEN + PE_PF_PROMPT_LEN + 2 * PE_PF_VALUE_LEN + PE_PF_REPORT_NAME_LEN
Global Const PE_SIZEOF_PARAMETER_FIELD_INFO = 7 * PE_WORD_LEN + PE_PF_NAME_LEN + PE_PF_PROMPT_LEN + 2 * PE_PF_VALUE_LEN + PE_PF_REPORT_NAME_LEN + PE_PF_EDITMASK_LEN + 2 * PE_DOUBLE_LEN

Declare Function PEGetNParameterFields Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthParameterField Lib "crpe32.dll" (ByVal printJob%, ByVal varN%, varInfo As PEParameterFieldInfo) As Integer

Declare Function PESetNthParameterField Lib "crpe32.dll" (ByVal printJob%, ByVal varN%, varInfo As PEParameterFieldInfo) As Integer

'*** Converting parameterInfo default value or current value into value info ****
Global Const PE_VI_STRING_LEN = 256

' define value type
Global Const PE_VI_NUMBER = 0
Global Const PE_VI_CURRENCY = 1
Global Const PE_VI_BOOLEAN = 2
Global Const PE_VI_DATE = 3
Global Const PE_VI_STRING = 4
Global Const PE_VI_DATETIME = 5
Global Const PE_VI_TIME = 6
Global Const PE_VI_INTEGER = 7
Global Const PE_VI_COLOR = 8
Global Const PE_VI_CHAR = 9
Global Const PE_VI_LONG = 10
Global Const PE_VI_NOVALUE = 100

Type PEValueInfo
    StructSize As Integer
    valueType As Integer  'a PE_VI_ constant
    viNumber As Double
    viCurrency As Double
    viBoolean As Long
    viString As String * PE_VI_STRING_LEN
    viDate(0 To 2) As Integer ' year, month, day
    viDateTime(0 To 5) As Integer ' year, month, day, hour, minute, second
    viTime(0 To 2) As Integer  ' hour, minute, second
    viColor As Long
    viInteger As Integer
    viC As Byte
    ignored As Byte 'for 4 byte alignment. ignored.
    viLong As Long
End Type

Global Const PE_SIZEOF_VALUE_INFO = 2 * PE_BYTE_LEN + _
                                   15 * PE_WORD_LEN + _
                                    3 * PE_LONG_LEN + _
                                    2 * PE_DOUBLE_LEN + _
                                    1 * PE_VI_STRING_LEN
                                     
Declare Function PEConvertPFInfoToVInfo Lib "crpe32.dll" (ByVal Value As Any, ByVal valueType%, valueInfo As PEValueInfo) As Integer
Declare Function PEConvertVInfoToPFInfo Lib "crpe32.dll" (valueInfo As PEValueInfo, valueType%, ByVal Value As Any) As Integer

' Default values for Parameter fields.
' - If return value is -1 then an error has occurred.
' ------------------------------------
Declare Function PEGetNParameterDefaultValues Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$) As Integer

Declare Function PEGetNthParameterDefaultValue Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, ByVal Index%, valueInfo As PEValueInfo) As Integer

Declare Function PESetNthParameterDefaultValue Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, ByVal Index%, valueInfo As PEValueInfo) As Integer

Declare Function PEAddParameterDefaultValue Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, valueInfo As PEValueInfo) As Integer

Declare Function PEDeleteNthParameterDefaultValue Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, ByVal Index%) As Integer

' Min/Max values for Parameter fields.
' ------------------------------------
Declare Function PEGetParameterMinMaxValue Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, valueMin As PEValueInfo, valueMax As PEValueInfo) As Integer

Declare Function PESetParameterMinMaxValue Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, valueMin As PEValueInfo, valueMax As PEValueInfo) As Integer

' Pick list options in Parameter fields.
Declare Function PEGetNthParameterValueDescription Lib "crpe32.dll" (ByVal printJob As Integer, ByVal parameterFieldName As String, ByVal reportName As String, ByVal Index As Integer, valueDesc As Long, valueDescLength As Integer) As Integer
Declare Function PESetNthParameterValueDescription Lib "crpe32.dll" (ByVal printJob As Integer, ByVal parameterFieldName As String, ByVal reportName As String, ByVal Index As Integer, ByVal valueDesc As String) As Integer

' constants for sortMethod in PEParameterPickListOption struct
Global Const PE_OR_NO_SORT = 0
Global Const PE_OR_ALPHANUMERIC_ASCENDING = 1
Global Const PE_OR_ALPHANUMERIC_DESCENDING = 2
Global Const PE_OR_NUMERIC_ASCENDING = 3
Global Const PE_OR_NUMERIC_DESCENDING = 4

Type PEParameterPickListOption
    StructSize As Integer       ' initialize to PE_SIZEOF_PICK_LIST_OPTION
    showDescOnly As Integer     ' boolean value or PE_UNCHANGED
    sortMethod As Integer       ' enum type const, PE_UNCHANGED for no change
    sortBasedOnDesc As Integer  ' boolean value or PE_UNCHANGED
End Type

Global Const PE_SIZEOF_PICK_LIST_OPTION = (4 * PE_WORD_LEN)

Declare Function PEGetParameterPickListOption Lib "crpe32.dll" (ByVal printJob As Integer, ByVal parameterFieldName As String, ByVal reportName As String, pickListOption As PEParameterPickListOption) As Integer
Declare Function PESetParameterPickListOption Lib "crpe32.dll" (ByVal printJob As Integer, ByVal parameterFieldName As String, ByVal reportName As String, pickListOption As PEParameterPickListOption) As Integer

' Parameter current values
' ------------------------

'parameter field origin
Global Const PE_PO_REPORT = 0
Global Const PE_PO_STOREDPROC = 1
Global Const PE_PO_QUERY = 2

' range info
Global Const PE_RI_INCLUDEUPPERBOUND = 1
Global Const PE_RI_INCLUDELOWERBOUND = 2
Global Const PE_RI_NOUPPERBOUND = 4
Global Const PE_RI_NOLOWERBOUND = 8

Global Const PE_DR_HASRANGE = 0
Global Const PE_DR_HASDISCRETE = 1
Global Const PE_DR_HASDISCRETEANDRANGE = 2

Type PEParameterValueInfo
    StructSize As Integer
    isNullable As Integer               ' Boolean value or PE_UNCHANGED for no change.
    disallowEditing As Integer          ' Boolean value or PE_UNCHANGED for no change.
    allowMultipleValues As Integer      ' Boolean value or PE_UNCHANGED for no change.
    hasDiscreteValues As Integer        ' PE_UNCHANGED for no change.
                                        ' 0 means has ranges, 1 means has discrete values
                                        ' 2 means has discrete and ranged values
    partOfGroup As Integer              ' Boolean value or PE_UNCHANGED for no change.
    groupNum As Integer                 ' a group number or PE_UNCHANGED for no change.
    mutuallyExclusiveGroup As Integer   ' Boolean value or PE_UNCHANGED for no change.
End Type

Global Const PE_SIZEOF_PARAMETER_VALUE_INFO = 8 * PE_WORD_LEN

Declare Function PEGetParameterValueInfo Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, valueInfo As PEParameterValueInfo) As Integer

Declare Function PESetParameterValueInfo Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, valueInfo As PEParameterValueInfo) As Integer

' If return value is -1 then an error has occurred.
Declare Function PEGetNParameterCurrentValues Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$) As Integer

Declare Function PEGetNthParameterCurrentValue Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, ByVal Index%, currentValue As PEValueInfo) As Integer

Declare Function PEAddParameterCurrentValue Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, currentValue As PEValueInfo) As Integer

' If return value is -1 then an error has occurred.
Declare Function PEGetNParameterCurrentRanges Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$) As Integer

Declare Function PEGetNthParameterCurrentRange Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, ByVal Index%, rangeStart As PEValueInfo, rangeEnd As PEValueInfo, ByVal rangeInfo%) As Integer

Declare Function PEAddParameterCurrentRange Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$, rangeStart As PEValueInfo, rangeEnd As PEValueInfo, ByVal rangeInfo%) As Integer

Declare Function PEGetNthParameterType Lib "crpe32.dll" (ByVal printJob%, ByVal Index%) As Integer 'returns PE_PO_* or -1 if index is invalid.

Declare Function PEClearParameterCurrentValuesAndRanges Lib "crpe32.dll" (ByVal printJob%, ByVal parameterFieldName$, ByVal reportName$) As Integer

' Subreport Functions
' -------------------------------------------

Global Const PE_SUBREPORT_NAME_LEN = 128

' PE_SRI_ONOPENJOB: Reimport the subreport when opening the main report.
' PE_SRI_ONFUNCTIONCALL: Reimport the subreport when the api is called.
Global Const PE_SRI_UNDEFINED = -1
Global Const PE_SRI_ONOPENJOB = 0
Global Const PE_SRI_ONFUNCTIONCALL = 1

Type PESubreportInfo
    StructSize As Integer            ' Initialize to PE_SIZEOF_SUBREPORT_INFO.
    
    'Strings are null-terminated.
    Name As String * PE_SUBREPORT_NAME_LEN

    'number of links
    NLinks As Integer

    ' subreport placement.
    isOnDemand As Integer       ' TRUE if the subreport is on demand subreport.
    external As Integer         ' 1: the subreport is imported; 0: otherwise.
    reimportOption As Integer   ' PE_SRI_ONOPENJOB or PE_SRI_ONFUNCTIONCALL
End Type

Global Const PE_SIZEOF_SUBREPORT_INFO = 4 * PE_WORD_LEN + PE_SUBREPORT_NAME_LEN + 1

Declare Function PEGetSubreportInfo Lib "crpe32.dll" (ByVal printJob%, ByVal subreportHandle As Long, subreportInfo As PESubreportInfo) As Integer

Declare Function PEReimportSubreport Lib "crpe32.dll" (ByVal printJob As Integer, ByVal subreportHandle As Long, linkChanged As Long, reimported As Long) As Integer

' End Subreports


' Controlling sort order and group sort order
' -------------------------------------------

Global Const PE_SF_MAXNAMELEN = 50  ' maximum length of a sort field name

Global Const PE_SF_DESC = 0         ' values for the Direction parameter
Global Const PE_SF_ASC = 1

Declare Function PEGetNSortFields Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortNumber%, nameHandle As Long, nameLength%, Direction%) As Integer

Declare Function PESetNthSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortNumber%, ByVal SortFieldName$, ByVal Direction%) As Integer

Declare Function PEDeleteNthSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortFieldN%) As Integer

Declare Function PEGetNGroupSortFields Lib "crpe32.dll" (ByVal printJob%) As Integer

Declare Function PEGetNthGroupSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortFieldN%, nameHandle As Long, nameLength%, Direction%) As Integer

Declare Function PESetNthGroupSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortFieldN%, ByVal SortGroupName$, ByVal Direction%) As Integer

Declare Function PEDeleteNthGroupSortField Lib "crpe32.dll" (ByVal printJob%, ByVal SortFieldN%) As Integer


' Controlling databases
' ---------------------
'
' The following functions allow retrieving and updating database info
' in an opened report, so that a report can be printed using different
' session, server, database, user and/or table location settings.  Any
' changes made to the report via these functions are not permanent, and
' only last as long as the report is open.
'
' The following database functions (except for PELogOnServer and
' PELogOffServer) must be called after PEOpenPrintJob and before
' PEStartPrintJob.

' The function PEGetNTables is called to fetch the number of tables in
' the report.  This includes all PC databases (e.g. Paradox, xBase)
' as well as SQL databases (e.g. SQL Server, Oracle, Netware).

Declare Function PEGetNTables Lib "crpe32.dll" (ByVal printJob%) As Integer

' The function PEGetNthTableType allows the application to determine the
' type of each table.  The application can test DBType (equal to
' PE_DT_STANDARD or PE_DT_SQL), or test the database DLL name used to
' create the report.  DLL names have the following naming convention:
'     - PDB*.DLL for standard (non-SQL) databases,
'     - PDS*.DLL for SQL databases.
'
' In the case of ODBC (PDSODBC.DLL) the DescriptiveName includes the
' ODBC data source name.

Global Const PE_DLL_NAME_LEN = 64
Global Const PE_FULL_NAME_LEN = 256
Global Const PE_SIZEOF_TABLE_TYPE = 324 ' # bytes in PETableType

Global Const PE_DT_STANDARD = 1  ' values for DBType
Global Const PE_DT_SQL = 2
Global Const PE_DT_SQL_STORED_PROCEDURE = 3

Type PETableType
    StructSize As Integer   ' initialize to # bytes in PETableType

    DLLName As String * PE_DLL_NAME_LEN
    DescriptiveName  As String * PE_FULL_NAME_LEN

    DBType As Integer
End Type

Declare Function PEGetNthTableType Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, TableType As PETableType) As Integer

' The functions PEGetNthTableSessionInfo and PESetNthTableSessionInfo
' are only used when connecting to MS Access databases (which require a
' session to be opened first)

Global Const PE_SESS_USERID_LEN = 128
Global Const PE_SESS_PASSWORD_LEN = 128
Global Const PE_SIZEOF_SESSION_INFO = 262  ' # bytes in PESessionInfo

Type PESessionInfo
    'initialize to # bytes in PESessionInfo
    StructSize As Integer

    ' Password is undefined when getting information from report.
    UserID As String * PE_SESS_USERID_LEN
    Password As String * PE_SESS_PASSWORD_LEN

    ' SessionHandle is undefined when getting information from report.
    ' When setting information, if it is = 0 the UserID and Password
    ' settings are used, otherwise the SessionHandle is used.
    SessionHandle As Long
End Type

Declare Function PEGetNthTableSessionInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, SessionInfo As PESessionInfo) As Integer

Declare Function PESetNthTableSessionInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, SessionInfo As PESessionInfo, ByVal PropagateAcrossTables%) As Integer

' Logging on is performed when printing the report, but the correct
' log on information must first be set using PESetNthTableLogOnInfo.
' Only the password is required, but the server, database, and
' user names may optionally be overriden as well.
'
' If the parameter propagateAcrossTables is TRUE, the new log on info
' is also applied to any other tables in this report that had the
' same original server and database names as this table.  If FALSE
' only this table is updated.
'
' Logging off is performed automatically when the print job is closed.

Global Const PE_SERVERNAME_LEN = 128
Global Const PE_DATABASENAME_LEN = 128
Global Const PE_USERID_LEN = 128
Global Const PE_PASSWORD_LEN = 128
Global Const PE_SIZEOF_LOGON_INFO = 514  ' # bytes in PELogOnInfo

Type PELogOnInfo
    ' initialize to # bytes in PELogOnInfo
    StructSize As Integer

    ' For any of the following values an empty string ("") means to use
    ' the value already set in the report.  To override a value in the
    ' report use a non-empty string (e.g. "Server A").
    '
    ' For Netware SQL, pass the dictionary path name in ServerName and
    ' data path name in DatabaseName.

    ServerName As String * PE_SERVERNAME_LEN
    DatabaseName  As String * PE_DATABASENAME_LEN
    UserID As String * PE_USERID_LEN

    ' Password is undefined when getting information from report.

    Password  As String * PE_PASSWORD_LEN
End Type

Declare Function PEGetNthTableLogOnInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, LogOnInfo As PELogOnInfo) As Integer
Declare Function PESetNthTableLogOnInfo Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, LogOnInfo As PELogOnInfo, ByVal Propagate%) As Integer

' A table's location is fetched and set using PEGetNthTableLocation and
' PESetNthTableLocation.  This name is database-dependent, and must be
' formatted correctly for the expected database.  For example:
'     - Paradox: "c:\crw\ORDERS.DB"
'     - SQL Server: "publications.dbo.authors"

Global Const PE_TABLE_LOCATION_LEN = 256
Global Const PE_CONNECTION_BUFFER_LEN = 512

Type PETableLocation
    ' initialize to # bytes in PETableLocation
    StructSize As Integer
    Location  As String * PE_TABLE_LOCATION_LEN
    SubLocation As String * PE_TABLE_LOCATION_LEN
    ConnectBuffer As String * PE_CONNECTION_BUFFER_LEN  ' Connection Info for attached tables
End Type

Global Const PE_SIZEOF_TABLE_LOCATION = 1026  ' # bytes in PETableLocation

Declare Function PEGetNthTableLocation Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, Location As PETableLocation) As Integer
Declare Function PESetNthTableLocation Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%, Location As PETableLocation) As Integer

' The function PETestNthTableConnectivity tests whether a database
' table's settings are valid and ready to be reported on.  It returns
' true if the database session, log on, and location info is all
' correct.
'
' This is useful, for example, in prompting the user and testing a
' server password before printing begins.
'
' This function may require a significant amount of time to complete,
' since it will first open a user session (if required), then log onto
' the database server (if required), and then open the appropriate
' database table (to test that it exists).  It does not read any data,
' and closes the table immediately once successful.  Logging off is
' performed when the print job is closed.
'
' If it fails in any of these steps, the error code set indicates
' which database info needs to be updated using functions above:
'    - If it is unable to begin a session, PE_ERR_DATABASESESSION is set,
'      and the application should update with PESetNthTableSessionInfo.
'    - If it is unable to log onto a server, PE_ERR_DATABASELOGON is set,
'      and the application should update with PESetNthTableLogOnInfo.
'    - If it is unable open the table, PE_ERR_DATABASELOCATION is set,
'      and the application should update with PESetNthTableLocation.

Declare Function PETestNthTableConnectivity Lib "crpe32.dll" (ByVal printJob%, ByVal TableN%) As Integer

' PELogOnServer and PELogOffServer can be called at any time to log on
' and off of a database server.  These functions are not required if
' function PESetNthTableLogOnInfo above was already used to set the
' password for a table.
'
' These functions require a database DLL name, which can be retrieved
' using PEGetNthTableType above.
'
' This function can also be used for non-SQL tables, such as password-
' protected Paradox tables.  Call this function to set the password
' for the Paradox DLL before beginning printing.
'
' Note: When printing using PEStartPrintJob the ServerName passed in
' PELogOnServer must agree exactly with the server name stored in the
' report.  If this is not true use PESetNthTableLogOnInfo to perform
' logging on instead.

Declare Function PELogOnServer Lib "crpe32.dll" (ByVal DLLName$, LogOnInfo As PELogOnInfo) As Integer
Declare Function PELogOffServer Lib "crpe32.dll" (ByVal DLLName$, LogOnInfo As PELogOnInfo) As Integer
Declare Function PELogOnSQLServerWithPrivateInfo Lib "crpe32.dll" (ByVal DLLName$, ByVal PrivateInfo As Long) As Integer


' Overriding SQL query in report
' ------------------------------
'
' PEGetSQLQuery returns the same query as appears in the Show SQL Query
' dialog in CRW, in syntax specific to the database driver you are using.
'
' PESetSQLQuery is mostly useful for reports with SQL queries that
' were explicitly edited in the Show SQL Query dialog in CRW, i.e. those
' reports that needed database-specific selection criteria or joins.
' (Otherwise it is usually best to continue using function calls such as
' PESetSelectionFormula and let CRW build the SQL query automatically.)
'
' PESetSQLQuery has the same restrictions as editing in the Show SQL
' Query dialog; in particular that changes are accepted in the FROM and
' WHERE clauses but ignored in the SELECT list of fields.

Declare Function PEGetSQLQuery Lib "crpe32.dll" (ByVal printJob%, TextHandle As Long, TextLength%) As Integer

Declare Function PESetSQLQuery Lib "crpe32.dll" (ByVal printJob%, ByVal QueryString$) As Integer

'Field mapping types
Global Const PE_FM_AUTO_FLD_MAP = 0          'Automatic field name mapping
                                             'unmapped report fields will be removed
Global Const PE_FM_CRPE_PROMPT_FLD_MAP = 1   'CRPE provided dialog box to map field manually

Declare Function PESetFieldMappingType Lib "crpe32.dll" (ByVal printJob%, ByVal mappingtype As Integer) As Integer

Declare Function PEGetFieldMappingType Lib "crpe32.dll" (ByVal printJob%, mappingtype As Integer) As Integer

Declare Function PEVerifyDatabase Lib "crpe32.dll" (ByVal printJob%) As Integer

' constants returned from PECheckNthTableDifferences, can be any
' combination of the following:
Global Const PE_TCD_OKAY = &H0
Global Const PE_TCD_DATABASENOTFOUND = &H1
Global Const PE_TCD_SERVERNOTFOUND = &H2
Global Const PE_TCD_SERVERNOTOPENED = &H4
Global Const PE_TCD_ALIASCHANGED = &H8
Global Const PE_TCD_INDEXESCHANGED = &H10
Global Const PE_TCD_DRIVERCHANGED = &H20
Global Const PE_TCD_DICTIONARYCHANGED = &H40
Global Const PE_TCD_FILETYPECHANGED = &H80
Global Const PE_TCD_RECORDSIZECHANGED = &H100
Global Const PE_TCD_ACCESSCHANGED = &H200
Global Const PE_TCD_PARAMETERSCHANGED = &H400
Global Const PE_TCD_LOCATIONCHANGED = &H800
Global Const PE_TCD_DATABASEOTHER = &H1000
Global Const PE_TCD_NUMFIELDSCHANGED = &H10000
Global Const PE_TCD_FIELDOTHER = &H20000
Global Const PE_TCD_FIELDNAMECHANGED = &H40000
Global Const PE_TCD_FIELDDESCCHANGED = &H80000
Global Const PE_TCD_FIELDTYPECHANGED = &H100000
Global Const PE_TCD_FIELDSIZECHANGED = &H200000
Global Const PE_TCD_NATIVEFIELDTYPECHANGED = &H400000
Global Const PE_TCD_NATIVEFIELDOFFSETCHANGED = &H800000
Global Const PE_TCD_NATIVEFIELDSIZECHANGED = &H1000000
Global Const PE_TCD_FIELDDECPLACESCHANGED = &H2000000

Type PETableDifferenceInfo
    StructSize As Integer
    tableDifferences As Long    ' any combination of PE_TC_*
    reserved1 As Long           ' reserved - do not use
    reserved2 As Long           ' reserved - do not use
End Type

Global Const PE_SIZEOF_TABLE_DIFFERENCE_INFO = (1 * PE_WORD_LEN) + (3 * PE_LONG_LEN)

' Not implemented for reports based on dictionary, returns PE_ERR_NOTIMPLEMENTED
Declare Function PECheckNthTableDifferences Lib "crpe32.dll" (ByVal printJob As Integer, ByVal TableN As Integer, tabledifferenceinfo As PETableDifferenceInfo) As Integer

' Saved data
' ----------
'
' Use PEHasSavedData to find out if a report currently has saved data
' associated with it.  This may or may not be TRUE when a print job is
' first opened from a report file.  Since data is saved during a print,
' this will always be TRUE immediately after a report is printed.
'
' Use PEDiscardSavedData to release the saved data associated with a
' report.  The next time the report is printed, it will get current data
' from the database.
'
' The default behavior is for a report to use its saved data, rather than
' refresh its data from the database when printing a report.

Declare Function PEHasSavedData Lib "crpe32.dll" (ByVal printJob%, HasSavedData As Long) As Integer

Declare Function PEDiscardSavedData Lib "crpe32.dll" (ByVal printJob%) As Integer


' Report title
' ------------

Declare Function PEGetReportTitle Lib "crpe32.dll" (ByVal printJob%, TitleHandle As Long, titleLength%) As Integer
Declare Function PESetReportTitle Lib "crpe32.dll" (ByVal printJob%, ByVal Title$) As Integer


' Controlling print to window
' ---------------------------

Declare Function PEOutputToWindow Lib "crpe32.dll" (ByVal printJob%, ByVal Title$, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal height As Long, ByVal style As Long, ByVal PWindow As Long) As Integer

Type PEWindowOptions
    StructSize As Integer            'initialize to PE_SIZEOF_WINDOW_OPTIONS

    hasGroupTree As Integer          '0 or 1, except use PE_UNCHANGED for no change
    canDrillDown As Integer          '0 or 1, except use PE_UNCHANGED for no change
    hasNavigationControls As Integer '0 or 1, except use PE_UNCHANGED for no change
    hasCancelButton As Integer       '0 or 1, except use PE_UNCHANGED for no change
    hasPrintButton As Integer        '0 or 1, except use PE_UNCHANGED for no change
    hasExportButton As Integer       '0 or 1, except use PE_UNCHANGED for no change
    hasZoomControl As Integer        '0 or 1, except use PE_UNCHANGED for no change
    hasCloseButton As Integer        '0 or 1, except use PE_UNCHANGED for no change
    hasProgressControls As Integer   '0 or 1, except use PE_UNCHANGED for no change
    hasSearchButton As Integer       '0 or 1, except use PE_UNCHANGED for no change
    hasPrintSetupButton As Integer   '0 or 1, except use PE_UNCHANGED for no change
    hasRefreshButton As Integer      '0 or 1, except use PE_UNCHANGED for no change
    showToolbarTips As Integer       ' BOOL value, except use PE_UNCHANGED for no change
                                     ' default is TRUE (*Show* tooltips on toolbar)
    showDocumentTips As Integer      ' BOOL value, except use PE_UNCHANGED for no change
                                     ' default is FALSE (*Hide* tooltips on document)
    hasLaunchButton As Integer       ' Launch Seagate Analysis button on toolbar.
                                     ' BOOL value, except use PE_UNCHANGED for no change
                                     ' default is FALSE
End Type

Global Const PE_SIZEOF_WINDOW_OPTIONS = (16 * PE_WORD_LEN)

Declare Function PEGetWindowOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEWindowOptions) As Integer

Declare Function PESetWindowOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEWindowOptions) As Integer

Declare Function PEGetWindowHandle Lib "crpe32.dll" (ByVal printJob%) As Long

Declare Sub PECloseWindow Lib "crpe32.dll" (ByVal printJob%)


' Controlling printed pages
' -------------------------

Declare Function PEShowNextPage Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEShowFirstPage Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEShowPreviousPage Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEShowLastPage Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEGetNPages Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEShowNthPage Lib "crpe32.dll" (ByVal printJob%, ByVal pageN%) As Integer

Global Const PE_ZOOM_FULL_SIZE = 0
Global Const PE_ZOOM_SIZE_FIT_ONE_SIDE = 1
Global Const PE_ZOOM_SIZE_FIT_BOTH_SIDES = 2

Declare Function PEZoomPreviewWindow Lib "crpe32.dll" (ByVal printJob%, ByVal ZoomLevel%) As Integer
' ZoomLevel is a percent from 25 to 400 or a PE_ZOOM_ constant

' Controlling print window when print control buttons hidden
' ----------------------------------------------------------

Declare Function PEShowPrintControls Lib "crpe32.dll" (ByVal printJob%, ByVal ShowPrintControls%) As Integer

Declare Function PEPrintControlsShowing Lib "crpe32.dll" (ByVal printJob%, ControlsShowing As Long) As Integer

Declare Function PEPrintWindow Lib "crpe32.dll" (ByVal printJob%, ByVal WaitNoWait%) As Integer

Declare Function PEExportPrintWindow Lib "crpe32.dll" (ByVal printJob%, ByVal ToMail%, ByVal WaitUntilDone%) As Integer

Declare Function PENextPrintWindowMagnification Lib "crpe32.dll" (ByVal printJob%) As Integer


' Changing printer selection
' --------------------------

Declare Function PESelectPrinter Lib "crpe32.dll" (ByVal printJob%, ByVal PrinterDriver$, ByVal PrinterName$, ByVal PortName$, DevMode As Any) As Integer

Declare Function PEGetSelectedPrinter Lib "crpe32.dll" (ByVal printJob%, DriverHandle As Long, DriverLength%, PrinterHandle As Long, PrinterLength%, PortHandle As Long, PortLength%, DevMode As Any) As Integer

Declare Function PEFreeDevMode Lib "crpe32.dll" (ByVal printJob As Integer, mode As Any) As Integer



' Controlling print to printer
' ----------------------------

Declare Function PEOutputToPrinter Lib "crpe32.dll" (ByVal printJob%, ByVal nCopies%) As Integer

Declare Function PESetNDetailCopies Lib "crpe32.dll" (ByVal printJob%, ByVal NDetailCopies%) As Integer

Declare Function PEGetNDetailCopies Lib "crpe32.dll" (ByVal printJob%, NDetailCopies%) As Integer

' Extension to PESetPrintOptions function: If the 2nd parameter
' (pointer to PEPrintOptions) is set to 0 (null) the function prompts
' the user for these options.
'
' With this change, you can get the behaviour of the print-to-printer
' button in the print window by calling PESetPrintOptions with a
' null pointer and then calling PEPrintWindow.

Global Const PE_MAXPAGEN = 65535
Global Const PE_FILE_PATH_LEN = 512

Global Const PE_UNCOLLATED = 0
Global Const PE_COLLATED = 1
Global Const PE_DEFAULTCOLLATION = 2

Type PEPrintOptions
    StructSize As Integer   ' initialize to # bytes in PEPrintOptions

    ' page and copy numbers are 1-origin
    ' use 0 to preserve the existing settings
    StartPageN As Integer
    stopPageN As Integer

    nReportCopies As Integer
    collation As Integer
    outputFileName As String * PE_FILE_PATH_LEN
End Type

Global Const PE_SIZEOF_PRINT_OPTIONS = 5 * PE_WORD_LEN + PE_FILE_PATH_LEN

Declare Function PESetPrintOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEPrintOptions) As Integer

Declare Function PEGetPrintOptions Lib "crpe32.dll" (ByVal printJob%, Options As PEPrintOptions) As Integer


' Controlling print to file and export
' ------------------------------------

Global Const PE_FT_RECORD = 0
Global Const PE_FT_TABSEPARATED = 1
Global Const PE_FT_TEXT = 2
Global Const PE_FT_DIF = 3
Global Const PE_FT_CSV = 4
Global Const PE_FT_CHARSEPARATED = 5
Global Const PE_FT_TABFORMATTED = 6

Global Const PE_FIELDDELIMLEN = 17

Declare Function PEOutputToFile Lib "crpe32.dll" (ByVal printJob%, ByVal OutputFilePath$, ByVal FileType%, Options As Any) As Integer

Type PEExportOptions
    StructSize As Integer   'initialize to # bytes in PEExportOptions

    FormatDLLName As String * PE_DLL_NAME_LEN
    FormatType1 As Integer
    FormatType2 As Integer
    FormatOptions1 As Integer
    FormatOptions2 As Integer

    DestinationDLLName As String * PE_DLL_NAME_LEN
    DestinationType1 As Integer
    DestinationType2 As Integer
    DestinationOptions1 As Integer
    DestinationOptions2 As Integer

    ' following are set by PEGetExportOptions,
    ' and ignored by PEExportTo.
    NFormatOptionsBytes As Integer
    NDestinationOptionsBytes As Integer
End Type

Global Const PE_SIZEOF_EXPORT_OPTIONS = 11 * PE_WORD_LEN + 2 * PE_DLL_NAME_LEN

Declare Function PEGetExportOptions Lib "crpe32.dll" (ByVal printJob%, ExportOptions As PEExportOptions) As Integer

Declare Function PEExportTo Lib "crpe32.dll" (ByVal printJob%, ExportOptions As PEExportOptions) As Integer


' Setting page margins
' --------------------

Global Const PE_SM_DEFAULT = &H8000

Declare Function PESetMargins Lib "crpe32.dll" (ByVal printJob%, ByVal LeftMargin%, ByVal RightMargin%, ByVal TopMargin%, ByVal BottomMargin%) As Integer

Declare Function PEGetMargins Lib "crpe32.dll" (ByVal printJob%, LeftMargin%, RightMargin%, TopMargin%, BottomMargin%) As Integer


'Report Summary Info
'-------------------

Global Const PE_SI_APPLICATION_NAME_LEN = 128
Global Const PE_SI_TITLE_LEN = 128
Global Const PE_SI_SUBJECT_LEN = 128
Global Const PE_SI_AUTHOR_LEN = 128
Global Const PE_SI_KEYWORDS_LEN = 128
Global Const PE_SI_COMMENTS_LEN = 512
Global Const PE_SI_REPORT_TEMPLATE_LEN = 128

Type PEReportSummaryInfo
    StructSize As Integer
    applicationName As String * PE_SI_APPLICATION_NAME_LEN ' read only.
    Title As String * PE_SI_TITLE_LEN
    subject As String * PE_SI_SUBJECT_LEN
    author As String * PE_SI_AUTHOR_LEN
    keywords As String * PE_SI_KEYWORDS_LEN
    comments As String * PE_SI_COMMENTS_LEN
    reportTemplate As String * PE_SI_REPORT_TEMPLATE_LEN
    savePreviewPicture As Integer   ' BOOL PE_UNCHANGED for no change
End Type

Global Const PE_SIZEOF_REPORT_SUMMARY_INFO = 1 * PE_WORD_LEN + _
                                      PE_SI_APPLICATION_NAME_LEN + _
                                      PE_SI_TITLE_LEN + _
                                      PE_SI_SUBJECT_LEN + _
                                      PE_SI_AUTHOR_LEN + _
                                      PE_SI_KEYWORDS_LEN + _
                                      PE_SI_COMMENTS_LEN + _
                                      PE_SI_REPORT_TEMPLATE_LEN + _
                                      1 * PE_WORD_LEN
                                
Declare Function PEGetReportSummaryInfo Lib "crpe32.dll" (ByVal printJob%, summaryInfo As PEReportSummaryInfo) As Integer
Declare Function PESetReportSummaryInfo Lib "crpe32.dll" (ByVal printJob%, summaryInfo As PEReportSummaryInfo) As Integer

' Setting section height and format
' ---------------------------------

Declare Function PEGetNSections Lib "crpe32.dll" (ByVal printJob%) As Integer
Declare Function PEGetNSectionsInArea Lib "crpe32.dll" (ByVal printJob As Integer, ByVal areaCode As Integer) As Integer ' return -1 if failed

Declare Function PEGetSectionCode Lib "crpe32.dll" (ByVal printJob%, ByVal sectionN%) As Integer

' MinimumHeight is in twips - 1440 twips to the inch
'This is the replacement API Call for PEGetMinimumSectionHeight which is obsolete.
'The obsolete API will still work for older applications, but use this for all new development
Declare Function PEGetSectionHeight Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, height%) As Integer

'This is the replacement API Call for PESetMinimumSectionHeight which is obsolete.
'The obsolete API will still work for older applications, but use this for all new development
Declare Function PESetSectionHeight Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal height%) As Integer

Type PESectionOptions
    StructSize As Integer   ' initialize to PE_SIZEOF_SECTION_OPTIONS

    ' use 0 to turn off, 1 to turn on and PE_UNCHANGED to preserve each attribute
    Visible As Integer
    NewPageBefore As Integer
    NewPageAfter As Integer
    KeepTogether As Integer
    SuppressBlankSection As Integer
    ResetPageNAfter As Integer
    PrintAtBottomOfPage As Integer
    backgroundColor As Long   ' Use PE_UNCHANGED_COLOR to preserve the
                              ' existing color.
    underlaySection As Integer
    showArea As Integer
    freeFormPlacement As Integer
    reserveMinimumPageFooter As Integer 'BOOLEAN or PE_UNCHANGED; Sets the size of the Page Footer area                        // to the size of the largest Page Footer section. Used with
                     ' PEGetAreaFormat/PESetAreaFormat; ignored when used with PEGetSectionFormat/PESetSectionFormat.
End Type

Global Const PE_SIZEOF_SECTION_OPTIONS = 12 * PE_WORD_LEN + 1 * 4

'Format formula name
'Old naming convention
Global Const SECTION_VISIBILITY = 58
Global Const NEW_PAGE_BEFORE = 60
Global Const NEW_PAGE_AFTER = 61
Global Const KEEP_SECTION_TOGETHER = 62
Global Const SUPPRESS_BLANK_SECTION = 63
Global Const RESET_PAGE_N_AFTER = 64
Global Const PRINT_AT_BOTTOM_OF_PAGE = 65
Global Const UNDERLAY_SECTION = 66
Global Const SECTION_BACK_COLOUR = 67

'New naming convention
Global Const PE_FFN_AREASECTION_VISIBILITY = 58
Global Const PE_FFN_SECTION_VISIBILITY = 58
Global Const PE_FFN_SHOW_AREA = 59
Global Const PE_FFN_NEW_PAGE_BEFORE = 60
Global Const PE_FFN_NEW_PAGE_AFTER = 61
Global Const PE_FFN_KEEP_SECTION_TOGETHER = 62
Global Const PE_FFN_KEEP_TOGETHER = 62
Global Const PE_FFN_SUPPRESS_BLANK_SECTION = 63
Global Const PE_FFN_RESET_PAGE_N_AFTER = 64
Global Const PE_FFN_PRINT_AT_BOTTOM_OF_PAGE = 65
Global Const PE_FFN_UNDERLAY_SECTION = 66
Global Const PE_FFN_SECTION_BACK_COLOUR = 67
Global Const PE_FFN_SECTION_BACK_COLOR = 67

Declare Function PEGetSectionFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal formulaName%, TextHandle As Long, TextLength%) As Integer

Declare Function PESetSectionFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal formulaName%, ByVal formulaString$) As Integer

Declare Function PEGetSectionFormat Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, Options As PESectionOptions) As Integer

Declare Function PESetSectionFormat Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, Options As PESectionOptions) As Integer

' Setting area format
' -------------------
                                                                                                               
Declare Function PEGetAreaFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal areaCode%, ByVal formulaName%, TextHandle As Long, TextLength%) As Integer

Declare Function PESetAreaFormatFormula Lib "crpe32.dll" (ByVal printJob%, ByVal areaCode%, ByVal formulaName%, ByVal formulaString$) As Integer

Declare Function PEGetAreaFormat Lib "crpe32.dll" (ByVal printJob%, ByVal areaCode%, Options As PESectionOptions) As Integer

Declare Function PESetAreaFormat Lib "crpe32.dll" (ByVal printJob%, ByVal areaCode%, Options As PESectionOptions) As Integer


' Setting font info
' -----------------

' values for ScopeCode - may be ORed together
Global Const PE_FIELDS = 1
Global Const PE_TEXT = 2

Global Const PE_UNCHANGED = -1

' to preserve the existing setting, use the following
'   for FontFamily%    use  FF_DONTCARE
'   for FontPitch%     use  DEFAULT_PITCH
'   for CharSet%       use  DEFAULT_CHARSET
'   for PointSize%     use  0
'   for isItalic%      use  PE_UNCHANGED
'   for isUnderlined%  use  PE_UNCHANGED
'   for isStruckOut%   use  PE_UNCHANGED
'   for Weight%        use  0
Declare Function PESetFont Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal scopecode%, ByVal faceName$, ByVal fontFamily%, ByVal fontPitch%, ByVal charSet%, ByVal pointSize%, ByVal isItalic%, ByVal isUnderlined%, ByVal isStruckOut%, ByVal weight%) As Integer

' Subreports
' ----------
Declare Function PEGetNSubreportsInSection Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%) As Integer

Declare Function PEGetNthSubreportInSection Lib "crpe32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal subreportN%) As Long


'
' G R A P H I N G
'
' - graph type

Global Const PE_GT_BARCHART = 0
Global Const PE_GT_LINECHART = 1
Global Const PE_GT_AREACHART = 2
Global Const PE_GT_PIECHART = 3
Global Const PE_GT_DOUGHNUTCHART = 4
Global Const PE_GT_THREEDRISERCHART = 5
Global Const PE_GT_THREEDSURFACECHART = 6
Global Const PE_GT_SCATTERCHART = 7
Global Const PE_GT_RADARCHART = 8
Global Const PE_GT_BUBBLECHART = 9
Global Const PE_GT_STOCKCHART = 10
Global Const PE_GT_USERDEFINEDCHART = 50        ' for PEGetGraphTypeInfo only -
Global Const PE_GT_UNKNOWNTYPECHART = 100       ' do not use in PESetGraphTypeInfo.

' graph subtype
' - bar charts
Global Const PE_GST_SIDEBYSIDEBARCHART = 0
Global Const PE_GST_STACKEDBARCHART = 1
Global Const PE_GST_PERCENTBARCHART = 2
Global Const PE_GST_FAKED3DSIDEBYSIDEBARCHART = 3
Global Const PE_GST_FAKED3DSTACKEDBARCHART = 4
Global Const PE_GST_FAKED3DPERCENTBARCHART = 5
                                                    
' - line charts
Global Const PE_GST_REGULARLINECHART = 10
Global Const PE_GST_STACKEDLINECHART = 11
Global Const PE_GST_PERCENTAGELINECHART = 12
Global Const PE_GST_LINECHARTWITHMARKERS = 13
Global Const PE_GST_STACKEDLINECHARTWITHMARKERS = 14
Global Const PE_GST_PERCENTAGELINECHARTWITHMARKERS = 15
                                                    
' - charts
Global Const PE_GST_ABSOLUTEAREACHART = 20
Global Const PE_GST_STACKEDAREACHART = 21
Global Const PE_GST_PERCENTAREACHART = 22
Global Const PE_GST_FAKED3DABSOLUTEAREACHART = 23
Global Const PE_GST_FAKED3DSTACKEDAREACHART = 24
Global Const PE_GST_FAKED3DPERCENTAREACHART = 25
                                                    
' - pie charts
Global Const PE_GST_REGULARPIECHART = 30
Global Const PE_GST_FAKED3DREGULARPIECHART = 31
Global Const PE_GST_MULTIPLEPIECHART = 32
Global Const PE_GST_MULTIPLEPROPORTIONALPIECHART = 33
                                                    
' - mmm, doughnut charts
Global Const PE_GST_REGULARDOUGHNUTCHART = 40
Global Const PE_GST_MULTIPLEDOUGHNUTCHART = 41
Global Const PE_GST_MULTIPLEPROPORTIONALDOUGHNUTCHART = 42

' 3D riser charts
Global Const PE_GST_THREEDREGULARCHART = 50
Global Const PE_GST_THREEDPYRAMIDCHART = 51
Global Const PE_GST_THREEDOCTAGONCHART = 52
Global Const PE_GST_THREEDCUTCORNERSCHART = 53
                                                        
' 3D surface charts
Global Const PE_GST_THREEDSURFACEREGULARCHART = 60
Global Const PE_GST_THREEDSURFACEWITHSIDESCHART = 61
Global Const PE_GST_THREEDSURFACEHONEYCOMBCHART = 62
                                                    
' scatter charts
Global Const PE_GST_XYSCATTERCHART = 70
Global Const PE_GST_XYSCATTERDUALAXISCHART = 71
Global Const PE_GST_XYSCATTERWITHLABELSCHART = 72
Global Const PE_GST_XYSCATTERDUALAXISWITHLABELSCHART = 73
                                                    
' radar charts
Global Const PE_GST_REGULARRADARCHART = 80
Global Const PE_GST_STACKEDRADARCHART = 81
Global Const PE_GST_RADARDUALAXISCHART = 82
                                                    
' bubble charts
Global Const PE_GST_REGULARBUBBLECHART = 90
Global Const PE_GST_DUALAXISBUBBLECHART = 91
                                                    
' stocked charts
Global Const PE_GST_HIGHLOWCHART = 100
Global Const PE_GST_HIGHLOWDUALAXISCHART = 101
Global Const PE_GST_HIGHLOWOPENCHART = 102
Global Const PE_GST_HIGHLOWOPENDUALAXISCHART = 103
Global Const PE_GST_HIGHLOWOPENCLOSECHART = 104
Global Const PE_GST_HIGHLOWOPENCLOSEDUALAXISCHART = 105

Global Const PE_GST_UNKNOWNSUBTYPECHART = 1000

Type PEGraphTypeInfo
    StructSize As Integer
    GraphType As Integer        ' PE_GT_*, PE_UNCHANGED for no change
    graphSubtype As Integer     ' PE_GST_*, PE_UNCHANGED for no change
End Type

Global Const PE_SIZEOF_GRAPH_TYPE_INFO = (3 * PE_WORD_LEN)

Declare Function PEGetGraphTypeInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, graphTypeInfo As PEGraphTypeInfo) As Integer
Declare Function PESetGraphTypeInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, graphTypeInfo As PEGraphTypeInfo) As Integer

' graph text

Global Const PE_GTT_TITLE = 0
Global Const PE_GTT_SUBTITLE = 1
Global Const PE_GTT_FOOTNOTE = 2
Global Const PE_GTT_SERIESTITLE = 3
Global Const PE_GTT_GROUPSTITLE = 4
Global Const PE_GTT_XAXISTITLE = 5
Global Const PE_GTT_YAXISTITLE = 6
Global Const PE_GTT_ZAXISTITLE = 7

' graph text fonts
Global Const PE_GTF_TITLEFONT = 0
Global Const PE_GTF_SUBTITLEFONT = 1
Global Const PE_GTF_FOOTNOTEFONT = 2
Global Const PE_GTF_GROUPSTITLEFONT = 3
Global Const PE_GTF_DATATITLEFONT = 4
Global Const PE_GTF_LEGENDFONT = 5
Global Const PE_GTF_GROUPLABELSFONT = 6
Global Const PE_GTF_DATALABELSFONT = 7

Global Const PE_FACE_NAME_LEN = 64

' Graph text max length
Global Const PE_GRAPH_TEXT_LEN = 128

Type PEFontColorInfo
    StructSize As Integer
    faceName As String * PE_FACE_NAME_LEN       ' empty string for no change
    fontFamily As Integer       ' FF_DONTCARE for no change
    fontPitch As Integer        ' DEFAULT_PITCH for no change
    charSet As Integer          ' DEFAULT_CHARSET for no change
    pointSize As Integer        ' 0 for no change
    isItalic As Integer         ' BOOL value, except use PE_UNCHANGED for no change.
    isUnderlined As Integer     ' BOOL value, except use PE_UNCHANGED for no change.
    isStruckOut As Integer      ' BOOL value, except use PE_UNCHANGED for no change.
    weight As Integer           ' 0 for no change
    color As Long               ' COLORREF - PE_UNCHANGED_COLOR for no change.
    twipSize As Integer         ' Font size in twips, 0 for no change.
                                ' Use one of pointSize or twipSize. If both pointSize and twipSize
                                ' are non-zero, twipSize will be used and pointSize will be ignored.
End Type

Global Const PE_SIZEOF_FONT_COLOR_INFO = PE_WORD_LEN + PE_FACE_NAME_LEN + (9 * PE_WORD_LEN) + PE_LONG_LEN

' titleType - PE_GTT_*
' title - HANDLE
Declare Function PEGetGraphTextInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, ByVal titleType As Integer, Title As Long, titleLength As Integer) As Integer
Declare Function PESetGraphTextInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, ByVal titleType As Integer, ByVal Title As String) As Integer

' enable/disable graph default titles
'
' titleType - PE_GTT_*
'
Declare Function PESetGraphTextDefaultOption Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, ByVal titleType As Integer, ByVal useDefault As Long) As Integer
Declare Function PEGetGraphTextDefaultOption Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, ByVal titleType As Integer, useDefault As Long) As Integer

' graph font
' titleFontType - PE_GTF_*
Declare Function PEGetGraphFontInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, ByVal titleFontType As Integer, fontColourInfo As PEFontColorInfo) As Integer
Declare Function PESetGraphFontInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, ByVal titleFontType As Integer, fontColourInfo As PEFontColorInfo) As Integer

' graph options
Global Const PE_GLP_PLACEUPPERRIGHT = 0
Global Const PE_GLP_PLACEBOTTOMCENTER = 1
Global Const PE_GLP_PLACETOPCENTER = 2
Global Const PE_GLP_PLACERIGHT = 3
Global Const PE_GLP_PLACELEFT = 4

' legend layout
'
Global Const PE_GLL_PERCENTAGE = 0
Global Const PE_GLL_AMOUNT = 1
Global Const PE_GLL_CUSTOM = 2   ' for PEGetGraphOptionInfo, do not use in PESetGraphOptionInfo.

' bar sizes
Global Const PE_GBS_MINIMUMBARSIZE = 0
Global Const PE_GBS_SMALLBARSIZE = 1
Global Const PE_GBS_AVERAGEBARSIZE = 2
Global Const PE_GBS_LARGEBARSIZE = 3
Global Const PE_GBS_MAXIMUMBARSIZE = 4

' mmm, pie sizes
Global Const PE_GPS_MINIMUMPIESIZE = 64
Global Const PE_GPS_SMALLPIESIZE = 48
Global Const PE_GPS_AVERAGEPIESIZE = 32
Global Const PE_GPS_LARGEPIESIZE = 16
Global Const PE_GPS_MAXIMUMPIESIZE = 0

' detached pie slice
Global Const PE_GDPS_NODETACHMENT = 0
Global Const PE_GDPS_SMALLESTSLICE = 1
Global Const PE_GDPS_LARGESTSLICE = 2

' marker sizes
Global Const PE_GMS_SMALLMARKERS = 0
Global Const PE_GMS_MEDIUMSMALLMARKERS = 1
Global Const PE_GMS_MEDIUMMARKERS = 2
Global Const PE_GMS_MEDIUMLARGEMARKERS = 3
Global Const PE_GMS_LARGEMARKERS = 4

' marker shapes
Global Const PE_GMSP_RECTANGLESHAPE = 1
Global Const PE_GMSP_CIRCLESHAPE = 4
Global Const PE_GMSP_DIAMONDSHAPE = 5
Global Const PE_GMSP_TRIANGLESHAPE = 8

' chart color
Global Const PE_GCR_COLORCHART = 0
Global Const PE_GCR_BLACKANDWHITECHART = 1

' chart data points
Global Const PE_GDP_NONE = 0
Global Const PE_GDP_SHOWLABEL = 1
Global Const PE_GDP_SHOWVALUE = 2

' number formats
Global Const PE_GNF_NODECIMAL = 0
Global Const PE_GNF_ONEDECIMAL = 1
Global Const PE_GNF_TWODECIMAL = 2
Global Const PE_GNF_CURRENCYNODECIMAL = 3
Global Const PE_GNF_CURRENCYTWODECIMAL = 4
Global Const PE_GNF_PERCENTNODECIMAL = 5
Global Const PE_GNF_PERCENTONEDECIMAL = 6
Global Const PE_GNF_PERCENTTWODECIMAL = 7

' viewing angles
Global Const PE_GVA_STANDARDVIEW = 1
Global Const PE_GVA_TALLVIEW = 2
Global Const PE_GVA_TOPVIEW = 3
Global Const PE_GVA_DISTORTEDVIEW = 4
Global Const PE_GVA_SHORTVIEW = 5
Global Const PE_GVA_GROUPEYEVIEW = 6
Global Const PE_GVA_GROUPEMPHASISVIEW = 7
Global Const PE_GVA_FEWSERIESVIEW = 8
Global Const PE_GVA_FEWGROUPSVIEW = 9
Global Const PE_GVA_DISTORTEDSTDVIEW = 10
Global Const PE_GVA_THICKGROUPSVIEW = 11
Global Const PE_GVA_SHORTERVIEW = 12
Global Const PE_GVA_THICKSERIESVIEW = 13
Global Const PE_GVA_THICKSTDVIEW = 14
Global Const PE_GVA_BIRDSEYEVIEW = 15
Global Const PE_GVA_MAXVIEW = 16

Type PEGraphOptionInfo
    StructSize As Integer

    graphColour As Integer      ' PE_GCR_*, PE_UNCHANGED for no change
        
    ShowLegend As Integer       ' BOOL, PE_UNCHANGED for no change
    legendPosition As Integer   ' PE_GLP_*, if showLegend == 0, means no legend

' pie charts and doughut charts
    pieSize As Integer  ' PE_GPS_*, PE_UNCHANGED for no change
    detachedPieSlice As Integer ' PE_GDPS_* or PE_UNCHANGED for no change

' bar chart
    barSize As Integer  ' PE_GBS_*, PE_UNCHANGED for no change
    VerticalBars As Integer     ' BOOL, PE_UNCHANGED for no change

' markers (used for line and bar charts)
    markerSize As Integer       ' PE_GMS_*, PE_UNCHANGED for no change
    markerShape As Integer      ' PE_GMSP_*, PE_UNCHANGED for no change

' data points
    dataPoints  As Integer      ' PE_GDP_*, PE_UNCHANGED for no change
    dataValueNumberFormat As Integer    ' PE_GNF_*, PE_UNCHANGED for no change
' 3d
    viewingAngle As Integer     ' PE_GVA_*, PE_UNCHANGED for no change

    legendLayout As Integer     ' PE_GLL_*
End Type

Global Const PE_SIZEOF_GRAPH_OPTION_INFO = (14 * PE_WORD_LEN)


Declare Function PEGetGraphOptionInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, graphOptionInfo As PEGraphOptionInfo) As Integer
Declare Function PESetGraphOptionInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, graphOptionInfo As PEGraphOptionInfo) As Integer

' graph axes
Global Const PE_GGT_NOGRIDLINES = 0
Global Const PE_GGT_MINORGRIDLINES = 1
Global Const PE_GGT_MAJORGRIDLINES = 2
Global Const PE_GGT_MAJORANDMINORGRIDLINES = 3

' axis division method
Global Const PE_ADM_AUTOMATIC = 0
Global Const PE_ADM_MANUAL = 1

Type PEGraphAxisInfo
    StructSize As Integer

    groupAxisGridLine As Integer        ' PE_GGT_*, PE_UNCHANGED for no change
    dataAxisYGridLine As Integer        ' PE_GGT_*, PE_UNCHANGED for no change
    dataAxisY2GridLine As Integer       ' PE_GGT_*, PE_UNCHANGED for no change
    seriesAxisGridline As Integer       ' PE_GGT_*, PE_UNCHANGED for no change

    dataAxisYMinValue As Double
    dataAxisYMaxValue As Double
    dataAxisY2MinValue As Double
    dataAxisY2MaxValue As Double
    seriesAxisMinValue As Double
    seriesAxisMaxValue As Double

    dataAxisYNumberFormat As Integer    ' PE_GNF_*, PE_UNCHANGED for no change
    dataAxisY2NumberFormat As Integer   ' PE_GNF_*, PE_UNCHANGED for no change
    seriesAxisNumberFormat As Integer   ' PE_GNF_*, PE_UNCHANGED for no change

    dataAxisYAutoRange As Integer       ' BOOL, PE_UNCHANGED for no change
    dataAxisY2AutoRange As Integer      ' BOOL, PE_UNCHANGED for no change
    seriesAxisAutoRange As Integer      ' BOOL, PE_UNCHANGED for no change

    dataAxisYAutomaticDivision As Integer       ' PE_ADM_* or PE_UNCHANGED for no change
    dataAxisY2AutomaticDivision As Integer      ' PE_ADM_* or PE_UNCHANGED for no change
    seriesAxisAutomaticDivision As Integer      ' PE_ADM_* or PE_UNCHANGED for no change

    dataAxisYManualDivision As Long     ' if dataAxisYAutomaticDivision is PE_ADM_AUTOMATIC, this field is ignored
    dataAxisY2ManualDivision As Long    ' if dataAxisY2AutomaticDivision is PE_ADM_AUTOMATIC, this field is ignored
    seriesAxisManualDivision As Long    ' if seriesAxisAutomaticDivision is PE_ADM_AUTOMATIC, this field is ignored

    dataAxisYAutoScale As Integer       ' BOOL, PE_UNCHANGED for no change
    dataAxisY2AutoScale As Integer      ' BOOL, PE_UNCHANGED for no change
    seriesAxisAutoScale As Integer      ' BOOL, PE_UNCHANGED for no change
End Type

Global Const PE_SIZEOF_GRAPH_AXIS_INFO = (5 * PE_WORD_LEN) + (6 * PE_DOUBLE_LEN) + (9 * PE_WORD_LEN) + (3 * PE_LONG_LEN) + (3 * PE_WORD_LEN)


Declare Function PEGetGraphAxisInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, graphAxisInfo As PEGraphAxisInfo) As Integer
Declare Function PESetGraphAxisInfo Lib "crpe32.dll" (ByVal printJob As Integer, ByVal sectionN As Integer, ByVal GraphN As Integer, graphAxisInfo As PEGraphAxisInfo) As Integer

' Subreport(s)
Declare Function PEOpenSubreport Lib "crpe32.dll" (ByVal parentJob%, ByVal subreportName$) As Integer
Declare Function PECloseSubreport Lib "crpe32.dll" (ByVal printJob%) As Integer

' Formula Syntax

Global Const PE_FS_SIZE = 2

' Syntax Types
Global Const PE_FST_CRYSTAL = 0
Global Const PE_FST_BASIC = 1

Type PEFormulaSyntax
    StructSize As Integer               ' Set to PE_SIZEOF_FORMULA_SYNTAX
    FormulaSyntax(0 To 1) As Integer    ' PE_FS_* or PE_UNCHANGED.
End Type

Global Const PE_SIZEOF_FORMULA_SYNTAX = (3 * PE_WORD_LEN)

' PESetFormulaSyntax
' Use this API to set the syntax to use in the next (and all successive)
' formula API call(s).
' Set one of PE_FST_* into formulaSyntax[0];
' formulaSyntax[1] is reserved for internal use.
' ***Default Behaviour: If any Formula API is called before calling
'                      PESetFormulaSyntax then PE_FST_CRYSTAL is assumed.
Declare Function PESetFormulaSyntax Lib "crpe32.dll" (ByVal printJob As Integer, FormulaSyntax As PEFormulaSyntax) As Integer

' PEGetFormulaSyntax
' Indicates the syntax used by the formula addressed in the last formula API call.
' The syntax type is returned in formulaSyntax[0];
' formulaSyntax[1] is reserved for internal use.
' ***Default Behaviour: If this API is called before any Formula API is called
'                       then the values returned will be PE_UNCHANGED.
Declare Function PEGetFormulaSyntax Lib "crpe32.dll" (ByVal printJob As Integer, FormulaSyntax As PEFormulaSyntax) As Integer


Function PE_SECTION_CODE(sectionType%, groupN%, sectionN%) As Integer
' A function to create section codes:
' (This representation allows up to 25 groups and 40 sections of a given
' type, although Crystal Reports itself has no such limitations.)
    PE_SECTION_CODE = (((sectionType) * 1000) + ((groupN) Mod 25) + (((sectionN) Mod 40) * 25))
End Function

Function PE_AREA_CODE(sectionType%, groupN%) As Integer
'A function to create area codes:
    PE_AREA_CODE = PE_SECTION_CODE(sectionType, groupN, 0)
End Function

Function PE_GROUP_N(sectionCode%) As Integer
' Function to decode Group Number from section codes:
    PE_GROUP_N = ((sectionCode) Mod 25)
End Function

Function PE_SECTION_N(sectionCode) As Integer
' Function to decode Section Number from section codes:
   PE_SECTION_N = (((sectionCode \ 25) Mod 40))
End Function

Function PE_SECTION_TYPE(sectionCode%) As Integer
' Function to decode type from section codes:
    PE_SECTION_TYPE = ((sectionCode) \ 1000)
End Function

'Function to simplify PEGetVersion
Function PEVBGetVersion(ByVal component%) As Single
    Dim version As Integer
    Dim major As Integer
    Dim minor As Integer
    version = PEGetVersion(component)
    If version = 0 Then
        PEVBGetVersion = 0
    Else
        major = version / 256
        minor = version Mod 256
        PEVBGetVersion = major + (minor / 10)
    End If
End Function


Function PEGetJobStatus(ByVal job As Integer, Info As PEJobInfo) As Integer
' To work around the problem of 4 - Byte alignment the PEGetJobStatus
' call has been re-declared here. When your application calls PEGetJobStatus
' it is calling this function which in turn calls CRPE32.DLL.
Dim splitinfo As SplitPEJobInfo
Dim temp1 As Long
Dim temp2 As Long

splitinfo.StructSize = PE_SIZEOF_JOB_INFO
PEGetJobStatus = RealPEGetJobStatus(job, splitinfo)
If PEGetJobStatus <> -1 Then
    temp1 = splitinfo.NumRecordsRead1
    If temp1 < 0 Then
        temp1 = 65536 + temp1
    End If
    temp2 = splitinfo.NumRecordsRead2
    If temp2 < 0 Then
        temp2 = 65536 + temp2
    End If
    temp2 = temp2 * 65536
    Info.NumRecordsRead = temp1 + temp2
    
    temp1 = splitinfo.NumRecordsSelected1
    If temp1 < 0 Then
        temp1 = 65536 + temp1
    End If
    temp2 = splitinfo.NumRecordsSelected2
    If temp2 < 0 Then
        temp2 = 65536 + temp2
    End If
    temp2 = temp2 * 65536
    Info.NumRecordsSelected = temp1 + temp2
    
    temp1 = splitinfo.NumRecordsPrinted1
    If temp1 < 0 Then
        temp1 = 65536 + temp1
    End If
    temp2 = splitinfo.NumRecordsPrinted2
    If temp2 < 0 Then
        temp2 = 65536 + temp2
    End If
    Info.NumRecordsPrinted = temp1 + temp2
    Info.LatestPageN = splitinfo.LatestPageN
    Info.StartPageN = splitinfo.StartPageN
    Info.DisplayPageN = splitinfo.DisplayPageN
    Info.PrintEnded = splitinfo.PrintEnded
End If
End Function

' End Of Declarations
