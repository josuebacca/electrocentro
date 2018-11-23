**  Delphi Declarations for CRPE(32).DLL
**  ====================================
**
**  File     : CRDELPHI.PAS
**  Authors  : James Anderson, Andre Couturier, Frank Zimmerman
**  Date     : 14 Oct 1999
**  Purpose  : This file presents the API to the Crystal Reports
**             Print Engine DLL.  This file will compile and execute in
**	         Delphi 1, 2, 3, 4 and 5.
**  Language : Delphi 1, 2, 3, 4 and 5.
**  Copyright (c) 1991-1999 Seagate Software Information Management Group, Inc.
**
** Notes:  This file will create a CRDELPHI.DCU when you use compile and build.
**         To use the CRDELPHI.DCU, you will need to add the following line
**         of code to your 'uses' section of your Delphi application
**            eg.
**            uses
**               Classes, CRDELPHI, ...;
*)

unit CRDelphi;
{$X+}
{$A-}

interface

uses
  WinTypes, WinProcs;

const
  { Unchanged constants }
  PE_UNCHANGED       = -1;
  PE_UNCHANGED_COLOR = -2;
  PE_NO_COLOR        = $FFFFFFFF; { unsigned long }
  { Variable sizes }
  PE_WORD_LEN      =  2;
  PE_DWORD_LEN     =  4;
  PE_SMALLINT_LEN  =  2;
  PE_BOOL_LEN	   =  1;
  PE_ALLLINES      = -1;
  PE_LONGPTR_LEN   =  4;

{******************************************************************************}
{ Error Codes returned by the CRPE.DLL                                         }
{******************************************************************************}
  PE_ERR_NOERROR               = 0;
  PE_ERR_NOTENOUGHMEMORY       = 500;
  PE_ERR_INVALIDJOBNO          = 501;
  PE_ERR_INVALIDHANDLE         = 502;
  PE_ERR_STRINGTOOLONG         = 503;
  PE_ERR_NOSUCHREPORT          = 504;
  PE_ERR_NODESTINATION         = 505;
  PE_ERR_BADFILENUMBER         = 506;
  PE_ERR_BADFILENAME           = 507;
  PE_ERR_BADFIELDNUMBER        = 508;
  PE_ERR_BADFIELDNAME          = 509;
  PE_ERR_BADFORMULANAME        = 510;
  PE_ERR_BADSORTDIRECTION      = 511;
  PE_ERR_ENGINENOTOPEN         = 512;
  PE_ERR_INVALIDPRINTER        = 513;
  PE_ERR_PRINTFILEEXISTS       = 514;
  PE_ERR_BADFORMULATEXT        = 515;
  PE_ERR_BADGROUPSECTION       = 516;
  PE_ERR_ENGINEBUSY            = 517;
  PE_ERR_BADSECTION            = 518;
  PE_ERR_NOPRINTWINDOW         = 519;
  PE_ERR_JOBALREADYSTARTED     = 520;
  PE_ERR_BADSUMMARYFIELD       = 521;
  PE_ERR_NOTENOUGHSYSRES       = 522;
  PE_ERR_BADGROUPCONDITION     = 523;
  PE_ERR_JOBBUSY               = 524;
  PE_ERR_BADREPORTFILE         = 525;
  PE_ERR_NODEFAULTPRINTER      = 526;
  PE_ERR_SQLSERVERERROR        = 527;
  PE_ERR_BADLINENUMBER         = 528;
  PE_ERR_DISKFULL              = 529;
  PE_ERR_FILEERROR             = 530;
  PE_ERR_INCORRECTPASSWORD     = 531;
  PE_ERR_BADDATABASEDLL        = 532;
  PE_ERR_BADDATABASEFILE       = 533;
  PE_ERR_ERRORINDATABASEDLL    = 534;
  PE_ERR_DATABASESESSION       = 535;
  PE_ERR_DATABASELOGON         = 536;
  PE_ERR_DATABASELOCATION      = 537;
  PE_ERR_BADSTRUCTSIZE         = 538;
  PE_ERR_BADDATE               = 539;
  PE_ERR_BADEXPORTDLL          = 540;
  PE_ERR_ERRORINEXPORTDLL      = 541;
  PE_ERR_PREVATFIRSTPAGE       = 542;
  PE_ERR_NEXTATLASTPAGE        = 543;
  PE_ERR_CANNOTACCESSREPORT    = 544;
  PE_ERR_USERCANCELLED         = 545;
  PE_ERR_OLE2NOTLOADED         = 546;
  PE_ERR_BADCROSSTABGROUP      = 547;
  PE_ERR_NOCTSUMMARIZEDFIELD   = 548;
  PE_ERR_DESTINATIONNOTEXPORT  = 549;
  PE_ERR_INVALIDPAGENUMBER     = 550;
  PE_ERR_NOTSTOREDPROCEDURE    = 552;
  PE_ERR_INVALIDPARAMETER      = 553;
  PE_ERR_GRAPHNOTFOUND         = 554;
  PE_ERR_INVALIDGRAPHTYPE      = 555;
  PE_ERR_INVALIDGRAPHDATA      = 556;
  PE_ERR_CANNOTMOVEGRAPH       = 557;
  PE_ERR_INVALIDGRAPHTEXT      = 558;
  PE_ERR_INVALIDGRAPHOPT       = 559;
  PE_ERR_BADSECTIONHEIGHT            = 560;
  PE_ERR_BADVALUETYPE                = 561;
  PE_ERR_INVALIDSUBREPORTNAME        = 562;
  PE_ERR_NOPARENTWINDOW              = 564; {dialog parent window}
  PE_ERR_INVALIDZOOMFACTOR           = 565; {zoom factor}
  PE_ERR_PAGESIZEOVERFLOW            = 567;
  PE_ERR_LOWSYSTEMRESOURCES          = 568;
  PE_ERR_BADGROUPNUMBER              = 570;
  PE_ERR_INVALIDOBJECTFORMATNAME     = 571;
  PE_ERR_INVALIDNEGATIVEVALUE        = 572;
  PE_ERR_INVALIDMEMORYPOINTER        = 573;
  PE_ERR_INVALIDOBJECTTYPE           = 574;
  PE_ERR_INVALIDGRAPHDATATYPE        = 577;
  PE_ERR_INVALIDSUBREPORTLINKNUMBER  = 582;
  PE_ERR_SUBREPORTLINKEXIST          = 583;
  PE_ERR_BADROWCOLVALUE              = 584;
  PE_ERR_INVALIDSUMMARYNUMBER        = 585;
  PE_ERR_INVALIDGRAPHDATAFIELDNUMBER = 586;
  PE_ERR_INVALIDSUBREPORTNUMBER      = 587;
  PE_ERR_INVALIDFIELDSCOPE           = 588;
  PE_ERR_FIELDINUSE                  = 590;
  PE_ERR_INVALIDPARAMETERNUMBER      = 594;
  PE_ERR_INVALIDPAGEMARGINS          = 595;
  PE_ERR_REPORTONSECUREQUERY         = 596;
  PE_ERR_CANNOTOPENSECUREQUERY       = 597;
  PE_ERR_INVALIDSECTIONNUMBER        = 598;
  PE_ERR_SQLSERVERNOTOPENED          = 599;
  PE_ERR_TABLENAMEEXIST              = 606;
  PE_ERR_INVALIDCURSOR               = 607;
  PE_ERR_FIRSTPASSNOTFINISHED        = 608;
  PE_ERR_CREATEDATASOURCE            = 609;
  PE_ERR_CREATEDRILLDOWNPARAMETERS   = 610;
  PE_ERR_CHECKFORDATASOURCECHANGES   = 613;
  PE_ERR_STARTBACKGROUNDPROCESSING   = 614;
  PE_ERR_SQLSERVERINUSE              = 619;
  PE_ERR_GROUPSORTFIELDNOTSET        = 620;
  PE_ERR_CANNOTSETSAVESUMMARIES      = 621;
  PE_ERR_LOADOLAPDATABASEMANAGER     = 622;
  PE_ERR_OPENOLAPCUBE                = 623;
  PE_ERR_READOLAPCUBEDATA            = 624;
  PE_ERR_CANNOTSAVEQUERY             = 626;
  PE_ERR_CANNOTREADQUERYDATA         = 627;
  PE_ERR_MAINREPORTFIELDLINKED       = 629;
  PE_ERR_INVALIDMAPPINGTYPEVALUE     = 630;
  PE_ERR_HITTESTFAILED               = 636;
  PE_ERR_BADSQLEXPRESSIONNAME        = 637; { no SQL expression by the specified *name* exists in this report. }
  PE_ERR_BADSQLEXPRESSIONNUMBER      = 638; { no SQL expression by the specified *number* exists in this report. }
  PE_ERR_BADSQLEXPRESSIONTEXT        = 639; { not a valid SQL expression }
  PE_ERR_INVALIDDEFAULTVALUEINDEX    = 641; { invalid index for default value of a parameter. }
  PE_ERR_NOMINMAXVALUE               = 642; { the specified PE_PF_* type does not have min/max values. }
  PE_ERR_INCONSISTANTTYPES           = 643; { if both min and max values are specified in PESetParameterMinMaxValue, }
                                            { the value types for the min and max must be the same. }
  PE_ERR_CANNOTLINKTABLES            = 645;
  PE_ERR_CREATEROUTER                = 646;
  PE_ERR_INVALIDFIELDINDEX           = 647;
  PE_ERR_INVALIDGRAPHTITLETYPE       = 648;
  PE_ERR_INVALIDGRAPHTITLEFONTTYPE   = 649;
  PE_ERR_PARAMTYPEDIFFERENT          = 650; { the type used in a add/set value API for a }
                                            { parameter differs with it's existing type. }
  PE_ERR_INCONSISTANTRANGETYPES      = 651; { the value type for both start & end range }
                                            { values must be the same. }
  PE_ERR_RANGEORDISCRETE             = 652; { an operation was attempted on a discrete parameter that is }
                                            { only legal for range parameters or vice versa. }
  PE_ERR_NOTMAINREPORT               = 654; { an operation was attempted that is disallowed for subreports. }
  PE_ERR_INVALIDCURRENTVALUEINDEX    = 655; { invalid index for current value of a parameter. }
  PE_ERR_LINKEDPARAMVALUE            = 656; { operation illegal on linked parameter. }
  PE_ERR_INVALIDPARAMETERRANGEINFO   = 672; { Invalid PE_RI_* combination}
  PE_ERR_INVALIDSORTMETHODINDEX      = 674; { Invalid sort method index. }
  PE_ERR_INVALIDGRAPHSUBTYPE         = 675; { Invalid PE_GST_* or }
                                            { PE_GST_* does not match PE_GT_* or }
                                            { PE_GST_* current graph type. }
  PE_ERR_BADGRAPHOPTIONINFO          = 676; { one of them members of PEGraphOptionInfo is out of range. }
  PE_ERR_BADGRAPHAXISINFO            = 677; { one of them members of PEGraphAxisInfo is out of range. }

  PE_ERR_NOTEXTERNALSUBREPORT        = 680; { the subreport is not imported }
  PE_ERR_INVALIDPARAMETERVALUE       = 687;
  PE_ERR_INVALIDFORMULASYNTAXTYPE    = 688; { specified formula syntax not in PE_FST_* }
  PE_ERR_INVALIDCROPVALUE            = 689;
  PE_ERR_INVALIDCOLLATIONVALUE       = 690;
  PE_ERR_STARTPAGEGREATERSTOPPAGE    = 691;
  PE_ERR_INVALIDEXPORTFORMAT		 = 692;


  PE_ERR_OTHERERROR                  = 997;
  PE_ERR_INTERNALERROR               = 998; { programming error }
  PE_ERR_NOTIMPLEMENTED              = 999;


type
  AreaType = Word;

  DWord = LongInt;

{******************************************************************************}
{ Open and Close Print Engine                                                  }
{******************************************************************************}
function PEOpenEngine: Bool stdcall;

{$ifdef VER120}
procedure PECloseEngine() stdcall;
{$else}
procedure PECloseEngine stdcall;
{$endif VER120}

function PECanCloseEngine: Bool stdcall;

{******************************************************************************}
{ Get Version Info                                                             }
{******************************************************************************}
const
  PE_GV_DLL                  = 100;
  PE_GV_ENGINE               = 200;

function PEGetVersion (versionRequested : Smallint): Smallint stdcall;


{******************************************************************************}
{ Open, Print and Close Report                                                 }
{  - used when no changes needed to report - Complete Default                  }
{******************************************************************************}
function PEPrintReport (
  reportFilePath    : PChar;
  toDefaultPrinter  : Bool;
  toWindow          : Bool;
  title             : PChar;
  left              : integer;
  top               : integer;
  width             : integer;
  height            : integer;
  style             : DWord;
  parentWindow      : Hwnd): Smallint stdcall;


{******************************************************************************}
{ Report Options                                                               }
{******************************************************************************}
const
  PE_RPTOPT_CVTDATETIMETOSTR  = 0;
  PE_RPTOPT_CVTDATETIMETODATE = 1;
  PE_RPTOPT_KEEPDATETIMETYPE  = 2;

  PE_RPTOPT_PROMPT_NONE   = 0;
  PE_RPTOPT_PROMPT_NORMAL = 1;
  PE_RPTOPT_PROMPT_ALWAYS = 2;

type
  PEReportOptions = record
    StructSize                       : Word;  	  {initialize to PE_SIZEOF_REPORT_OPTIONS}
    saveDataWithReport               : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    saveSummariesWithReport          : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    useIndexForSpeed                 : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    translateDOSStrings              : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    translateDOSMemos                : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    convertDateTimeType              : Smallint;  {a PE_RPTOPT_ value, except use PE_UNCHANGED for no change}
    convertNullFieldToDefault        : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    morePrintEngineErrorMessages     : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    caseInsensitiveSQLData           : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    verifyOnEveryPrint               : Smallint;  {BOOL value, except use PE_UNCHAGED for no change}
    zoomMode                         : Smallint;  {a PE_ZOOM_ constant, except use PE_UNCHANGED for no change}
    hasGroupTree                     : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    dontGenerateDataForHiddenObjects : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    performGroupingOnServer          : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
    promptMode                       : Smallint;  {PE_RPTOPT_PROMPT_NONE, PE_RPTOPT_PROMPT_NORMAL, 
                                                   PE_RPTOPT_PROMPT_ALWAYS, PE_UNCHANGED}
    selectDistinctRecords            : Smallint;  {BOOL value, except use PE_UNCHANGED for no change}
  end;

const
  PE_SIZEOF_REPORT_OPTIONS = SizeOf(PEReportOptions);

function PEGetReportOptions(
  printJob          : Smallint;
  var reportOptions : PEReportOptions): Bool stdcall;

function PESetReportOptions(
  printJob          : Smallint;
  var reportOptions : PEReportOptions): Bool stdcall;


{******************************************************************************}
{ Open and Close Print Job (i.e. Report)                                       }
{******************************************************************************}
function PEOpenPrintJob (reportFilePath : PChar): Smallint stdcall;
function PEClosePrintJob (printJob : Smallint): Bool stdcall;


{******************************************************************************}
{ Start and Cancel Print Job                                                   }
{  - i.e. print the Report, usually after changing report                      }
{******************************************************************************}
function PEStartPrintJob (
  printJob      : Smallint;
  waitUntilDone : Bool): Bool stdcall;

procedure PECancelPrintJob (printJob : Smallint) stdcall;


{******************************************************************************}
{ Open and Close Subreport                                                     }
{******************************************************************************}
function PEOpenSubreport (
  parentJob     : Smallint;
  subreportName : PChar): Smallint stdcall;

function PECloseSubreport (printJob : Smallint): Bool stdcall;


{******************************************************************************}
{ Print Job Status                                                             }
{******************************************************************************}
function PEIsPrintJobFinished (printJob : Smallint): Bool stdcall;

{ Status returns for PEGetJobStatus }
const
  PE_JOBNOTSTARTED = 1;
  PE_JOBINPROGRESS = 2;
  PE_JOBCOMPLETED  = 3;
  PE_JOBFAILED     = 4;
  PE_JOBCANCELLED  = 5;
  PE_JOBHALTED     = 6;

type
  PEJobInfo = record
    StructSize         : Word;  (*initialize to SizeOf(PEJobInfo)*)
    NumRecordsRead     : DWord;
    NumRecordsSelected : DWord;
    NumRecordsPrinted  : DWord;
    DisplayPageN       : Word;  (*the page being displayed in window*)
    LatestPageN        : Word;  (*the page being generated*)
    StartPageN         : Word;  (*user opted, default to 1*)
    PrintEnded         : Bool;  (*full report print completed?*)
  end;

const
  PE_SIZEOF_JOB_INFO = SizeOf(PEJobInfo);

function PEGetJobStatus (
  printJob    : Smallint;
  var jobInfo : PEJobInfo): Smallint stdcall;


{******************************************************************************}
{ Controlling Dialogs                                                          }
{******************************************************************************}
function PESetDialogParentWindow (
  printJob     : Smallint;
  parentWindow : HWnd): Bool stdcall;

function PEEnableProgressDialog (
  printJob : Smallint;
  enable   : Bool): Bool stdcall;


{******************************************************************************}
{ Controlling Parameter Field Prompting Dialog                                 }
{   Set boolean to indicate whether CRPE is allowed to prompt                  }
{   for parameter values during printing.                                      }
{******************************************************************************}
function PEGetAllowPromptDialog (printJob : Smallint): Bool stdcall;
function PESetAllowPromptDialog (
  printJob         : Smallint;
  showPromptDialog : Bool): Bool stdcall;


{******************************************************************************}
{ Print Job Error Codes and Messages                                           }
{******************************************************************************}
function PEGetErrorCode (printJob : Smallint): Smallint stdcall;

function PEGetErrorText (
  printJob       : Smallint;
  var textHandle : HWnd;
  var textLength : Smallint): Bool stdcall;

function PEGetHandleString (
  textHandle   : HWnd;
  buffer       : PChar;
  bufferLength : Smallint): Bool stdcall;


{******************************************************************************}
{ Print Date                                                                   }
{******************************************************************************}
function PEGetPrintDate (
  printJob  : Smallint;
  var year  : Smallint;
  var month : Smallint;
  var day   : Smallint): Bool stdcall;

function PESetPrintDate (
  printJob : Smallint;
  year     : Smallint;
  month    : Smallint;
  day      : Smallint): Bool stdcall;


{******************************************************************************}
{ Encoding and Decoding Section Codes                                          }
{******************************************************************************}
{ Section constants }
{ For use with PE_SECTION_CODE, PE_SECTION_TYPE,
  PE_GROUP_N and PE_SECTION_N functions }
const
  PE_SECT_REPORT_HEADER = 1;
  PE_SECT_PAGE_HEADER   = 2;
  PE_SECT_GROUP_HEADER  = 3;
  PE_SECT_DETAIL        = 4;
  PE_SECT_GROUP_FOOTER  = 5;
  PE_SECT_PAGE_FOOTER   = 7;
  PE_SECT_REPORT_FOOTER = 8;

{ The old section constants with comment showing them in terms of the new: }
{ Note that PE_GRANDTOTALSECTION and PE_SUMMARYSECTION both map }
{ to PE_SECT_REPORT_FOOTER. }
  PE_ALLSECTIONS       =    0;
  PE_HEADERSECTION     = 2000;  {PE_SECTION_CODE (PE_SECT_PAGE_HEADER, 0, 0)}
  PE_FOOTERSECTION     = 7000;  {PE_SECTION_CODE (PE_SECT_PAGE_FOOTER, 0, 0)}
  PE_TITLESECTION      = 1000;  {PE_SECTION_CODE (PE_SECT_REPORT_HEADER, 0, 0)}
  PE_SUMMARYSECTION    = 8000;  {PE_SECTION_CODE (PE_SECT_REPORT_FOOTER, 0, 0)}
  PE_GROUPHEADER       = 3000;  {PE_SECTION_CODE (PE_SECT_GROUP_HEADER, 0, 0)}
  PE_GROUPFOOTER       = 5000;  {PE_SECTION_CODE (PE_SECT_GROUP_FOOTER, 0, 0)}
  PE_DETAILSECTION     = 4000;  {PE_SECTION_CODE (PE_SECT_DETAIL, 0, 0)}
  PE_GRANDTOTALSECTION = PE_SUMMARYSECTION;

{ A function to create section codes: }
{ This representation allows up to 25 groups and 40 sections of a given }
{ type, although Crystal Reports itself has no such limitations. }
function PE_SECTION_CODE (sectionType, groupN, sectionN : Smallint): Smallint stdcall;

{ Functions to decode section and area codes: }
function PE_SECTION_TYPE (sectionCode : Smallint): Smallint stdcall;
function PE_GROUP_N (sectionCode : Smallint): Smallint stdcall;
function PE_SECTION_N (sectionCode : Smallint): Smallint stdcall;

{ A function to create section codes: }
{  This representation allows up to 25 groups although Crystal Reports }
{  itself has no such limitations. }
function PE_AREA_CODE (sectionType, groupN : Smallint): Smallint stdcall;


{******************************************************************************}
{ Group Conditions (i.e. group breaks)                                         }
{******************************************************************************}
const
  PE_SF_MAX_NAME_LENGTH = 50;

  PE_SF_DESCENDING   = 0;
  PE_SF_ASCENDING    = 1;
  PE_SF_ORIGINAL     = 2;  {only for group condition}
  PE_SF_SPECIFIED    = 3;  {only for group condition}

  { use PE_ANYCHANGE for all field types except Date }
  PE_GC_ANYCHANGE    = 0;

  { use these constants for Date fields }
  PE_GC_DAILY        = 0;
  PE_GC_WEEKLY       = 1;
  PE_GC_BIWEEKLY     = 2;
  PE_GC_SEMIMONTHLY  = 3;
  PE_GC_MONTHLY      = 4;
  PE_GC_QUARTERLY    = 5;
  PE_GC_SEMIANNUALLY = 6;
  PE_GC_ANNUALLY     = 7;

  { use these constants for Time  and DateTime fields }
  PE_GC_BYSECOND     =  8;
  PE_GC_BYMINUTE     =  9;
  PE_GC_BYHOUR       = 10;
  PE_GC_BYAMPM       = 11;

  { use these constants for Boolean fields }
  PE_GC_TOYES        = 1;
  PE_GC_TONO         = 2;
  PE_GC_EVERYYES     = 3;
  PE_GC_EVERYNO      = 4;
  PE_GC_NEXTISYES    = 5;
  PE_GC_NEXTISNO     = 6;

function PESetGroupCondition (
  printJob       : Smallint;
  sectionCode    : Smallint;
  conditionField : PChar;
  condition      : Smallint;  (* a PE_GC_ constant *)
  sortDirection  : Smallint   (* a PE_SF_ constant *)
  ): Bool stdcall;

function PEGetNGroups (printJob : Smallint): Smallint stdcall;

{ for PEGetGroupCondition, condition encodes both }
{ the condition and the type of the condition field }
const
  PE_GC_CONDITIONMASK = $00FF;
  PE_GC_TYPEMASK      = $0F00;

  PE_GC_TYPEOTHER     = $0000;
  PE_GC_TYPEDATE      = $0200;
  PE_GC_TYPEBOOLEAN   = $0400;
  PE_GC_TYPETIME      = $0800;

function PEGetGroupCondition (
  printJob                 : Smallint;
  sectionCode              : Smallint;
  var conditionFieldHandle : Hwnd;
  var conditionFieldLength : Smallint;
  var condition            : Smallint;
  var sortDirection        : Smallint): Bool stdcall;

const
  PE_FIELD_NAME_LEN = 512;

  PE_GO_TBN_ALL_GROUPS_UNSORTED = 0;
  PE_GO_TBN_ALL_GROUPS_SORTED   = 1;
  PE_GO_TBN_TOP_N_GROUPS        = 2;
  PE_GO_TBN_BOTTOM_N_GROUPS     = 3;

type
  PEFieldNameType = array[0..PE_FIELD_NAME_LEN-1] of Char;
  PEGroupOptions = record
    StructSize:                Word;
    (* when setting, pass a PE_GC_ constant, or PE_UNCHANGED for no change.
       when getting, use PE_GC_TYPEMASK and PE_GC_CONDITIONMASK to
       decode the condition.*)
    condition                 : Smallint;  {a PE_GC_ constant, or PE_UNCHANGED for no change.}
    FieldName                 : PEFieldNameType; { formula form, or empty for no change.}
    sortDirection             : Smallint;  { a PE_SF_ const, or PE_UNCHANGED for no change.}
    repeatGroupHeader         : Smallint;  { BOOL value, or PE_UNCHANGED for no change.}
    keepGroupTogether         : Smallint;  { BOOL value, or PE_UNCHANGED for no change.}
    topOrBottomNGroups        : Smallint;  { a PE_GO_TBN_ constant, or PE_UNCHANGED for no change.}
    topOrBottomNSortFieldName : PEFieldNameType; { formula form, or empty for no change.}
    nTopOrBottomGroups        : Smallint;  { the number of groups to keep,
                                             0 for all, or PE_UNCHANGED for no change.}
    discardOtherGroups        : Smallint;  { BOOL value, or PE_UNCHANGED for no change.}
    ignored                   : Integer;   { ignored - used for 4-byte alignment }
    hierarchicalSorting       : Smallint;  { Boolean or PE_UNCHANGED for no change.}
    instanceIDField           : PEFieldNameType; { for hierarchical grouping }
    parentIDField             : PEFieldNameType; { for hierarchical grouping }
    groupIndent               : LongInt;   { twips }
  end;

const
  PE_SIZEOF_GROUP_OPTIONS = SizeOf(PEGroupOptions);

function PEGetGroupOptions (
  printJob         : Smallint;
  groupN           : Smallint;
  var groupOptions : PEGroupOptions): Bool stdcall;

function PESetGroupOptions (
  printJob         : Smallint;
  groupN           : Smallint;
  var groupOptions : PEGroupOptions): Bool stdcall;


{******************************************************************************}
{ Getting Number of Sections and Section Code                                  }
{******************************************************************************}
function PEGetNSections (printJob: Smallint): Smallint stdcall;
{ return -1 if failed }

function PEGetSectionCode (
  printJob: Smallint;
  sectionN: Smallint): Smallint stdcall;

function PEGetNSectionsInArea (printJob: Smallint; areaCode: Smallint): Smallint stdcall;
{ return -1 if failed }


{******************************************************************************}
{ Section Height                                                               }
{   The PEGetMinimumSectionHeight/PESetMinimumSectionHeight API Calls          }
{   are obsolete, please use PEGetSectionHeight/PESetSectionHeight instead.    }
{   The obsolete API calls will still work in older applications but please    }
{   use the NEW calls for all new Development.                                 }
{******************************************************************************}
function PESetSectionHeight (
  printJob    : Smallint;
  sectionCode :	Smallint;
  Height      : Smallint): Bool stdcall;
  {Height is in twips }

function PEGetSectionHeight (
  printJob :		Smallint;
  sectionCode : 	Smallint;
  var Height :	Smallint): Bool stdcall;
  {Height is in twips }


{******************************************************************************}
{ Section or Area Format                                                       }
{******************************************************************************}
type
  PESectionOptions = record
    (*Initialize to PE_SIZEOF_SECTION_OPTIONS*)
    StructSize           : Word;
    Visible              : Smallint; (* BOOLEAN values, except use PE_UNCHANGED *)
    newPageBefore        : Smallint; (* to preserve the existing settings       *)
    newPageAfter         : Smallint;
    keepTogether         : Smallint;
    suppressBlankSection : Smallint;
    resetPageNAfter      : Smallint;
    printAtBottomOfPage  : Smallint;
    backgroundColor      : DWord;    (* Use PE_UNCHANGED_COLOR to preserve the
                                        existing color. *)
    underlaySection      : Smallint; (* BOOLEAN values, except use PE_UNCHANGED  *)
    showArea             : Smallint; (* to preserve the existing settings        *)
    freeFormPlacement    : Smallint;
    reserveMinimumPageFooter : Smallint; { BOOLEAN or PE_UNCHANGED; Sets the size    }
										 { of the Page Footer *area* to the size     }
										 { of the largest Page Footer *section*.     }
										 { Used with PEGetAreaFormat/PESetAreaFormat }
										 { ignored when used with PEGetSectionFormat }
										 { PESetSectionFormat.                       }
  end;

const PE_SIZEOF_SECTION_OPTIONS = SizeOf(PESectionOptions);

function PEGetSectionFormat (
  printJob    : Smallint;
  sectionCode : Smallint;
  var options : PESectionOptions): Bool stdcall;

function PESetSectionFormat (
  printJob    : Smallint;
  sectionCode : Smallint;
  var options : PESectionOptions): Bool stdcall;

function PEGetAreaFormat (
  printJob    : Smallint;
  areaCode    : Smallint;
  var options : PESectionOptions): Bool stdcall;

function PESetAreaFormat (
  printJob    : Smallint;
  areaCode    : Smallint;
  var options : PESectionOptions): Bool stdcall;


{******************************************************************************}
{ Setting Font info                                                            }
{******************************************************************************}
const
  PE_FIELDS = $0001;
  PE_TEXT   = $0002;

function PESetFont (
  printJob     : Smallint;
  sectionCode  : Smallint;
  scopeCode    : Smallint;
  faceName     : PChar;    (* 0 for no change               *)
  fontFamily   : Smallint; (* FF_DONTCARE for no change     *)
  fontPitch    : Smallint; (* DEFAULT_PITCH for no change   *)
  charSet      : Smallint; (* DEFAULT_CHARSET for no change *)
  pointSize    : Smallint; (* 0 for no change               *)
  isItalic     : Smallint; (* PE_UNCHANGED for no change    *)
  isUnderlined : Smallint; (* PE_UNCHANGED for no change    *)
  isStruckOut  : Smallint; (* PE_UNCHANGED for no change    *)
  weight       : Smallint; (* 0 for no change               *)
  ):Bool stdcall;


{******************************************************************************}
{ Subreports                                                                   }
{******************************************************************************}
function PEGetNSubreportsInSection (
  printJob: Smallint;
  sectionCode : Smallint): Smallint stdcall;

function PEGetNthSubreportInSection (
  printJob: Smallint;
  sectionCode: Smallint;
  subreportN : Smallint): DWord stdcall;

const
  PE_SUBREPORT_NAME_LEN = 128;

type
  PESubreportNameType = array [0..PE_SUBREPORT_NAME_LEN-1] of Char;
  PESubreportInfo = record
    structSize : Word;          { Initialize to PE_SIZEOF_SUBREPORT_INFO }
    name       : PESubreportNameType;
    NLinks     : Smallint;      { Number of Links }
    IsOnDemand : Smallint;      { TRUE if subreport is On Demand }

    { New for SCR8 }
    isExternal   : Smallint;    { 1: the subreport is imported; 0: otherwise. }
    reimportOption : Smallint;  { PE_SRI_ONOPENJOB or PE_SRI_ONFUNCTIONCALL }
  end;

const
  PE_SIZEOF_SUBREPORT_INFO = SizeOf(PESubreportInfo);

function PEGetSubreportInfo (
  printJob          : Smallint;
  subreportHandle   : DWord;
  var subreportInfo : PESubreportInfo): Bool stdcall;

function PEReimportSubreport(
  printJob          : Smallint;
  subreportHandle   : DWord;
  var linkChanged   : Bool;
  var reimported    : Bool): Bool stdcall;

{******************************************************************************}
{ Graphing                                                                     }
{******************************************************************************}
{------------------------------------------------------------------------------}
{ Graph Type                                                                   }
{------------------------------------------------------------------------------}
const
  { Graph Type }
  PE_GT_BARCHART           =   0;
  PE_GT_LINECHART          =   1;
  PE_GT_AREACHART          =   2;
  PE_GT_PIECHART           =   3;
  PE_GT_DOUGHNUTCHART      =   4;
  PE_GT_THREEDRISERCHART   =   5;
  PE_GT_THREEDSURFACECHART =   6;
  PE_GT_SCATTERCHART       =   7;
  PE_GT_RADARCHART         =   8;
  PE_GT_BUBBLECHART        =   9;
  PE_GT_STOCKCHART         =  10;
  { These next two are for PEGetGraphTypeInfo only }
  { Do not use in PESetGraphTypeInfo }
  PE_GT_USERDEFINEDCHART   =  50;
  PE_GT_UNKNOWNTYPECHART   = 100;

  { Graph Subtype }
  { Bar charts }
  PE_GST_SIDEBYSIDEBARCHART                =    0;
  PE_GST_STACKEDBARCHART                   =    1;
  PE_GST_PERCENTBARCHART                   =    2;
  PE_GST_FAKED3DSIDEBYSIDEBARCHART         =    3;
  PE_GST_FAKED3DSTACKEDBARCHART            =    4;
  PE_GST_FAKED3DPERCENTBARCHART            =    5;

  { Line charts }
  PE_GST_REGULARLINECHART                  =   10;
  PE_GST_STACKEDLINECHART                  =   11;
  PE_GST_PERCENTAGELINECHART               =   12;
  PE_GST_LINECHARTWITHMARKERS              =   13;
  PE_GST_STACKEDLINECHARTWITHMARKERS       =   14;
  PE_GST_PERCENTAGELINECHARTWITHMARKERS    =   15;

  { Area charts }
  PE_GST_ABSOLUTEAREACHART                 =   20;
  PE_GST_STACKEDAREACHART                  =   21;
  PE_GST_PERCENTAREACHART                  =   22;
  PE_GST_FAKED3DABSOLUTEAREACHART          =   23;
  PE_GST_FAKED3DSTACKEDAREACHART           =   24;
  PE_GST_FAKED3DPERCENTAREACHART           =   25;

  { Pie charts }
  PE_GST_REGULARPIECHART                   =   30;
  PE_GST_FAKED3DREGULARPIECHART            =   31;
  PE_GST_MULTIPLEPIECHART                  =   32;
  PE_GST_MULTIPLEPROPORTIONALPIECHART      =   33;

  { Doughnut charts }
  PE_GST_REGULARDOUGHNUTCHART              =   40;
  PE_GST_MULTIPLEDOUGHNUTCHART             =   41;
  PE_GST_MULTIPLEPROPORTIONALDOUGHNUTCHART =   42;

  { 3D Riser charts }
  PE_GST_THREEDREGULARCHART                =   50;
  PE_GST_THREEDPYRAMIDCHART                =   51;
  PE_GST_THREEDOCTAGONCHART                =   52;
  PE_GST_THREEDCUTCORNERSCHART             =   53;

  { 3D surface charts }
  PE_GST_THREEDSURFACEREGULARCHART         =   60;
  PE_GST_THREEDSURFACEWITHSIDESCHART       =   61;
  PE_GST_THREEDSURFACEHONEYCOMBCHART       =   62;

  { Scatter charts }
  PE_GST_XYSCATTERCHART                    =   70;
  PE_GST_XYSCATTERDUALAXISCHART            =   71;
  PE_GST_XYSCATTERWITHLABELSCHART          =   72;
  PE_GST_XYSCATTERDUALAXISWITHLABELSCHART  =   73;

  { Radar charts }
  PE_GST_REGULARRADARCHART                 =   80;
  PE_GST_STACKEDRADARCHART                 =   81;
  PE_GST_RADARDUALAXISCHART                =   82;

  { Bubble charts }
  PE_GST_REGULARBUBBLECHART                =   90;
  PE_GST_DUALAXISBUBBLECHART               =   91;

  { Stocked charts }
  PE_GST_HIGHLOWCHART                      =  100;
  PE_GST_HIGHLOWDUALAXISCHART              =  101;
  PE_GST_HIGHLOWOPENCHART                  =  102;
  PE_GST_HIGHLOWOPENDUALAXISCHART          =  103;
  PE_GST_HIGHLOWOPENCLOSECHART             =  104;
  PE_GST_HIGHLOWOPENCLOSEDUALAXISCHART     =  105;

  PE_GST_UNKNOWNSUBTYPECHART               = 1000;

type
  PEGraphTypeInfo = record
    StructSize   : Word;
    graphType    : Smallint; { PE_GT_*, PE_UNCHANGED for no change }
    graphSubtype : Smallint; { PE_GST_*, PE_UNCHANGED for no change }
  end;

const
 PE_SIZEOF_GRAPH_TYPE_INFO = SizeOf(PEGraphTypeInfo);

function PEGetGraphTypeInfo (
  printJob          : Smallint;
  sectionN          : Smallint;
  graphN            : Smallint;
  var graphTypeInfo : PEGraphTypeInfo): Bool stdcall;

function PESetGraphTypeInfo (
  printJob          : Smallint;
  sectionN          : Smallint;
  graphN            : Smallint;
  var graphTypeInfo : PEGraphTypeInfo): Bool stdcall;

{------------------------------------------------------------------------------}
{ Graph Text                                                                   }
{------------------------------------------------------------------------------}
const
  PE_GTT_TITLE           = 0;
  PE_GTT_SUBTITLE        = 1;
  PE_GTT_FOOTNOTE        = 2;
  PE_GTT_SERIESTITLE     = 3;
  PE_GTT_GROUPSTITLE     = 4;
  PE_GTT_XAXISTITLE      = 5;
  PE_GTT_YAXISTITLE      = 6;
  PE_GTT_ZAXISTITLE      = 7;

function PEGetGraphTextInfo (
  printJob        : Smallint;
  sectionN        : Smallint;
  graphN          : Smallint;
  titleType       : Word;     { PE_GTT_ constant }
  var title       : Hwnd;
  var titleLength : Smallint): Bool stdcall;

function PESetGraphTextInfo (
  printJob  : Smallint;
  sectionN  : Smallint;
  graphN    : Smallint;
  titleType : Word;   { PE_GTT_ constant }
  title     : PChar): Bool stdcall;

{ enable/disable graph default titles }
function PESetGraphTextDefaultOption (
  printJob   : Smallint;
  sectionN   : Smallint; 
  graphN     : Smallint; 
  titleType  : Word;  { PE_GTT_ constant }
  useDefault : Bool): Bool stdcall;

function PEGetGraphTextDefaultOption (
  printJob       : Smallint;
  sectionN       : Smallint; 
  graphN         : Smallint; 
  titleType      : Word;   { PE_GTT_ constant }
  var useDefault : Bool): Bool stdcall;

{------------------------------------------------------------------------------}
{ Graph Font                                                                   }
{------------------------------------------------------------------------------}
const
  PE_GTF_TITLEFONT       = 0;
  PE_GTF_SUBTITLEFONT    = 1;
  PE_GTF_FOOTNOTEFONT    = 2;
  PE_GTF_GROUPSTITLEFONT = 3;
  PE_GTF_DATATITLEFONT   = 4;
  PE_GTF_LEGENDFONT      = 5;
  PE_GTF_GROUPLABELSFONT = 6;
  PE_GTF_DATALABELSFONT  = 7;

  PE_FACE_NAME_LEN       = 64;

type
  PEFaceNameType = array [0..PE_FACE_NAME_LEN-1] of Char;
  PEFontColorInfo = record
    StructSize   : Word;
    faceName     : PEFaceNameType; { empty string for no change }
    fontFamily   : Smallint; { FF_DONTCARE for no change }
    fontPitch    : Smallint; { DEFAULT_PITCH for no change }
    charSet      : Smallint; { DEFAULT_CHARSET for no change }
    pointSize    : Smallint; { 0 for no change }
    isItalic     : Smallint; { BOOL value, except use PE_UNCHANGED for no change }
    isUnderlined : Smallint; { BOOL value, except use PE_UNCHANGED for no change }
    isStruckOut  : Smallint; { BOOL value, except use PE_UNCHANGED for no change }
    weight       : Smallint; { 0 for no change }
    color        : COLORREF; { PE_UNCHANGED_COLOR for no change }
    twipSize     : Smallint; { Font size in twips, 0 for no change.        }
      { Use either pointSize or twipSize. If both pointSize and twipSize   }
      { are non-zero, twipSize will be used and pointSize will be ignored. }
  end;

const
  PE_SIZEOF_FONT_COLOR_INFO = SizeOf(PEFontColorInfo);

function PEGetGraphFontInfo (
  printJob           : Smallint;
  sectionN           : Smallint;
  graphN             : Smallint;
  titleFontType      : Word;  { PE_GTF_ constant }
  var fontColourInfo : PEFontColorInfo): Bool stdcall;

function PESetGraphFontInfo (
  printJob           : Smallint;
  sectionN           : Smallint;
  graphN             : Smallint;
  titleFontType      : Word; { PE_GTF_ constant }
  var fontColourInfo : PEFontColorInfo): Bool stdcall;

{------------------------------------------------------------------------------}
{ Graph Options                                                                }
{------------------------------------------------------------------------------}
const
  PE_GLP_PLACEUPPERRIGHT    =  0;
  PE_GLP_PLACEBOTTOMCENTER  =  1;
  PE_GLP_PLACETOPCENTER     =  2;
  PE_GLP_PLACERIGHT         =  3;
  PE_GLP_PLACELEFT          =  4;

  { Bar Sizes }
  PE_GBS_MINIMUMBARSIZE     =  0;
  PE_GBS_SMALLBARSIZE       =  1;
  PE_GBS_AVERAGEBARSIZE     =  2;
  PE_GBS_LARGEBARSIZE       =  3;
  PE_GBS_MAXIMUMBARSIZE     =  4;

  { Pie Sizes }
  PE_GPS_MINIMUMPIESIZE     = 64;
  PE_GPS_SMALLPIESIZE       = 48;
  PE_GPS_AVERAGEPIESIZE     = 32;
  PE_GPS_LARGEPIESIZE       = 16;
  PE_GPS_MAXIMUMPIESIZE     =  0;

  { Detached Pie Slice }
  PE_GDPS_NODETACHMENT      =  0;
  PE_GDPS_SMALLESTSLICE     =  1;
  PE_GDPS_LARGESTSLICE      =  2;

  { Marker Sizes }
  PE_GMS_SMALLMARKERS       =  0;
  PE_GMS_MEDIUMSMALLMARKERS =  1;
  PE_GMS_MEDIUMMARKERS      =  2;
  PE_GMS_MEDIUMLARGEMARKERS =  3;
  PE_GMS_LARGEMARKERS       =  4;

  { Marker shapes }
  PE_GMSP_RECTANGLESHAPE    =  1;
  PE_GMSP_CIRCLESHAPE       =  4;
  PE_GMSP_DIAMONDSHAPE      =  5;
  PE_GMSP_TRIANGLESHAPE     =  8;

  { Chart Color }
  PE_GCR_COLORCHART         =  0;
  PE_GCR_BLACKANDWHITECHART =  1;

  { Chart Data points }
  PE_GDP_NONE               =  0;
  PE_GDP_SHOWLABEL          =  1;
  PE_GDP_SHOWVALUE          =  2;

  { Number formats }
  PE_GNF_NODECIMAL          =  0;
  PE_GNF_ONEDECIMAL         =  1;
  PE_GNF_TWODECIMAL         =  2;
  PE_GNF_CURRENCYNODECIMAL  =  3;
  PE_GNF_CURRENCYTWODECIMAL =  4;
  PE_GNF_PERCENTNODECIMAL   =  5;
  PE_GNF_PERCENTONEDECIMAL  =  6;
  PE_GNF_PERCENTTWODECIMAL  =  7;

  { Viewing Angles }
  PE_GVA_STANDARDVIEW       =  1;
  PE_GVA_TALLVIEW           =  2;
  PE_GVA_TOPVIEW            =  3;
  PE_GVA_DISTORTEDVIEW      =  4;
  PE_GVA_SHORTVIEW          =  5;
  PE_GVA_GROUPEYEVIEW       =  6;
  PE_GVA_GROUPEMPHASISVIEW  =  7;
  PE_GVA_FEWSERIESVIEW      =  8;
  PE_GVA_FEWGROUPSVIEW      =  9;
  PE_GVA_DISTORTEDSTDVIEW   = 10;
  PE_GVA_THICKGROUPSVIEW    = 11;
  PE_GVA_SHORTERVIEW        = 12;
  PE_GVA_THICKSERIESVIEW    = 13;
  PE_GVA_THICKSTDVIEW       = 14;
  PE_GVA_BIRDSEYEVIEW       = 15;
  PE_GVA_MAXVIEW            = 16;

  {legend layout}
  PE_GLL_PERCENTAGE         =  0;
  PE_GLL_AMOUNT             =  1;
  PE_GLL_CUSTOM             =  2;  { for PEGetGraphOptionInfo, do not use }
                                   { in PESetGraphOptionInfo. }
type
  PEGraphOptionInfo = record
    StructSize       : Word;
    graphColour      : Smallint; { PE_GCR_*, PE_UNCHANGED for no change }
    showLegend       : Smallint; { BOOL, PE_UNCHANGED for no change }
    legendPosition   : Smallint; { PE_GLP_*, if showLegend == 0, means no legend }
    { Pie Charts and Doughut Charts }
    pieSize          : Smallint; { PE_GPS_*, PE_UNCHANGED for no change }
    detachedPieSlice : Smallint; { PE_GDPS_*, PE_UNCHANGED for no change }
    { Bar Chart }
    barSize          : Smallint; { PE_GBS_*, PE_UNCHANGED for no change }
    verticalBars     : Smallint; { BOOL, PE_UNCHANGED for no change }
    { Markers (used for line and bar charts) }
    markerSize       : Smallint; { PE_GMS_*, PE_UNCHANGED for no change }
    markerShape      : Smallint; { PE_GMSP_*, PE_UNCHANGED for no change }
    { Data Points }
    dataPoints       : Smallint; { PE_GDP_*, PE_UNCHANGED for no change }
    dataValueNumberFormat : Smallint; { PE_GNF_*, PE_UNCHANGED for no change }
    { 3D }
    viewingAngle     : Smallint; { PE_GVA_*, PE_UNCHANGED for no change }
    legendLayout     : Smallint; { PE_GLL_ constant }
  end;

const
  PE_SIZEOF_GRAPH_OPTION_INFO = SizeOf(PEGraphOptionInfo);

function PEGetGraphOptionInfo (
  printJob            : Smallint;
  sectionN            : Smallint;
  graphN              : Smallint;
  var graphOptionInfo : PEGraphOptionInfo): Bool stdcall;

function PESetGraphOptionInfo (
  printJob           : Smallint;
  sectionN           : Smallint;
  graphN             : Smallint;
  var graphOptionInfo : PEGraphOptionInfo): Bool stdcall;

{------------------------------------------------------------------------------}
{ Graph Axes                                                                   }
{------------------------------------------------------------------------------}
const
  PE_GGT_NOGRIDLINES            =  0;
  PE_GGT_MINORGRIDLINES         =  1;
  PE_GGT_MAJORGRIDLINES         =  2;
  PE_GGT_MAJORANDMINORGRIDLINES =  3;
  PE_ADM_AUTOMATIC              =  0;
  PE_ADM_MANUAL                 =  1;

type
  PEGraphAxisInfo = record
    StructSize         : Word;
    { Grid Lines }
    groupAxisGridLine  : Smallint; { PE_GGT_*, PE_UNCHANGED for no change }
    dataAxisYGridLine  : Smallint; { PE_GGT_*, PE_UNCHANGED for no change }
    dataAxisY2GridLine : Smallint; { PE_GGT_*, PE_UNCHANGED for no change }
    seriesAxisGridline : Smallint; { PE_GGT_*, PE_UNCHANGED for no change }
    { Min/Max Values }
    dataAxisYMinValue  : double;
    dataAxisYMaxValue  : double;
    dataAxisY2MinValue : double;
    dataAxisY2MaxValue : double;
    seriesAxisMinValue : double;
    seriesAxisMaxValue : double;
    { Number Format }
    dataAxisYNumberFormat  : Smallint; { PE_GNF_*, PE_UNCHANGED for no change }
    dataAxisY2NumberFormat : Smallint; { PE_GNF_*, PE_UNCHANGED for no change }
    seriesAxisNumberFormat : Smallint; { PE_GNF_*, PE_UNCHANGED for no change }
    { Auto Range }
    dataAxisYAutoRange  : Smallint; { BOOL, PE_UNCHANGED for no change }
    dataAxisY2AutoRange : Smallint; { BOOL, PE_UNCHANGED for no change }
    seriesAxisAutoRange : Smallint; { BOOL, PE_UNCHANGED for no change }
    { Automatic Division }
    dataAxisYAutomaticDivision  : Smallint; { PE_ADM_*, PE_UNCHANGED for no change }
    dataAxisY2AutomaticDivision : Smallint; { PE_ADM_*, PE_UNCHANED for no change }
    seriesAxisAutomaticDivision : Smallint; { PE_ADM_*, PE_UNCHANED for no change }
    { Manual Division }
    dataAxisYManualDivision  : LongInt; { if dataAxisYAutomaticDivision is PE_ADM_AUTOMATIC, this field is ignored }
    dataAxisY2ManualDivision : LongInt; { if dataAxisY2AutomaticDivision is PE_ADM_AUTOMATIC, this field is ignored }
    seriesAxisManualDivision : LongInt; { if seriesAxisAutomaticDivision is PE_ADM_AUTOMATIC, this field is ignored }
    { Auto Scale }
    dataAxisYAutoScale  : Smallint;  { BOOL, PE_UNCHANGED for no change }
    dataAxisY2AutoScale : Smallint;  { BOOL, PE_UNCHANGED for no change }
    seriesAxisAutoScale : Smallint;  { BOOL, PE_UNCHANGED for no change }
  end;

const
  PE_SIZEOF_GRAPH_AXIS_INFO = SizeOf(PEGraphAxisInfo);

function PEGetGraphAxisInfo (
  printJob          : Smallint;
  sectionN          : Smallint;
  graphN            : Smallint;
  var graphAxisInfo : PEGraphAxisInfo): Bool stdcall;

function PESetGraphAxisInfo (
  printJob          : Smallint;
  sectionN          : Smallint;
  graphN            : Smallint;
  var graphAxisInfo : PEGraphAxisInfo): Bool stdcall;


{******************************************************************************}
{  Formulas, Selection Formulas and Group Selection Formulas                   }
{******************************************************************************}
{------------------------------------------------------------------------------}
{ Formulas                                                                     }
{------------------------------------------------------------------------------}
function PEGetNFormulas (printJob : Smallint): Smallint stdcall;

function PEGetNthFormula (
  printJob       : Smallint;
  formulaN       : Smallint;
  var nameHandle : Hwnd;
  var nameLength : Smallint;
  var textHandle : Hwnd;
  var textLength : Smallint): Bool stdcall;

function PEGetFormula (
  printJob       : Smallint;
  formulaName    : PChar;
  var textHandle : HWnd;
  var textLength : Smallint): Bool stdcall;

function PESetFormula (
  printJob      : Smallint;
  formulaName   : PChar;
  formulaString : PChar): Bool stdcall;

function PECheckFormula (
  printJob    : Smallint;
  formulaName : PChar): Bool stdcall;

{------------------------------------------------------------------------------}
{ Selection Formula                                                            }
{------------------------------------------------------------------------------}
function PEGetSelectionFormula (
  printJob       : Smallint;
  var textHandle : HWnd;
  var textLength : Smallint): Bool stdcall;

function PESetSelectionFormula (
  printJob      : Smallint;
  formulaString : PChar): Bool stdcall;

function PECheckSelectionFormula (printJob : Smallint): Bool stdcall;

{------------------------------------------------------------------------------}
{ Group Selection Formula                                                      }
{------------------------------------------------------------------------------}
function PEGetGroupSelectionFormula (
  printJob       : Smallint;
  var textHandle : HWnd;
  var textLength : Smallint): Bool stdcall;

function PESetGroupSelectionFormula (
  printJob      : Smallint;
  formulaString : PChar): Bool stdcall;

function PECheckGroupSelectionFormula (printJob : Smallint): Bool stdcall;

{------------------------------------------------------------------------------}
{ Section Format Formulas                                                      }
{------------------------------------------------------------------------------}
const
  { Format formula name : Old naming convention }
  SECTION_VISIBILITY      = 58;
  NEW_PAGE_BEFORE         = 60;
  NEW_PAGE_AFTER          = 61;
  KEEP_SECTION_TOGETHER   = 62;
  SUPPRESS_BLANK_SECTION  = 63;
  RESET_PAGE_N_AFTER      = 64;
  PRINT_AT_BOTTOM_OF_PAGE = 65;
  UNDERLAY_SECTION        = 66;
  SECTION_BACK_COLOUR     = 67;
  { Format formula name : New naming convention}
  PE_FFN_AREASECTION_VISIBILITY  = 58;
  PE_FFN_SECTION_VISIBILITY      = 58;
  PE_FFN_SHOW_AREA               = 59;
  PE_FFN_NEW_PAGE_BEFORE         = 60;
  PE_FFN_NEW_PAGE_AFTER          = 61;
  PE_FFN_KEEP_SECTION_TOGETHER   = 62;
  PE_FFN_KEEP_TOGETHER           = 62;
  PE_FFN_SUPPRESS_BLANK_SECTION  = 63;
  PE_FFN_RESET_PAGE_N_AFTER      = 64;
  PE_FFN_PRINT_AT_BOTTOM_OF_PAGE = 65;
  PE_FFN_UNDERLAY_SECTION        = 66;
  PE_FFN_SECTION_BACK_COLOR      = 67;

function PEGetSectionFormatFormula (
  printJob       : Smallint;
  sectionCode    : Smallint;
  formulaName    : Smallint;
  var textHandle : HWnd;
  var textLength : Smallint): Bool stdcall;

function PESetSectionFormatFormula (
  printJob      : Smallint;
  sectionCode   : Smallint;
  formulaName   : Smallint;
  formulaString : PChar): Bool stdcall;

{------------------------------------------------------------------------------}
{ Area Format Formulas                                                         }
{------------------------------------------------------------------------------}
function PEGetAreaFormatFormula (
  printJob       : Smallint;
  areaCode       : Smallint;
  formulaName    : Smallint;
  var textHandle : HWnd;
  var textLength : Smallint): Bool stdcall;

function PESetAreaFormatFormula (
  printJob      : Smallint;
  areaCode      : Smallint;
  formulaName   : Smallint;
  formulaString : PChar): Bool stdcall;


{******************************************************************************}
{ SQL Expressions                                                              }
{******************************************************************************}
function PEGetNSQLExpressions (printJob : Smallint): Smallint stdcall;

function PEGetNthSQLExpression (
  printJob       : Smallint;
  expressionN    : Smallint;
  var nameHandle : Hwnd;
  var nameLength : Smallint;
  var textHandle : Hwnd;
  var textLength : Smallint): Bool stdcall;

function PEGetSQLExpression (
  printJob             : Smallint;
  const expressionName : PChar;
  var textHandle       : Hwnd;
  var textLength       : Smallint): Bool stdcall;

function PESetSQLExpression (
  printJob               : Smallint;
  const expressionName   : PChar;
  const expressionString : PChar): Bool stdcall;

function PECheckSQLExpression (
  printJob             : Smallint;
  const expressionName : PChar): Bool stdcall;


(********************************************************************************/
** NOTE : Stored Procedures
**   The previous Stored Procedure API Calls PEGetNParams, PEGetNthParam,
** PEGetNthParamInfo and PESetNthParam have been made obsolete.  Older
** applications that used these API Calls will still work as before, but for new
** development please use the new Parameter API calls below,
**   The Stored Procedure Parameters have now been unified with the Parameter
** Fields.
**
** The replacements for these calls are as follows :
**		PEGetNParams	  = PEGetNParameterFields
**		PEGetNthParam	  = PEGetNthParameterField
**		PEGetNthParamInfo = PEGetParameterValueInfo
**		PESetNthParam	  = PESetNthParameterField
**
** NOTE : To tell if a Parameter Field is a Stored Procedure, use the
**	  PEGetNthParameterType or PEGetNthParameterField API Calls
**
** If you wish to SET a parameter to NULL then set the CurrentValue to CRWNULL.
** The CRWNULL is of Type String and is independant of the datatype of the
** parameter.
*******************************************************************************)


{******************************************************************************}
{ Parameter Fields                                                             }
{******************************************************************************}
{------------------------------------------------------------------------------}
{ Getting/Setting Parameter Fields                                             }
{------------------------------------------------------------------------------}
const
  PE_PF_REPORT_NAME_LEN = 128;
  PE_PF_NAME_LEN        = 256;
  PE_PF_PROMPT_LEN      = 256;
  PE_PF_VALUE_LEN       = 256;
  PE_PF_EDITMASK_LEN    = 256;

  PE_PF_NUMBER     = 0;
  PE_PF_CURRENCY   = 1;
  PE_PF_BOOLEAN    = 2;
  PE_PF_DATE       = 3;
  PE_PF_STRING     = 4;
  PE_PF_DATETIME   = 5;
  PE_PF_TIME       = 6;

type
  PE_PF_ReportNameType         = array [0..PE_PF_REPORT_NAME_LEN-1] of Char;
  PEParameterFieldNameType     = array [0..PE_PF_NAME_LEN-1] of Char;
  PEParameterFieldTextType     = array [0..PE_PF_PROMPT_LEN-1] of Char;
  PEParameterFieldValueType    = array [0..PE_PF_VALUE_LEN-1] of Char;
  PEParameterFieldEditMaskType = array [0..PE_PF_EDITMASK_LEN-1] of Char;
  PEParameterFieldInfo = record
    structSize        : Word; { Initialize to PE_SIZEOF_PARAMETER_FIELD_INFO }
    ValueType         : Word; { PE_PF_ constant }
    DefaultValueSet   : Word; { Indicate the default value is set in PEParameterFieldInfo }
    CurrentValueSet   : Word; { Indicate the current value is set in PEParameterFieldInfo }
    Name              : PEParameterFieldNameType;
    Prompt            : PEParameterFieldTextType;
    { Next 2 Could be Number, Date, DateTime, Time, Boolean, or String }
    DefaultValue      : PEParameterFieldValueType;
    CurrentValue      : PEParameterFieldValueType;
    { Name of report where the field belongs, only used in PEGetNthParameterField }
    ReportName        : PE_PF_ReportNameType;
    { Returns false (0) if parameter is linked, not in use, or has current value set }
    needsCurrentValue : Word;
    { For String values this will be TRUE if the string is limited on length, for }
    { other types it will be TRUE if the parameter is limited by a range }
    isLimited         : Word;
    { For string fields, these are the minimum/maximum length of the string }
    { For numeric fields, they are the minimum/maximum numeric value }
    { For other fields, use PEGetParameterMinMaxValue }
    MinSize           : double;
    MaxSize           : double;
    { An edit mask that restricts what may be entered for string parameters }
    EditMask          : PEParameterFieldEditMaskType;
    { Returns True if it is essbase sub var }
    isHidden          : Word;
  end;

const
  PE_SIZEOF_PARAMETER_FIELD_INFO = SizeOf(PEParameterFieldInfo);
  PE_SIZEOF_VARINFO_TYPE = SizeOf(PEParameterFieldInfo);

function PEGetNParameterFields (printJob: Smallint): Smallint stdcall;

function PEGetNthParameterField (
  printJob          : Smallint;
  parameterN        : Smallint;
  var parameterInfo : PEParameterFieldInfo): Bool stdcall;

function PESetNthParameterField (
  printJob          : Smallint;
  parameterN        : Smallint;
  var parameterInfo : PEParameterFieldInfo): Bool stdcall;

{------------------------------------------------------------------------------}
{ Converting Default/Current Value via PEValueInfo                             }
{------------------------------------------------------------------------------}
const
  { ValueInfo types }
  PE_VI_NUMBER      = 0;
  PE_VI_CURRENCY    = 1;
  PE_VI_BOOLEAN     = 2;
  PE_VI_DATE        = 3;
  PE_VI_STRING      = 4;
  PE_VI_DATETIME    = 5;
  PE_VI_TIME        = 6;
  PE_VI_INTEGER     = 7;
  PE_VI_COLOR       = 8;
  PE_VI_CHAR        = 9;
  PE_VI_LONG        = 10;
  PE_VI_NOVALUE    = 100;
  PE_VI_STRING_LEN = 256;

type
  PEValueInfoStringType     = array[0..PE_VI_STRING_LEN-1] of Char;
  PEValueInfoDateOrTimeType = array[0..2] of Smallint;
  PEValueInfoDateTimeType   = array[0..5] of Smallint;
  PEValueInfo = record
    StructSize : Word;
    valueType  : Word;  {a PE_VI_ constant}
    viNumber   : Double;
    viCurrency : Double;
    viBoolean  : Bool;
    viString   : PEValueInfoStringType;
    viDate     : PEValueInfoDateOrTimeType; {year, month, day}
    viDateTime : PEValueInfoDateTimeType;   {year, month, day, hour, minute, second}
    viTime     : PEValueInfoDateOrTimeType; {hour, minute, second}
    viColor    : DWord;
    viInteger  : Smallint;
    viC        : Char;
    ignored    : Char; {for 4 byte alignment. ignored.}
    viLong     : DWord;
  end;

const
  PE_SIZEOF_VALUE_INFO = SizeOf(PEValueInfo);

function PEConvertPFInfoToVInfo (
  value         : PChar;
  valueType     : Smallint;
  var valueInfo : PEValueInfo): Bool stdcall;

function PEConvertVInfoToPFInfo (
  var valueInfo  : PEValueInfo;
  var valueType  : Word;
  value : PChar): Bool stdcall;

{------------------------------------------------------------------------------}
{ Getting/Setting multiple Default Values                                      }
{------------------------------------------------------------------------------}
{ If return value is -1 then an error has occurred. }
function PEGetNParameterDefaultValues(
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar): Smallint stdcall;

function PEGetNthParameterDefaultValue (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  index                    : Smallint;
  var valueInfo            : PEValueInfo): Bool stdcall;

function PESetNthParameterDefaultValue (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  index                    : Smallint;
  var valueInfo            : PEValueInfo): Bool stdcall;

function PEAddParameterDefaultValue (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  var valueInfo            : PEValueInfo): Bool stdcall;

function PEDeleteNthParameterDefaultValue (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  index                    : Smallint): Bool stdcall;

{------------------------------------------------------------------------------}
{ Min/Max Values for Parameter Fields                                          }
{------------------------------------------------------------------------------}
function PEGetParameterMinMaxValue (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  var valueMin             : PEValueInfo;
  { Set to NULL to retrieve MAX only; must be non-NULL if valueMax is NULL }
  var valueMax             : PEValueInfo): Bool stdcall;
  { Set to NULL to retrieve MIN only; must be non-NULL if valueMin is NULL }


function PESetParameterMinMaxValue (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  var valueMin             : PEValueInfo;
  { Set to NULL to retrieve MAX only; must be non-NULL if valueMax is NULL }
  var valueMax             : PEValueInfo): Bool stdcall;
  { Set to NULL to retrieve MIN only; must be non-NULL if valueMin is NULL }
  { If both valueInfo and valueMax are non-NULL then }
  { valueMin->valueType MUST BE THE SAME AS valueMax->valueType. }
  { If different, PE_ERR_INCONSISTANTTYPES is returned. }

{------------------------------------------------------------------------------}
{ Pick List Options in Parameter Fields                                        }
{------------------------------------------------------------------------------}
function PEGetNthParameterValueDescription (
  printJob            : Smallint;
  parameterFieldName  : PChar;
  reportName          : PChar;
  index               : Smallint;
  var valueDesc       : HWnd;
  var valueDescLength : Smallint): Bool stdcall;

function PESetNthParameterValueDescription (
  printJob            : Smallint;
  parameterFieldName  : PChar;
  reportName          : PChar;
  index               : Smallint;
  valueDesc           : PChar): Bool stdcall;

{ Constants for sortMethod in PEParameterPickListOption }
const
  PE_OR_NO_SORT                 = 0;
  PE_OR_ALPHANUMERIC_ASCENDING  = 1;
  PE_OR_ALPHANUMERIC_DESCENDING = 2;
  PE_OR_NUMERIC_ASCENDING       = 3;
  PE_OR_NUMERIC_DESCENDING      = 4;

type
  PEParameterPickListOption = record
    StructSize      : Word;      { initialize to PE_SIZEOF_PICK_LIST_OPTION }
    showDescOnly    : Smallint;  { boolean value or PE_UNCHANGED }
    sortMethod      : Smallint;  { enum type const, PE_UNCHANGED for no change }
    sortBasedOnDesc : Smallint;  { boolean value or PE_UNCHANGED }
  end;

const
  PE_SIZEOF_PICK_LIST_OPTION = SizeOf(PEParameterPickListOption);

function PEGetParameterPickListOption (
  printJob           : Smallint;
  parameterFieldName : PChar;
  reportName         : PChar;
  var pickListOption : PEParameterPickListOption): Bool stdcall;

function PESetParameterPickListOption (
  printJob           : Smallint;
  parameterFieldName : PChar;
  reportName         : PChar;
  var pickListOption : PEParameterPickListOption): Bool stdcall;

{------------------------------------------------------------------------------}
{ Parameter Value Info - extra options for Parameter Fields                    }
{------------------------------------------------------------------------------}
type
  PEParameterValueInfo = record
    StructSize             : Word;
    isNullable             : Smallint; {Boolean value or PE_UNCHANGED for no change.}
    disallowEditing        : Smallint; {Boolean value or PE_UNCHANGED for no change.}
    allowMultipleValues    : Smallint; {Boolean value or PE_UNCHANGED for no change.}
    hasDiscreteValues      : Smallint; {Boolean value or PE_UNCHANGED for no change.}
                                       {True: has discrete values, False: has ranges}
    partOfGroup            : Smallint; {Boolean value or PE_UNCHANGED for no change.}
    groupNum               : Smallint; {a group number or PE_UNCHANGED for no change.}
    mutuallyExclusiveGroup : Smallint; {Boolean value or PE_UNCHANGED for no change.}
  end;

const
  PE_SIZEOF_PARAMETER_VALUE_INFO = SizeOf(PEParameterValueInfo);

function PEGetParameterValueInfo (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  var valueInfo            : PEParameterValueInfo): Bool stdcall;

function PESetParameterValueInfo (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  var valueInfo            : PEParameterValueInfo): Bool stdcall;

{------------------------------------------------------------------------------}
{ Getting/Setting multiple Current Values and Ranges                           }
{------------------------------------------------------------------------------}
const
  { Range Info }
  PE_RI_INCLUDEUPPERBOUND = 1;
  PE_RI_INCLUDELOWERBOUND = 2;
  PE_RI_NOUPPERBOUND      = 4;
  PE_RI_NOLOWERBOUND      = 8;

  PE_DR_HASRANGE            = 0;
  PE_DR_HASDISCRETE         = 1;
  PE_DR_HASDISCRETEANDRANGE = 2;

{ If return value is -1 then an error has occurred. }
function PEGetNParameterCurrentValues (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar): Word stdcall;

function PEGetNthParameterCurrentValue (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  index                    : Smallint;
  var currentValue         : PEValueInfo): Bool stdcall;

function PEAddParameterCurrentValue (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  var currentValue         : PEValueInfo): Bool stdcall;

{ If return value is -1 then an error has occurred. }
function PEGetNParameterCurrentRanges (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar) : Word stdcall;

function PEGetNthParameterCurrentRange (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  index                    : Smallint;
  var rangeStart           : PEValueInfo;
  var rangeEnd             : PEValueInfo;
  var rangeInfo            : Smallint): Bool stdcall;

function PEAddParameterCurrentRange (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar;
  var rangeStart           : PEValueInfo;
  var rangeEnd             : PEValueInfo;
  rangeInfo                : Smallint): Bool stdcall;

function PEClearParameterCurrentValuesAndRanges (
  printJob                 : Smallint;
  const parameterFieldName : PChar;
  const reportName         : PChar): Bool stdcall;

{------------------------------------------------------------------------------}
{ Parameter Field Type                                                         }
{------------------------------------------------------------------------------}
const
  { Parameter Field origin }
  PE_PO_REPORT     = 0;
  PE_PO_STOREDPROC = 1;
  PE_PO_QUERY      = 2;

{returns PE_PO_* or -1 if index is invalid.}
function PEGetNthParameterType (
  printJob : Smallint;
  index    : Smallint): Smallint stdcall;


{******************************************************************************}
{ Record and Group SortFields                                                  }
{******************************************************************************}
{------------------------------------------------------------------------------}
{ SortFields                                                                   }
{------------------------------------------------------------------------------}
function PEGetNSortFields (printJob: Smallint): Smallint stdcall;

function PEGetNthSortField (
  printJob       : Smallint;
  sortFieldN     : Smallint;
  var nameHandle : HWnd;
  var nameLength : Smallint;
  var direction  : Smallint): Bool stdcall;

function PESetNthSortField (
  printJob   : Smallint;
  sortFieldN : Smallint;
  Name       : PChar;
  direction  : Smallint): Bool stdcall;

function PEDeleteNthSortField (
  printJob   : Smallint;
  sortFieldN : Smallint): Bool stdcall;

{------------------------------------------------------------------------------}
{ GroupSortFields                                                              }
{------------------------------------------------------------------------------}
function PEGetNGroupSortFields (printJob : Smallint): Smallint stdcall;

function PEGetNthGroupSortField (
  printJob       : Smallint;
  sortFieldN     : Smallint;
  var nameHandle : HWnd;
  var nameLength : Smallint;
  var direction  : Smallint): Bool stdcall;

function PESetNthGroupSortField (
  printJob   : Smallint;
  sortFieldN : Smallint;
  Name       : PChar;
  direction  : Smallint): Bool stdcall;

function PEDeleteNthGroupSortField (
  printJob   : Smallint;
  sortFieldN : Smallint): Bool stdcall;


{******************************************************************************}
{ Controlling Databases                                                        }
{   The following functions allow retrieving and updating database info        }
{   in an opened report, so that a report can be printed using different       }
{   session, server, database, user and/or table location settings.  Any       }
{   changes made to the report via these functions are not permanent, and      }
{   only last as long as the report is open.                                   }
{                                                                              }
{   The following database functions (except for PELogOnServer and             }
{   PELogOffServer) must be called after PEOpenPrintJob and before             }
{   PEStartPrintJob.                                                           }
{******************************************************************************}
{------------------------------------------------------------------------------}
{ Getting Number of Tables                                                     }
{   The function PEGetNTables is called to fetch the number of tables in       }
{   the report.  This includes all PC databases (e.g. Paradox, xBase)          }
{   as well as SQL databases (e.g. SQL Server, Oracle, Netware).               }
{------------------------------------------------------------------------------}
function PEGetNTables (printJob: Smallint): Smallint stdcall;

{------------------------------------------------------------------------------}
{ Getting Table Type                                                           }
{   The function PEGetNthTableType allows the application to determine the     }
{   type of each table.  The application can test DBType (equal to             }
{   PE_DT_STANDARD or PE_DT_SQL), or test the database DLL name used to        }
{   create the report.  DLL names have the following naming convention:        }
{       - PDB*.DLL for standard (non-SQL) databases,                           }
{       - PDS*.DLL for SQL databases.                                          }
{                                                                              }
{   In the case of ODBC (PDSODBC.DLL) the DescriptiveName includes the         }
{   ODBC data source name.                                                     }
{------------------------------------------------------------------------------}
const
  PE_DLL_NAME_LEN            =  64;
  PE_FULL_NAME_LEN           = 256;
  { DBType constants }
  PE_DT_STANDARD             = 1;
  PE_DT_SQL                  = 2;
  PE_DT_SQL_STORED_PROCEDURE = 3;

type
  PEDllNameType  = array [0..PE_DLL_NAME_LEN - 1] of Char;
  PEFullNameType = array [0..PE_FULL_NAME_LEN - 1] of Char;
  PETableType = record
    StructSize      : Word; { Initialize to PE_SIZEOF_TABLE_TYPE }
    DLLName         : PEDllNameType;
    DescriptiveName : PEFullNameType;
    DBType          : Word;
  end;

const
  PE_SIZEOF_TABLE_TYPE = SizeOf(PETableType);

function PEGetNthTableType (
  printJob      : Smallint;
  tableN        : Smallint;
  var tableType : PETableType): Bool stdcall;

{------------------------------------------------------------------------------}
{ Getting Table Session Info                                                   }
{   The functions PEGetNthTableSessionInfo and PESetNthTableSessionInfo        }
{   are only used when connecting to MS Access databases (which require a      }
{   session to be opened first)                                                }
{------------------------------------------------------------------------------}
const
  PE_SESS_USERID_LEN   = 128;
  PE_SESS_PASSWORD_LEN = 128;

type
  PESesPassType = array [0..PE_SESS_PASSWORD_LEN - 1] of Char;
  PESessUserType = array [0..0PE_SESS_USERID_LEN - 1] of Char;
  PESessionInfo = record
    StructSize    : Word; { Initialize to PE_SIZEOF_SESSION_INFO }
    UserID        : PESesUserType;
    { Password is undefined when getting information from report. }
    Password      : PESesPassType;
    { SessionHandle is undefined when getting information from report. }
    { When setting information, if it is = 0 the UserID and Password }
    { settings are used, otherwise the SessionHandle is used. }
    SessionHandle : DWord;
  end;

const
  PE_SIZEOF_SESSION_INFO = SizeOf(PESessionInfo);

function PEGetNthTableSessionInfo (
  printJob        : Smallint;
  tableN          : Smallint;
  var sessionInfo : PESessionInfo): Bool stdcall;

function PESetNthTableSessionInfo (
  printJob              : Smallint;
  tableN                : Smallint;
  var sessionInfo       : PESessionInfo;
  propagateAcrossTables : Bool): Bool stdcall;

{------------------------------------------------------------------------------}
{ Table LogOn/LogOff                                                           }
{   Logging on is performed when printing the report, but the correct          }
{   log on information must first be set using PESetNthTableLogOnInfo.         }
{   Only the password is required, but the server, database, and               }
{   user names may optionally be overriden as well.                            }
{                                                                              }
{   If the parameter propagateAcrossTables is TRUE, the new log on info        }
{   is also applied to any other tables in this report that had the            }
{   same original server and database names as this table.  If FALSE           }
{   only this table is updated.                                                }
{                                                                              }
{   Logging off is performed automatically when the print job is closed.       }
{------------------------------------------------------------------------------}
const
  PE_SERVERNAME_LEN      = 128;
  PE_DATABASENAME_LEN    = 128;
  PE_USERID_LEN          = 128;
  PE_PASSWORD_LEN        = 128;

type
  PELogonServerType = array [0..PE_SERVERNAME_LEN - 1] of Char;
  PELogonDbType     = array [0..PE_DATABASENAME_LEN - 1] of Char;
  PELogonUserType   = array [0..PE_USERID_LEN - 1] of Char;
  PELogonPassType   = array [0..PE_PASSWORD_LEN - 1] of Char;
  PELogOnInfo = record
    StructSize   : Word; { Initialize to PE_SIZEOF_LOGON_INFO }
    (* For any of the following values an empty string ("") means to use
    ** the value already set in the report.  To override a value in the
    ** report use a non-empty string (e.g. "Server A").  All strings are
    ** null-terminated.
    **
    ** For Netware SQL, pass the dictionary path name in ServerName and
    ** data path name in DatabaseName. *)
    ServerName   : PELogonServerType;
    DatabaseName : PELogonDbType;
    UserId       : PELogonUserType;
    (* Password is undefined when getting information from report. *)
    Password     : PELogonPassType;
  end;

const
  PE_SIZEOF_LOGON_INFO = SizeOf(PELogOnInfo);

function PEGetNthTableLogOnInfo (
  printJob      : Smallint;
  tableN        : Smallint;
  var logOnInfo : PELogOnInfo): Bool stdcall;

function PESetNthTableLogOnInfo (
  printJob              : Smallint;
  tableN                : Smallint;
  var logOnInfo         : PELogOnInfo;
  propagateAcrossTables : Bool): Bool stdcall;

{------------------------------------------------------------------------------}
{ Table Location                                                               }
{   A table's location is fetched and set using PEGetNthTableLocation and      }
{   PESetNthTableLocation.  This name is database-dependent, and must be       }
{   formatted correctly for the expected database.  For example:               }
{       - Paradox: "c:\crw\ORDERS.DB"                                          }
{       - SQL Server: "publications.dbo.authors"                               }
{------------------------------------------------------------------------------}
const
  PE_TABLE_LOCATION_LEN    = 256;
  PE_CONNECTION_BUFFER_LEN = 512;

type
  PETableLocType = array [0..PE_TABLE_LOCATION_LEN - 1] of Char;
  PEConnectBufferType = array [0..PE_CONNECTION_BUFFER_LEN - 1] of Char;
  PETableLocation = record
    StructSize    : Word; { Initialize to PE_SIZEOF_TABLE_LOCATION }
    Location      : PETableLocType;
    SubLocation   : PETableLocType; { For MS Access Table Names }
    ConnectBuffer : PEConnectBufferType;
  end;

const
  PE_SIZEOF_TABLE_LOCATION = SizeOf(PETableLocation);

function PEGetNthTableLocation (
  printJob     : Smallint;
  tableN       : Smallint;
  var location : PETableLocation): Bool stdcall;

function PESetNthTableLocation (
  printJob     : Smallint;
  tableN       : Smallint;
  var location : PETableLocation): Bool stdcall;

{------------------------------------------------------------------------------}
{ Table Private Info - for CDO, ADO, etc.                                      }
{------------------------------------------------------------------------------}
type
  crBytePointer = ^Byte;
  PETablePrivateInfo = record
    StructSize : Word;  { initialize to PE_SIZEOF_TABLE_PRIVATE_INFO }
    nBytes     : Smallint;
    tag        : DWord;
    dataPtr    : crBytePointer;
  end;

const
  PE_SIZEOF_TABLE_PRIVATE_INFO = SizeOf(PETablePrivateInfo);

function PEGetNthTablePrivateInfo (
  printJob        : Smallint;
  tableN          : Smallint;
  var privateInfo : PETablePrivateInfo): Bool stdcall;

function PESetNthTablePrivateInfo (
  printJob        : Smallint;
  tableN          : Smallint;
  var privateInfo : PETablePrivateInfo): Bool stdcall;

{------------------------------------------------------------------------------}
{ Test Connectivity                                                            }
{   The function PETestNthTableConnectivity tests whether a database           }
{   table's settings are valid and ready to be reported on.  It returns        }
{   true if the database session, log on, and location info is all             }
{   correct.                                                                   }
{                                                                              }
{   This is useful, for example, in prompting the user and testing a           }
{   server password before printing begins.                                    }
{                                                                              }
{   This function may require a significant amount of time to complete,        }
{   since it will first open a user session (if required), then log onto       }
{   the database server (if required), and then open the appropriate           }
{   database table (to test that it exists).  It does not read any data,       }
{   and closes the table immediately once successful.  Logging off is          }
{   performed when the print job is closed.                                    }
{                                                                              }
{   If it fails in any of these steps, the error code set indicates            }
{   which database info needs to be updated using functions above:             }
{      - If it is unable to begin a session, PE_ERR_DATABASESESSION is set,    }
{        and the application should update with PESetNthTableSessionInfo.      }
{      - If it is unable to log onto a server, PE_ERR_DATABASELOGON is set,    }
{        and the application should update with PESetNthTableLogOnInfo.        }
{      - If it is unable open the table, PE_ERR_DATABASELOCATION is set,       }
{        and the application should update with PESetNthTableLocation.         }
{------------------------------------------------------------------------------}
function PETestNthTableConnectivity (
  printJob : Smallint;
  tableN   : Smallint): Bool stdcall;

{------------------------------------------------------------------------------}
{ Logging On/Off Server                                                        }
{   PELogOnServer and PELogOffServer can be called at any time to log on       }
{   and off of a database server.  These functions are not required if         }
{   function PESetNthTableLogOnInfo above was already used to set the          }
{   password for a table.                                                      }
{                                                                              }
{   These functions require a database DLL name, which can be retrieved        }
{   using PEGetNthTableType above.                                             }
{                                                                              }
{   This function can also be used for non-SQL tables, such as password-       }
{   protected Paradox tables.  Call this function to set the password          }
{   for the Paradox DLL before beginning printing.                             }
{                                                                              }
{   Note: When printing using PEStartPrintJob the ServerName passed in         }
{   PELogOnServer must agree exactly with the server name stored in the        }
{   report.  If this is not true use PESetNthTableLogOnInfo to perform         }
{   logging on instead.                                                        }
{------------------------------------------------------------------------------}
function PELogOnServer (
  dllName       : PChar;
  var logOnInfo : PELogOnInfo): Bool stdcall;

function PELogOffServer (
  dllName       : PChar;
  var logOnInfo : PELogOnInfo): Bool stdcall;

function PELogOnSQLServerWithPrivateInfo (
  dllName     : PChar;
  privateInfo : Pointer): Bool stdcall;

{------------------------------------------------------------------------------}
{ Verify Database                                                              }
{------------------------------------------------------------------------------}
function PEVerifyDatabase(printJob : Smallint): Bool stdcall;

{------------------------------------------------------------------------------}
{ Check Table Differences                                                      }
{------------------------------------------------------------------------------}
const
  { Constants returned from PECheckNthTableDifferences, }
  { can be any combination of the following: }
  PE_TCD_OKAY                     = $00000000;
  PE_TCD_DATABASENOTFOUND         = $00000001;
  PE_TCD_SERVERNOTFOUND           = $00000002;
  PE_TCD_SERVERNOTOPENED          = $00000004;
  PE_TCD_ALIASCHANGED             = $00000008;
  PE_TCD_INDEXESCHANGED           = $00000010;
  PE_TCD_DRIVERCHANGED            = $00000020;
  PE_TCD_DICTIONARYCHANGED        = $00000040;
  PE_TCD_FILETYPECHANGED          = $00000080;
  PE_TCD_RECORDSIZECHANGED        = $00000100;
  PE_TCD_ACCESSCHANGED            = $00000200;
  PE_TCD_PARAMETERSCHANGED        = $00000400;
  PE_TCD_LOCATIONCHANGED          = $00000800;
  PE_TCD_DATABASEOTHER            = $00001000;
  PE_TCD_NUMFIELDSCHANGED         = $00010000;
  PE_TCD_FIELDOTHER               = $00020000;
  PE_TCD_FIELDNAMECHANGED         = $00040000;
  PE_TCD_FIELDDESCCHANGED         = $00080000;
  PE_TCD_FIELDTYPECHANGED         = $00100000;
  PE_TCD_FIELDSIZECHANGED         = $00200000;
  PE_TCD_NATIVEFIELDTYPECHANGED   = $00400000;
  PE_TCD_NATIVEFIELDOFFSETCHANGED = $00800000;
  PE_TCD_NATIVEFIELDSIZECHANGED   = $01000000;
  PE_TCD_FIELDDECPLACESCHANGED    = $02000000;

type
  PETableDifferenceInfo = record
    StructSize       : Word;
    tableDifferences : DWord; { any combination of PE_TCD_* }
    reserved1        : DWord; { reserved - do not use }
    reserved2        : DWord; { reserved - do not use }
  end;

const
  PE_SIZEOF_TABLE_DIFFERENCE_INFO = SizeOf(PETableDifferenceInfo);

{ Not implemented for Reports based on Dictionaries: returns PE_ERR_NOTIMPLEMENTED }
function PECheckNthTableDifferences (
  printJob                : Smallint;
  tableN                  : Smallint;
  var tabledifferenceinfo : PETableDifferenceInfo): Bool stdcall;


{******************************************************************************}
{ Overriding SQL Query in Report                                               }
{   PEGetSQLQuery () returns the same query as appears in the Show SQL Query   }
{   dialog in CRW, in syntax specific to the database driver you are using.    }
{                                                                              }
{   PESetSQLQuery () is mostly useful for reports with SQL queries that        }
{   were explicitly edited in the Show SQL Query dialog in CRW, i.e. those     }
{   reports that needed database-specific selection criteria or joins.         }
{   (Otherwise it is usually best to continue using function calls such as     }
{   PESetSelectionFormula () and let CRW build the SQL query automatically.    }
{                                                                              }
{   PESetSQLQuery () has the same restrictions as editing in the Show SQL      }
{   Query dialog; in particular that changes are accepted in the FROM and      }
{   WHERE clauses but ignored in the SELECT list of fields.                    }
{******************************************************************************}
function PEGetSQLQuery (
  printJob       : Smallint;
  var textHandle : HWnd;
  var textLength : Smallint): Bool stdcall;

function PESetSQLQuery (
  printJob    : Smallint;
  queryString : PChar): Bool stdcall;


{******************************************************************************}
{ Saved Data                                                                   }
{   Use PEHasSavedData() to find out if a report currently has saved data      }
{   associated with it.  This may or may not be TRUE when a print job is       }
{   first opened from a report file.  Since data is saved during a print,      }
{   this will always be TRUE immediately after a report is printed.            }
{                                                                              }
{   Use PEDiscardSavedData() to release the saved data associated with a       }
{   report.  The next time the report is printed, it will get current data     }
{   from the database.                                                         }
{                                                                              }
{   The default behavior is for a report to use its saved data, rather than    }
{   refresh its data from the database when printing a report.                 }
{******************************************************************************}
function PEHasSavedData (
  printJob         : Smallint;
  var hasSavedData : Bool): Bool stdcall;

function PEDiscardSavedData (printJob : Smallint): Bool stdcall;


{******************************************************************************}
{ Report Title                                                                 }
{******************************************************************************}
function PEGetReportTitle (
  printJob        : Smallint;
  var titleHandle : HWnd;
  var titleLength : Smallint): Bool stdcall;

function PESetReportTitle (
  printJob : Smallint;
  title    : PChar): Bool stdcall;


{******************************************************************************}
{ Output to Window                                                             }
{******************************************************************************}
function PEOutputToWindow (
  printJob     : Smallint;
  title        : PChar;
  left         : Integer;
  top          : Integer;
  width        : Integer;
  height       : Integer;
  style        : DWord;
  parentWindow : HWnd): Bool stdcall;

function PEGetWindowHandle (printJob: Smallint): HWnd stdcall;

procedure PECloseWindow (printJob: Smallint) stdcall;

{------------------------------------------------------------------------------}
{ Window Options/Print Controls                                                }
{------------------------------------------------------------------------------}
type
  PEWindowOptions = record
    StructSize            : Word;     { initialize to PE_SIZEOF_WINDOW_OPTIONS}
    hasGroupTree          : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    canDrillDown          : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasNavigationControls : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasCancelButton       : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasPrintButton        : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasExportButton       : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasZoomControl        : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasCloseButton        : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasProgressControls   : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasSearchButton       : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasPrintSetupButton   : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    hasRefreshButton      : Smallint; { 0 or 1 except use PE_UNCHANGED for no change}
    showToolbarTips       : Smallint; { BOOL value, except use PE_UNCHANGED for no change}
                                      { default is TRUE (*Show* tooltips on toolbar)}
    showDocumentTips      : Smallint; { BOOL value, except use PE_UNCHANGED for no change}
                                      { default is FALSE (*Hide* tooltips on document)}
    hasLaunchButton       : Smallint; { Launch Seagate Analysis button on toolbar. }
                                      { BOOL value, except use PE_UNCHANGED for no change }
                                      { default is FALSE }
  end;

const
  PE_SIZEOF_WINDOW_OPTIONS = SizeOf(PEWindowOptions);

function PEGetWindowOptions (
  printJob    : Smallint;
  var options : PEWindowOptions): Bool stdcall;

function PESetWindowOptions (
  printJob    : Smallint;
  var options : PEWindowOptions): Bool stdcall;

function PEShowPrintControls (
  printJob          : Smallint;
  showPrintControls : Bool): Bool stdcall;

function PEPrintControlsShowing (
  printJob            : Smallint;
  var controlsShowing : Bool): Bool stdcall;

{------------------------------------------------------------------------------}
{ Paging                                                                       }
{------------------------------------------------------------------------------}
function PEShowFirstPage (printJob: Smallint): Bool stdcall;
function PEShowNextPage (printJob: Smallint): Bool stdcall;
function PEShowPreviousPage (printJob: Smallint): Bool stdcall;
function PEShowLastPage (printJob: Smallint): Bool stdcall;

function PEGetNPages (printJob: Smallint): Smallint stdcall;

function PEShowNthPage (
  printJob : Smallint;
  pageN    : Smallint): Bool stdcall;

{------------------------------------------------------------------------------}
{ Window Zoom                                                                  }
{------------------------------------------------------------------------------}
const
  PE_ZOOM_FULL_SIZE           = 0;
  PE_ZOOM_SIZE_FIT_ONE_SIDE   = 1;
  PE_ZOOM_SIZE_FIT_BOTH_SIDES = 2;

function PEZoomPreviewWindow (
  printJob : Smallint;
  level    : Smallint): Bool stdcall;
  {level: a percent from 25 to 400 or a PE_ZOOM_ constant}

function PENextPrintWindowMagnification (printJob:  Smallint): Bool stdcall;

{------------------------------------------------------------------------------}
{ Print/Export Preview Window                                                  }
{------------------------------------------------------------------------------}
function PEPrintWindow (
  printJob      : Smallint;
  waitUntilDone : Bool): Bool stdcall;

function PEExportPrintWindow (
  printJob      : Smallint;
  toMail        : Bool;
  waitUntilDone : Bool): Bool stdcall;


{******************************************************************************}
{ Output to Printer                                                            }
{******************************************************************************}
function PEOutputToPrinter (
  printJob : Smallint;
  nCopies  : Smallint): Bool stdcall;

function PESelectPrinter (
  printJob    : Smallint;
  driverName  : PChar;
  printerName : PChar;
  portName    : PChar;
  mode        : PDevMode): Bool stdcall;

function PEGetSelectedPrinter (
  printJob          : Smallint;
  var driverHandle  : Hwnd;
  var driverLength  : Smallint;
  var printerHandle : Hwnd;
  var printerLength : Smallint;
  var portHandle    : Hwnd;
  var portLength    : Smallint;
  var mode          : PDevMode): Bool stdcall;

function PEFreeDevMode (
  printJob : Smallint; 
  mode     : PDevMode): Bool stdcall;

{******************************************************************************}
{ Detail Copies                                                                }
{   Commonly used for label-style Reports                                      }
{******************************************************************************}
function PESetNDetailCopies (
  printJob : Smallint;
  nCopies  : Smallint): Bool stdcall;

function PEGetNDetailCopies (
  printJob          : Smallint;
  var nDetailCopies : Smallint): Bool stdcall;


{******************************************************************************}
{ Print Options                                                                }
{   Extension to PESetPrintOptions function: If the 2nd parameter              }
{   (pointer to PEPrintOptions) is set to 0 (null) the function prompts        }
{   the user for these options.                                                }
{                                                                              }
{   With this change, you can get the behaviour of the print-to-printer        }
{   button in the print window by calling PESetPrintOptions with a             }
{   null pointer and then calling PEPrintWindow.                               }
{******************************************************************************}
const
  {Start/Stop Page}
  PE_MAXPAGEN            = 65535;
  {Output FileName}
  PE_FILE_PATH_LEN       = 512;
  {Collation}
  PE_UNCOLLATED          = 0;
  PE_COLLATED            = 1;
  PE_DEFAULTCOLLATION    = 2;

type
  PEOutputFileNameType = array [0..PE_FILE_PATH_LEN-1] of Char;
  PEPrintOptions  = record
    StructSize     : Word; { initialize to SizeOf(PEPrintOptions) }
    { page and copy numbers are 1-origin }
    { use 0 to preserve the existing Strings }
    StartPageN     : Word;
    StopPageN      : Word;
    nReportCopies  : Word;
    Collation      : Word;
    outputFileName : PEOutputFileNameType;
  end;

function PESetPrintOptions (
  printJob    : Smallint;
  var options : PEPrintOptions): Bool stdcall;

function PEGetPrintOptions (
  printJob    : Smallint;
  var options : PEPrintOptions):Bool stdcall;


{******************************************************************************}
{ Exporting                                                                    }
{******************************************************************************}
const
  UXFCrystalReportType    = 0;
  UXFWordWinType          = 0;
  UXFWordDosType          = 1;
  UXFWordPerfectType      = 2;
  UXFTextType             = 0;
  UXFTabbedTextType       = 1;
  UXFRichTextFormatType   = 0;
  UXFLotusWksType         = 0;
  UXFLotusWk1Type         = 1;
  UXFLotusWk3Type         = 2;
  UXFReportDefinitionType = 0;

type
  PEExportOptions  = record
    { Initialize to SizeOfPEExportOptions }
    StructSize               : Word;
    formatDLLName            : PEDllNameType;
    formatType               : DWord;
    formatOptions            : Pointer;
    destinationDLLName       : PEDllNameType;
    destinationType          : DWord;
    destinationOptions       : Pointer;
    nFormatOptionsBytes      : Word; { Set by PEExportOptions }
                                     { Ignored by PEExportTo  }
    nDestinationOptionsBytes : Word; { Set by PEExportOptions }
                                     { Ignored by PEExportTo  }
  end;

const
  PE_SIZEOF_EXPORT_OPTIONS = SizeOf(PEExportOptions);

function PEGetExportOptions (
  printJob    : Smallint;
  var options : PEExportOptions): Bool stdcall;

function PEExportTo (
  printJob    : Smallint;
  var options : PEExportOptions): Bool stdcall;

{------------------------------------------------------------------------------}
{ Export to Excel                                                              }
{------------------------------------------------------------------------------}
const
  UXFXls2Type            = 0;
  UXFXls3Type            = 1;
  UXFXls4Type            = 2;
  UXFXls5Type            = 3;
  UXFXls5TypeTab         = 4; {Left in for backward compatibility}
  UXFXls5TypeExt         = 4;
  UXFXl7Type             = 5;
  UXFXl7TabType          = 6;
  UXFXl8Type             = 7;
  UXFXl8TabType          = 8;
  DEFAULT_COLUMN_WIDTH   = 10;

type
  UXFXlsOptions = record
    structSize        : Word;
    bColumnHeadings   : Bool;
    {TRUE -- has column headings, which come from}
    {        "Page Header" and "Report Header" areas.}
    {FALSE -- no column headings.}
    {The default value is FALSE.}
    bUseconstColWidth : Bool;
    {TRUE -- use constant column width}
    {FALSE -- set column width based on an area}
    {The default value is FALSE.}
    fconstColWidth    : double;
    {Column width, when bUseconstColWidth is TRUE.}
    {The default value is DEFAULT_COLUMN_WIDTH (10)}
    bTabularFormat    : Bool;
    {TRUE -- tabular format (flatten an area into a row)}
    {FALSE -- non-tabular format}
    {The default value is FALSE.}
    baseAreaType      : AreaType;
    {One of the 7 area types defined by PE_SECT_ constants.}
    {The default value is PE_SECT_DETAIL Details.}
    baseAreaGroupNum  : Word;
    {If baseAreaType is either GroupHeader or GroupFooter,
    { and there are more than one group, we need to give the group number.}
    {The default value is 1.}
    bUseWorksheetFunc  : Bool;
    { If TRUE, use Excel worksheet functions to represent }
    { subtotal fields in SCR. }
    { The default value is TRUE. }
  end;

const
  UXFXlsOptionsSize = SizeOf(UXFXlsOptions);

{------------------------------------------------------------------------------}
{ Export to Lotus Notes Database                                               }
{------------------------------------------------------------------------------}
const
  UXDNotesType = 3;

type
  UXDNotesOptions = record
    structSize : Word;
    szDBName   : PChar;
    szFormName : PChar;  // should be "Report Form"
    szComments : PChar;
  end;

const
  UXDNotesOptionsSize = SizeOf(UXDNotesOptions);

{------------------------------------------------------------------------------}
{ Export to Disk                                                               }
{------------------------------------------------------------------------------}
const
  UXDDiskType = 0;

type
  UXDDiskOptions = record
    {Initialize to UXDDiskOptionsSize}
    structSize : Word;
    fileName   : PChar;
  end;

const
  UXDDiskOptionsSize = SizeOf(UXDDiskOptions);

{------------------------------------------------------------------------------}
{ Export to MAPI                                                               }
{------------------------------------------------------------------------------}
const
  UXDMAPIType = 0;

type
  UXDMAPIOptions = record
    {Initialize to UXDMAPIOptionsSize}
    structSize  : Word;
    toList      : PChar;
    ccList      : PChar;
    subject     : PChar;
    mailmessage : PChar;
  end;

const
  UXDMAPIOptionsSize = SizeOf(UXDMAPIOptions);

{------------------------------------------------------------------------------}
{ Export to DIF                                                                }
{------------------------------------------------------------------------------}
const
  UXFDIFType = 0;

type
  UXFDIFOptions = record
    {Initialize to UXFDIFOptionsSize}
    structSize            : Word;
    useReportNumberFormat : Bool;
    useReportDateFormat   : Bool;
  end;

const
  UXFDIFOptionsSize = SizeOf(UXFDIFOptions);

{------------------------------------------------------------------------------}
{ Export to RecordStyle                                                        }
{------------------------------------------------------------------------------}
const
  UXFRecordType = 0;

type
  UXFRecordStyleOptions = record
    {Initialize to UXFRecordStyleOptionsSize}
    structSize            : Word;
    useReportNumberFormat : Bool;
    useReportDateFormat   : Bool;
  end;

const
  UXFRecordStyleOptionsSize = SizeOf(UXFRecordStyleOptions);

{------------------------------------------------------------------------------}
{ Export to Comma/Tab Separated                                                }
{------------------------------------------------------------------------------}
const
  UXFCommaSeparatedType = 0;
  UXFTabSeparatedType   = 1;

type
  UXFCommaTabSeparatedOptions = record
    {Initialize to UXFCommaTabSeparatedOptionsSize}
    structSize            : Word;
    useReportNumberFormat : Bool;
    useReportDateFormat   : Bool;
  end;

const
  UXFCommaTabSeparatedOptionsSize = SizeOf(UXFCommaTabSeparatedOptions);

{------------------------------------------------------------------------------}
{ Export to Char Separated                                                     }
{------------------------------------------------------------------------------}
const
  UXFCharSeparatedType = 2;

type
  UXFCharSeparatedOptions = record
    {Initialize to UXFCharSeparatedOptionsSize}
    structSize            : Word;
    useReportNumberFormat : Bool;
    useReportDateFormat   : Bool;
    stringDelimiter       : Char;
    fieldDelimiter        : PChar;
  end;

const
  UXFCharSeparatedOptionsSize = SizeOf(UXFCharSeparatedOptions);

{------------------------------------------------------------------------------}
{ Export to VIM                                                                }
{------------------------------------------------------------------------------}
const
  UXDVIMType = 0;

type
  UXDVIMOptions = record
    {Initialize to UXDVIMOptionsSize}
    structSize  : Word;
    toList      : PChar;
    ccList      : PChar;
    bccList     : PChar;
    subject     : PChar;
    mailmessage : PChar;
  end;

const
  UXDVIMOptionsSize = SizeOf(UXDVIMOptions);

{------------------------------------------------------------------------------}
{ Export to Exchange                                                           }
{------------------------------------------------------------------------------}
const
  UXDExchFolderType = 0;
  UXDPostDocMessage = 1011; {wDestType for folder messages}

type
  UXDPostFolderOptions = record
    structSize    : Word;
    pszProfile    : PChar;
    pszPassword   : PChar;
    wDestType     : Word;
    pszFolderPath : PChar;
  end;

  { pszFolderPath has to be in the following format:
     <Message Store Name>@<Folder Name>@<Folder Name> }

const
  UXDPostFolderOptionsSize = SizeOf(UXDPostFolderOptions);

{------------------------------------------------------------------------------}
{ Export to ODBC                                                               }
{------------------------------------------------------------------------------}
const
  UXFODBCType = 0;

type
  UXFODBCOptions = record
    structSize         : Word;
    dataSourceName     : PChar;
    dataSourceUserID   : PChar;
    dataSourcePassword : PChar;
    exportTableName    : PChar;
  end;

const
  UXFODBCOptionsSize = SizeOf(UXFODBCOptions);

{------------------------------------------------------------------------------}
{ Export to HTML                                                               }
{------------------------------------------------------------------------------}
const
  UXFHTML3Type     = 0; {Draft HTML 3.0 tags}
  UXFExplorer2Type = 1; {Include MS Explorer 2.0 tags}
  UXFNetscape2Type = 2; {Include Netscape 2.0 tags}
  UXFHTML32ExtType = 1; {HTML 3.2 tags + bg color extensions}
  UXFHTML32StdType = 2; {HTML 3.2 tags}

type
  UXFHTML3Options = record
    structSize : Word;  { set to UXFHTML3OptionsSize }
    fileName   : PChar; { ptr to full Windows filepath of HTML output file }
                        { e.g. "C:\pub\docs\boxoffic\default.htm" }
                        { NOTE: any exported GIF files will be  }
                        { located in the same directory as this file }
  end;

const
  UXFHTML3OptionsSize = SizeOf(UXFHTML3Options);

{------------------------------------------------------------------------------}
{ Export to Paginated Text                                                     }
{------------------------------------------------------------------------------}
const
  UXFPaginatedTextType = 2;

type
  UXFPaginatedTextOptions = record
    structSize     :Word;
    nLinesPerPage  :Word;
  end;

const
  UXFPaginatedTextOptionsSize = SizeOf(UXFPaginatedTextOptions);

{------------------------------------------------------------------------------}
{ Export to Application                                                        }
{------------------------------------------------------------------------------}
const
  UXDApplicationType = 0;

type
  UXDApplicationOptions = record
    structSize : word;
    fileName : PChar;
  end;

const
  UXDApplicationOptionsSize = SizeOf(UXDApplicationOptions);


{******************************************************************************}
{ Setting Page Margins                                                         }
{******************************************************************************}
const
  PE_SM_DEFAULT = $8000;

function PESetMargins (
  printJob : Smallint;
  left     : Smallint;
  right    : Smallint;
  top      : Smallint;
  bottom   : Smallint): Bool stdcall;

function PEGetMargins (
  printJob   : Smallint;
  var left   : Smallint;
  var right  : Smallint;
  var top    : Smallint;
  var bottom : Smallint): Bool stdcall;


{******************************************************************************}
{ Report Summary Info                                                          }
{******************************************************************************}
const
  PE_SI_APPLICATION_NAME_LEN  = 128;
  PE_SI_TITLE_LEN             = 128;
  PE_SI_SUBJECT_LEN           = 128;
  PE_SI_AUTHOR_LEN            = 128;
  PE_SI_KEYWORDS_LEN          = 128;
  PE_SI_COMMENTS_LEN          = 512;
  PE_SI_REPORT_TEMPLATE_LEN   = 128;

type
  PEApplicationNameType = array [0..PE_SI_APPLICATION_NAME_LEN-1] of char;
  PETitleType           = array [0..PE_SI_TITLE_LEN-1] of char;
  PESubjectType         = array [0..PE_SI_SUBJECT_LEN-1] of char;
  PEAuthorType          = array [0..PE_SI_AUTHOR_LEN-1] of char;
  PEKeywordsType        = array [0..PE_SI_KEYWORDS_LEN-1] of char;
  PECommentsType        = array [0..PE_SI_COMMENTS_LEN-1] of char;
  PEReportTemplateType  = array [0..PE_SI_REPORT_TEMPLATE_LEN-1] of char;
  PEReportSummaryInfo = record
    StructSize      : Word;
    applicationName : PEApplicationNameType; { read only.}
    title           : PETitleType;
    subject         : PESubjectType;
    author          : PEAuthorType;
    keywords        : PEKeywordsType;
    comments        : PECommentsType;
    reportTemplate  : PEReportTemplateType;
    savePreviewPicture : Smallint; { BOOL PE_UNCHANGED for no change }
  end;

const
  PE_SIZEOF_REPORT_SUMMARY_INFO = SizeOf(PEReportSummaryInfo);

function PEGetReportSummaryInfo (
  printJob        : Smallint;
  var summaryInfo : PEReportSummaryInfo): Bool stdcall;

function PESetReportSummaryInfo (
  printJob        : Smallint;
  var summaryInfo : PEReportSummaryInfo): Bool stdcall;


{******************************************************************************}
{ Callback Events                                                              }
{******************************************************************************}
const
  { event ID }
  PE_CLOSE_PRINT_WINDOW_EVENT           = 1;
  PE_ACTIVATE_PRINT_WINDOW_EVENT        = 2;
  PE_DEACTIVATE_PRINT_WINDOW_EVENT      = 3;
  PE_PRINT_BUTTON_CLICKED_EVENT         = 4;
  PE_EXPORT_BUTTON_CLICKED_EVENT        = 5;
  PE_ZOOM_LEVEL_CHANGING_EVENT          = 6;
(*  PE_ZOOM_CONTROL_SELECTED_EVENT        = 6;*)
  PE_FIRST_PAGE_BUTTON_CLICKED_EVENT    = 7;
  PE_PREVIOUS_PAGE_BUTTON_CLICKED_EVENT = 8;
  PE_NEXT_PAGE_BUTTON_CLICKED_EVENT     = 9;
  PE_LAST_PAGE_BUTTON_CLICKED_EVENT     = 10;
  PE_CANCEL_BUTTON_CLICKED_EVENT        = 11;
  PE_CLOSE_BUTTON_CLICKED_EVENT         = 12;
  PE_SEARCH_BUTTON_CLICKED_EVENT        = 13;
      PE_GROUP_TREE_BUTTON_CLICKED_EVENT     = 14;
  PE_PRINT_SETUP_BUTTON_CLICKED_EVENT   = 15;
  PE_REFRESH_BUTTON_CLICKED_EVENT       = 16;
  PE_SHOW_GROUP_EVENT                   = 17;
  PE_DRILL_ON_GROUP_EVENT	        = 18; { include drill on graph }
  PE_DRILL_ON_DETAIL_EVENT              = 19;
  PE_READING_RECORDS_EVENT              = 20;
  PE_START_EVENT                        = 21;
  PE_STOP_EVENT                         = 22;
  PE_MAPPING_FIELD_EVENT                = 23;
  PE_RIGHT_CLICK_EVENT                  = 24; { right mouse click }
  PE_LEFT_CLICK_EVENT                   = 25; { left mouse click }
  PE_MIDDLE_CLICK_EVENT                 = 26; { middle mouse click }

{ All Events are disabled by default; use PEEnableEvent to enable events }
type
  PEEnableEventInfo = record
    StructSize                 : Word;
    startStopEvent             : Smallint; {0 or 1, PE_UNCHANGED for no change}
    readingRecordEvent         : Smallint; {0 or 1, PE_UNCHANGED for no change}
    printWindowButtonEvent     : Smallint; {0 or 1, PE_UNCHANGED for no change}
    drillEvent                 : Smallint; {0 or 1, PE_UNCHANGED for no change}
    closePrintWindowEvent      : Smallint; {0 or 1, PE_UNCHANGED for no change}
    activatePrintWindowEvent   : Smallint; {0 or 1, PE_UNCHANGED for no change}
    fieldMappingEvent          : Smallint; {Bool value, PE_UNCHANGED for no change}
    mouseClickEvent            : Smallint; {Bool value, PE_UNCHANGED for no change}
    hyperlinkEvent             : Smallint; {Bool value, PE_UNCHANGED for no change}
    launchSeagateAnalysisEvent : Smallint; {Bool value, PE_UNCHANGED for no change}
  end;

const
  PE_SIZEOF_ENABLE_EVENT_INFO = SizeOf(PEEnableEventInfo);

function PEEnableEvent (
  printJob            : Smallint;
  var enableEventInfo : PEEnableEventInfo): Bool stdcall;

function PEGetEnableEventInfo (
  printJob            : Smallint;
  var enableEventInfo : PEEnableEventInfo): Bool stdcall;

{ Set callback function}
function PESetEventCallback (
  printJob     : Smallint;
  callbackProc : pointer;
  userData     : pointer): Bool stdcall;
  (*
  callbackProc function should be of form:
    function callbackProc (
      eventID : Smallint;  {event ID constant}
      param2  : pointer;   {pointer to Event structure for event ID}
      param3  : pointer    {user-defined pointer}
      ): Bool stdcall;
  *)  {should "export" be used for 16-bit callback?}

{------------------------------------------------------------------------------}
{ General PrintWindow Event                                                    }
{------------------------------------------------------------------------------}
{ use this structure for
  PE_CLOSE_PRINT_WINDOW_EVENT
  PE_ACTIVATE_PRINT_WINDOW_EVENT
  PE_DEACTIVATE_PRINT_WINDOW_EVENT
  PE_PRINT_BUTTON_CLICKED_EVENT
  PE_EXPORT_BUTTON_CLICKED_EVENT
  PE_FIRST_PAGE_BUTTON_CLICKED_EVENT
  PE_PREVIOUS_PAGE_BUTTON_CLICKED_EVENT
  PE_NEXT_PAGE_BUTTON_CLICKED_EVENT
  PE_LAST_PAGE_BUTTON_CLICKED_EVENT
  PE_CANCEL_BUTTON_CLICKED_EVENT
  PE_PRINT_SETUP_BUTTON_CLICKED_EVENT
  PE_REFRESH_BUTTON_CLICKED_EVENT }
type
  PEGeneralPrintWindowEventInfo = record
    StructSize   : Word;
    ignored      : Word; {for 4 byte alignment. ignore.}
    windowHandle : HWnd;
  end;

const
  PE_SIZEOF_GENERAL_PRINT_WINDOW_EVENT_INFO = SizeOf(PEGeneralPrintWindowEventInfo);

{------------------------------------------------------------------------------}
{ Mouse Click Event                                                            }
{------------------------------------------------------------------------------}
{ use this structure for
  PE_RIGHT_CLICK_EVENT
  PE_LEFT_CLICK_EVENT
  PE_MIDDLE_CLICK_EVENT }
const
  { mouse click action }
  PE_MOUSE_NOTSUPPORTED = 0;
  PE_MOUSE_DOWN         = 1;
  PE_MOUSE_UP           = 2;
  PE_MOUSE_DOUBLE_CLICK = 3;

  { mouse click flags (virtual key state-masks) }
  PE_CF_NONE       = $0000;
  PE_CF_LBUTTON    = $0001;
  PE_CF_RBUTTON    = $0002;
  PE_CF_SHIFTKEY   = $0004;
  PE_CF_CONTROLKEY = $0008;
  PE_CF_MBUTTON    = $0010;

{for PE_RIGHT_CLICK_EVENT}
type
  PEMouseClickEventInfo = record
    StructSize   : Word;
    windowHandle : LongInt;
    clickAction  : integer;     {PE_MOUSE constants: mouse button down or up}
    clickFlags   : integer;     {any combination of PE_CF_*}
    xOffset      : integer;     {x-coordinate of mouse click in pixels}
    yOffset      : integer;     {y-coordinate of mouse click in pixels}
    fieldValue   : PEValueInfo;	{value of object at click point if it is a field}
                                {object, excluding MEMO and BLOB fields,}
                                {else valueType element = PE_VI_NOVALUE.}
    objectHandle : DWord;	{the design view object}
    sectionCode  : Smallint;	{section in which click occurred.}
  end;

const
  PE_SIZEOF_MOUSE_CLICK_EVENT_INFO = SizeOf(PEMouseClickEventInfo);

{------------------------------------------------------------------------------}
{ Start/Stop Event                                                             }
{------------------------------------------------------------------------------}
{ use this structure for
  PE_START_EVENT
  PE_STOP_EVENT }
const
  { job destination }
  PE_TO_NOWHERE = 0;
  PE_TO_WINDOW  = 1;
  PE_TO_PRINTER = 2;
  PE_TO_EXPORT  = 3;
  PE_FROM_QUERY = 4;

{ for PE_START_EVENT}
type
  PEStartEventInfo = record
    StructSize  : Word;
    destination : Word;
  end;

const
  PE_SIZEOF_START_EVENT_INFO = SizeOf(PEStartEventInfo);

{ for PE_STOP_EVENT}
type
  PEStopEventInfo = record
    StructSize  : Word;
    destination : Word;
    jobStatus   : Word; {a PE_JOB constant}
  end;

const
  PE_SIZEOF_STOP_EVENT_INFO = SizeOf(PEStopEventInfo);

{------------------------------------------------------------------------------}
{ Reading Records Event                                                        }
{------------------------------------------------------------------------------}
{ for PE_READING_RECORDS_EVENT }
type
  PEReadingRecordsEventInfo = record
    StructSize      : Word;
    cancelled       : Smallint; {0 or 1.}
    recordsRead     : DWord;
    recordsSelected : DWord;
    done            : Smallint; {0 or 1}
  end;

const
  PE_SIZEOF_READING_RECORDS_EVENT_INFO = SizeOf(PEReadingRecordsEventInfo);

{------------------------------------------------------------------------------}
{ Zoom Control Event                                                           }
{------------------------------------------------------------------------------}
{ for PE_ZOOM_CONTROL_SELECTED_EVENT }
type
  PEZoomLevelChangingEventInfo = record
    StructSize   : Word;
    zoomLevel    : Word;
    windowHandle : HWnd;
 end;

const
  PE_SIZEOF_ZOOM_LEVEL_CHANGING_EVENT_INFO = SizeOf(PEZoomLevelChangingEventInfo);

{------------------------------------------------------------------------------}
{ Close Button Clicked Event                                                   }
{------------------------------------------------------------------------------}
{ for PE_CLOSE_BUTTON_CLICKED_EVENT}
type
  PECloseButtonClickedEventInfo = record
    StructSize   : Word;
    viewIndex    : Word; {which view is going to be closed. start from zero.}
    windowHandle : HWnd; {frame window handle the button is on.}
  end;

const
  PE_SIZEOF_CLOSE_BUTTON_CLICKED_EVENT_INFO = SizeOf(PECloseButtonClickedEventInfo);

{------------------------------------------------------------------------------}
{ Search Button Clicked Event                                                  }
{------------------------------------------------------------------------------}
{for PE_SEARCH_BUTTON_CLICKED_EVENT}
const
  PE_SEARCH_STRING_LEN = 128;

type
  PESearchStringType = array [0..PE_SEARCH_STRING_LEN-1] of Char;
  PESearchButtonClickedEventInfo = record
    windowHandle : HWnd;
    searchString : PESearchStringType;
    StructSize   : Word;
  end;

const
  PE_SIZEOF_SEARCH_BUTTON_CLICKED_EVENT_INFO = SizeOf(PESearchButtonClickedEventInfo);

{------------------------------------------------------------------------------}
{ GroupTree Button Clicked Event                                               }
{------------------------------------------------------------------------------}
{for PE_GROUPTREE_BUTTON_CLICKED_EVENT}
type
  PEGroupTreeButtonClickedEventInfo = record
    StructSize   : Word;
    Visible      : Smallint; {0 or 1. Current state of the group tree button}
    windowHandle : HWnd;
  end;

const
  PE_SIZEOF_GROUP_TREE_BUTTON_CLICKED_EVENT_INFO = SizeOf(PEGroupTreeButtonClickedEventInfo);

{------------------------------------------------------------------------------}
{ Show Group Event                                                             }
{------------------------------------------------------------------------------}
{for PE_SHOW_GROUP_EVENT}
type
  PEGroupArray = array[0..0] of PChar;
  PPEGroupArray = ^PEGroupArray;
  PEShowGroupEventInfo = record
    StructSize   : Word;
    groupLevel   : Word;
    windowHandle : HWnd;
    groupList    : PPEGroupArray;
    { points to an array of group name PChars }
    { PChar memory is freed after calling the call back function }
  end;

const
  PE_SIZEOF_SHOW_GROUP_EVENT_INFO = SizeOf(PEShowGroupEventInfo);

(******************************************************************************
** An example of using the PEGroupArray (array of pointers)
**   - because the array is defined with zero elements,
**     we turn range-checking off in code.
*******************************************************************************
  {Local procedure WindowCallback}
  function WindowCallback(eventID: Smallint; pEvent: pointer; nil): LongBool stdcall;
  var
    ShowGroupInfo : PEShowGroupEventInfo;
    slGroups      : TStringList;
    cnt           : integer;
  begin
    case eventID of
      PE_SHOW_GROUP_EVENT :
        begin
          ShowGroupInfo := PEShowGroupEventInfo(pEvent^);
          if (ShowGroupInfo.structSize = SizeOf(PEShowGroupEventInfo)) then
          begin
            slGroups := TStringList.Create;
            {If Range Checking is on, turn it off}
            {$IFOPT R+}
              {$DEFINE CKRANGE}
              {$R-}
            {$ENDIF}
            for cnt := 0 to ShowGroupInfo.groupLevel - 1 do
              slGroups.Add(String(ShowGroupInfo.groupList^[cnt]));
            {$IFDEF CKRANGE}
              {$UNDEF CKRANGE}
              {$R+}
            {$ENDIF}
            {...do something...}
            slGroups.Free;
            Result := True; //to continue processing
          end;
        end;
    end;
  end;
*)

{------------------------------------------------------------------------------}
{ Drill On Group Event                                                         }
{------------------------------------------------------------------------------}
{For PE_DRILL_ON_GROUP_EVENT}
const
  PE_DE_ON_GROUP     = 0;
  PE_DE_ON_GROUPTREE = 1;
  PE_DE_ON_GRAPH     = 2;
  PE_DE_ON_MAP       = 3;
  PE_DE_ON_SUBREPORT = 4;

type
  PEDrillOnGroupEventInfo = record
    StructSize   : Word;
    drillType    : Word; { a PE_DE_ constant}
    windowHandle : HWnd;
    { Points to an array of group name. Memory pointed by group list
    { is freed after calling the callback function }
    groupList    : PPEGroupArray;
    groupLevel   : Word;
  end;
  PPEDrillOnGroupEventInfo = ^PEDrillOnGroupEventInfo;

const
  PE_SIZEOF_DRILL_ON_GROUP_EVENT_INFO = SizeOf(PEDrillOnGroupEventInfo);

{------------------------------------------------------------------------------}
{ Drill On Detail Event                                                        }
{------------------------------------------------------------------------------}
{ for PE_DRILL_ON_DETAIL_EVENT }
type
  PEFieldValueInfo = record
    StructSize : Word;
    ignored    : Word;  {for 4 byte alignment. ignore.}
    fieldName:  PEFieldNameType; {defined in GroupCondition/GroupOptions}
    fieldValue: PEValueInfo;
  end;
  PPEFieldValueInfo     = ^PEFieldValueInfo;
  PEFieldValueInfoArray = array [0..0] of PPEFieldValueInfo;
  PPEFieldValueInfoArray = ^PEFieldValueInfoArray;

const
  PE_SIZEOF_FIELD_VALUE_INFO = SizeOf(PEFieldValueInfo);

type
  PEDrillOnDetailEventInfo = record
    StructSize         : Word;
    selectedFieldIndex : Smallint; {-1 if no field selected}
    windowHandle       : HWnd;
    { points to an array of PEFieldValue. memory
      pointed by fieldValueList is freed after
      calling the call back function }
    fieldValueList     : PPEFieldValueInfoArray;
    nFieldValue        : Smallint; {number of field value in fieldValueList}
  end;

const
  PE_SIZEOF_DRILL_ON_DETAIL_EVENT_INFO = SizeOf(PEDrillOnDetailEventInfo);

{------------------------------------------------------------------------------}
{ Field Mapping Event                                                          }
{------------------------------------------------------------------------------}
const
  PE_TABLE_NAME_LEN          = 128;
  PE_DATABASE_FIELD_NAME_LEN = 128;

  { Field value type }
  PE_FVT_INT8SFIELD           = 1;
  PE_FVT_INT8UFIELD           = 2;
  PE_FVT_INT16SFIELD          = 3;
  PE_FVT_INT16UFIELD          = 4;
  PE_FVT_INT32SFIELD          = 5;
  PE_FVT_INT32UFIELD          = 6;
  PE_FVT_NUMBERFIELD          = 7;
  PE_FVT_CURRENCYFIELD        = 8;
  PE_FVT_BOOLEANFIELD         = 9;
  PE_FVT_DATEFIELD            = 10;
  PE_FVT_TIMEFIELD            = 11;
  PE_FVT_STRINGFIELD          = 12;
  PE_FVT_TRANSIENTMEMOFIELD   = 13;
  PE_FVT_PERSISTENTMEMOFIELD  = 14;
  PE_FVT_BLOBFIELD            = 15;
  PE_FVT_DATETIMEFIELD        = 16;
  PE_FVT_BITMAPFIELD          = 17;
  PE_FVT_ICONFIELD            = 18;
  PE_FVT_PICTUREFIELD         = 19;
  PE_FVT_OLEFIELD             = 20;
  PE_FVT_GRAPHFIELD           = 21;
  PE_FVT_UNKNOWNFIELD         = 22;

  { Field mapping types }
  PE_FM_AUTO_FLD_MAP          = 0;
  { Auto field name mapping }
  { Unmapped Report fields will be removed }

  PE_FM_CRPE_PROMPT_FLD_MAP   = 1;
  { CRPE provides dialog box to map field manually }

  PE_FM_EVENT_DEFINED_FLD_MAP = 2;
  { CRPE provides list of field in report and new database }
  { User needs to activate the PE_MAPPING_FIELD_EVENT }
  { and define a callback function }

function PESetFieldMappingType (
  printJob    : Smallint;
  mappingType : Word): Bool stdcall;

function PEGetFieldMappingType (
  printJob        : Smallint;
  var mappingType : Word): Bool stdcall;

type
  PETableAliasNameType     = array [0..PE_TABLE_NAME_LEN -1] of Char;
  PEDatabaseFieldNameType  = array [0..PE_DATABASE_FIELD_NAME_LEN -1] of Char;
  PEReportFieldMappingInfo = record
    StructSize        : Word;
    valueType         : Word; { a PE_FVT_constant }
    tableAliasName    : PETableAliasNameType;
    databaseFieldName : PEDatabaseFieldNameType;
    mappingTo         : integer;
    { Mapped fields are assigned to the index of a field }
    { in array PEFieldMappingEventInfo->databaseFields. }
    { Unmapped fields are assigned to -1 }
  end;
  PPEReportFieldMappingInfo = ^PEReportFieldMappingInfo;
  PEMappingInfoArray = array[0..0] of PPEReportFieldMappingInfo;
  PPEMappingInfoArray = ^PEMappingInfoArray;

const
  PE_SIZEOF_REPORT_FIELDMAPPING_INFO = SizeOf(PEReportFieldMappingInfo);

type
  PEFieldMappingEventInfo = record
    StructSize      : Word;
    reportFields    : PPEMappingInfoArray;
    { An array of pointers to the fields in the report }
    { User need to modify the 'mappingTo' of each new mapped field
    { by assigning the value of the index of a field in the array of databaseFields }
    nReportFields   : Word; { Size of array reportFields }
    databaseFields  : PPEMappingInfoArray;
    { An array of pointers to the fields in the new database file }
    nDatabaseFields : Word; { Size of array databaseField }
  end;


const
  PE_SIZEOF_FIELDMAPPING_EVENT_INFO = SizeOf(PEFieldMappingEventInfo);

{------------------------------------------------------------------------------}
{ Hyperlink Event                                                              }
{------------------------------------------------------------------------------}
const
  PE_DRILL_ON_HYPERLINK_EVENT      = 27;
  PE_LAUNCH_SEAGATE_ANALYSIS_EVENT = 28;

type
  PEHyperlinkEventInfo = record
    StructSize    : Word;
    ignored       : Word;    { for 4 byte alignment. ignore. }
    windowHandle  : Longint; { HWND }
    hyperlinkText : PChar;   { points to the hyperlink text associated with the object. }
                             { memory pointed by hyperlinkText is freed after calling the callback function. }
  end;

const
  PE_SIZEOF_HYPERLINKEVENTINFO = SizeOf(PEHyperlinkEventInfo);

{------------------------------------------------------------------------------}
{ LaunchSeagateAnalysis Event                                                  }
{------------------------------------------------------------------------------}
type
  PELaunchSeagateAnalysisEventInfo = record
    StructSize   : Word;
    ignored      : Word;    { for 4 byte alignment. ignore. }
    windowHandle : Longint; { HWND }
    pathFile     : PChar;   { points to the path and filename of the temporary report. }
                            { memory pointed by pathFile is freed after calling the callback function. }
  end;

const
  PE_SIZEOF_LAUNCHSEAGATEANALYSISEVENTINFO = SizeOf(PELaunchSeagateAnalysisEventInfo);


{******************************************************************************}
{ Window Cursor                                                                }
{******************************************************************************}
{ All are window standard cursors except PE_TC_MAGNIFY_CURSOR.
  PE_TC_SIZEALL_CURSOR, PE_TC_NO_CURSOR, PE_TC_APPSTARTING_CURSOR
  and PE_TC_HELP_CURSOR are  not supported in 16-bit.
  PE_TC_SIZE_CURSOR and PE_TC_ICON_CURSOR are obsolete for 32-bit;
  use PE_TC_SIZEALL_CURSOR and PE_TC_ARROW_CURSOR instead. }
const
  PE_TC_DEFAULT_CURSOR     = 0; {CRPE set default cursor to be PE_TC_ARRAOW_CURSOR}
  PE_TC_ARROW_CURSOR       = 1;
  PE_TC_CROSS_CURSOR       = 2;
  PE_TC_IBEAM_CURSOR       = 3;
  PE_TC_UPARROW_CURSOR     = 4;
  PE_TC_SIZEALL_CURSOR     = 5;  {32-bit}
  PE_TC_SIZENWSE_CURSOR    = 6;
  PE_TC_SIZENESW_CURSOR    = 7;
  PE_TC_SIZEWE_CURSOR      = 8;
  PE_TC_SIZENS_CURSOR      = 9;
  PE_TC_NO_CURSOR          = 10; {32-bit}
  PE_TC_WAIT_CURSOR        = 11;
  PE_TC_APPSTARTING_CURSOR = 12; {32-bit}
  PE_TC_HELP_CURSOR	   = 13; {32-bit}
  PE_TC_SIZE_CURSOR        = 14; {16-bit}
  PE_TC_ICON_CURSOR        = 15; {16-bit}
  { 32-bit, CRPE specific cusors }
  PE_TC_BACKGROUND_PROCESS_CURSOR = 94;
  PE_TC_GRAB_HAND_CURSOR          = 95;
  PE_TC_ZOOM_IN_CURSOR            = 96;
  PE_TC_REPORT_SECTION_CURSOR     = 97;
  PE_TC_HAND_CURSOR               = 98;
  PE_TC_MAGNIFY_CURSOR            = 99;

type
  PETrackCursorInfo = record
    StructSize                  : Word;
    groupAreaCursor             : Smallint; {a PE_TC constant. PE_UNCHANGED for no change}
    groupAreaFieldCursor        : Smallint; {a PE_TC constant. PE_UNCHAGNED for no change}
    detailAreaCursor            : Smallint; {a PE_TC constant. PE_UNCHANGED for no change}
    detailAreaFieldCursor       : Smallint; {a PE_TC constant. PE_UNCHANGED for no change}
    graphCursor                 : Smallint; {a PE_TC constant. PE_UNCHANGED for no change}
    groupAreaCursorHandle       : DWord;    {reserved}
    groupAreaFieldCursorHandle  : DWord;    {reserved}
    detailAreaCursorHandle      : DWord;    {reserved}
    detailAreaFieldCursorHandle : DWord;    {reserved}
    graphCursorHandle           : DWord;    {reserved}
    ondemandSubreportCursor     : Smallint; { Cursor to show over on-demand subreports when }
                                            { drilldown for the window is enabled; }
                                            { default is PE_TC_MAGNIFY_CURSOR. }
    hyperlinkCursor             : Smallint; { Cursor to show over report object that has hyperlink text; }
                                            { default is PE_TC_HAND_CURSOR. }
  end;

const
  PE_SIZEOF_TRACK_CURSOR_INFO = SizeOf(PETrackCursorInfo);

{ set tracking cursor}
function PEGetTrackCursorInfo (
  printJob       : Smallint;
  var cursorInfo : PETrackCursorInfo): Bool stdcall;

function PESetTrackCursorInfo (
  printJob       : Smallint;
  var cursorInfo : PETrackCursorInfo): Bool stdcall;


{******************************************************************************}
{ Formula Syntax                                                               }
{******************************************************************************}
const
  PE_FST_CRYSTAL = 0;
  PE_FST_BASIC   = 1;
  PE_FS_SIZE    = 2;

type
  PEFormulaSyntaxArray = array[0..PE_FS_SIZE-1] of Smallint;
  PEFormulaSyntax = record
    StructSize    : Word;
    formulaSyntax : PEFormulaSyntaxArray; { PE_FS_* or PE_UNCHANGED. }
  end;

const
  PE_SIZEOF_FORMULA_SYNTAX = SizeOf(PEFormulaSyntax);

{ PEGetFormulaSyntax                                                               }
{ Indicates the syntax used by the formula addressed in the last formula API call. }
{ The syntax type is returned in formulaSyntax[0];                                 }
{ formulaSyntax[1] is reserved for internal use.                                   }
{ ***Default Behaviour: If this API is called before any Formula API is called     }
{                       then the values returned will be PE_UNCHANGED.             }


function PEGetFormulaSyntax (
  printJob          : Smallint;
  var formulaSyntax : PEFormulaSyntax): Bool stdcall;

{ PESetFormulaSyntax                                                      }
{ Use this API to set the syntax to use in the next (and all successive)  }
{ formula API call(s).                                                    }
{ Set one of PE_FST_* into formulaSyntax[0];                              }
{ formulaSyntax[1] is reserved for internal use.                          }
{ ***Default Behaviour: If any Formula API is called before calling       }
{                       PESetFormulaSyntax then PE_FST_CRYSTAL is assumed.}

function PESetFormulaSyntax (
  printJob          : Smallint; 
  var formulaSyntax : PEFormulaSyntax): Bool stdcall;



{******End of Interface Section******}

implementation

function PE_SECTION_CODE(sectionType, groupN, sectionN : Smallint): Smallint stdcall;
begin
  Result := (sectionType * 1000) + (groupN mod 25) + ((sectionN mod 40) * 25);
end;

function PE_AREA_CODE(sectionType, groupN: Smallint): Smallint stdcall;
begin
  Result := PE_SECTION_CODE(sectionType, groupN, 0);
end;

function PE_SECTION_TYPE(sectionCode: Smallint): Smallint;
begin
  Result := sectionCode div 1000;
end;

function PE_GROUP_N(sectionCode: Smallint): Smallint;
begin
  Result := sectionCode mod 25;
end;

function PE_SECTION_N(sectionCode: Smallint): Smallint;
begin
  Result := (sectionCode div 25) mod 40;
end;

const
  DLLName = 'CRPE32.DLL';

function  PEPrintReport;                          external DLLName;
function  PEOpenEngine;                           external DLLName;
procedure PECloseEngine;                          external DLLName;
function  PECanCloseEngine;                       external DLLName;
function  PEGetVersion;                           external DLLName;
function  PEOpenPrintJob;                         external DLLName;
function  PEClosePrintJob;                        external DLLName;
function  PEGetReportOptions;                     external DLLName;
function  PESetReportOptions;                     external DLLName;
function  PEStartPrintJob;                        external DLLName;
procedure PECancelPrintJob;                       external DLLName;
function  PEIsPrintJobFinished;                   external DLLName;
function  PEGetJobStatus;                         external DLLName;
function  PEGetErrorCode;                         external DLLName;
function  PEGetErrorText;                         external DLLName;
function  PEGetHandleString;                      external DLLName;
function  PEGetPrintDate;                         external DLLName;
function  PESetPrintDate;                         external DLLName;
function  PESetGroupCondition;                    external DLLName;
function  PEGetNGroups;                           external DLLName;
function  PEGetGroupCondition;                    external DLLName;
function  PEGetGroupOptions;                      external DLLName;
function  PESetGroupOptions;                      external DLLName;
function  PEGetNFormulas;                         external DLLName;
function  PEGetNthFormula;                        external DLLName;
function  PEGetFormula;                           external DLLName;
function  PESetFormula;                           external DLLName;
function  PECheckFormula;                         external DLLName;
function  PEGetSelectionFormula;                  external DLLName;
function  PESetSelectionFormula;                  external DLLName;
function  PECheckSelectionFormula;                external DLLName;
function  PEGetGroupSelectionFormula;             external DLLName;
function  PESetGroupSelectionFormula;             external DLLName;
function  PECheckGroupSelectionFormula;           external DLLName;
function  PEGetNSQLExpressions;                   external DLLName;
function  PEGetNthSQLExpression;                  external DLLName;
function  PEGetSQLExpression;                     external DLLName;
function  PESetSQLExpression;                     external DLLName;
function  PECheckSQLExpression;                   external DLLName;
function  PEGetNSortFields;                       external DLLName;
function  PEGetNthSortField;                      external DLLName;
function  PESetNthSortField;                      external DLLName;
function  PEDeleteNthSortField;                   external DLLName;
function  PEGetNGroupSortFields;                  external DLLName;
function  PEGetNthGroupSortField;                 external DLLName;
function  PESetNthGroupSortField;                 external DLLName;
function  PEDeleteNthGroupSortField;              external DLLName;
function  PEGetNTables;                           external DLLName;
function  PEGetNthTableType;                      external DLLName;
function  PEGetNthTableSessionInfo;               external DLLName;
function  PESetNthTableSessionInfo;               external DLLName;
function  PEGetNthTableLogOnInfo;                 external DLLName;
function  PESetNthTableLogOnInfo;                 external DLLName;
function  PEGetNthTableLocation;                  external DLLName;
function  PESetNthTableLocation;                  external DLLName;
function  PEGetNthTablePrivateInfo;               external DLLName;
function  PESetNthTablePrivateInfo;               external DLLName;
function  PETestNthTableConnectivity;             external DLLName;
function  PELogOnServer;                          external DLLName;
function  PELogOffServer;                         external DLLName;
function  PELogOnSQLServerWithPrivateInfo;        external DLLName;
function  PEGetSQLQuery;                          external DLLName;
function  PESetSQLQuery;                          external DLLName;
function  PEVerifyDatabase;                       external DLLName;
function  PEHasSavedData;                         external DLLName;
function  PEDiscardSavedData;                     external DLLName;
function  PEGetReportTitle;                       external DLLName;
function  PESetReportTitle;                       external DLLName;
function  PEOutputToWindow;                       external DLLName;
function  PEGetWindowOptions;                     external DLLName;
function  PESetWindowOptions;                     external DLLName;
function  PEGetWindowHandle;                      external DLLName;
procedure PECloseWindow;                          external DLLName;
function  PEShowNextPage;                         external DLLName;
function  PEShowFirstPage;                        external DLLName;
function  PEShowPreviousPage;                     external DLLName;
function  PEShowLastPage;                         external DLLName;
function  PEZoomPreviewWindow;                    external DLLName;
function  PEShowPrintControls;                    external DLLName;
function  PEPrintControlsShowing;                 external DLLName;
function  PEPrintWindow;                          external DLLName;
function  PEExportPrintWindow;                    external DLLName;
function  PENextPrintWindowMagnification;         external DLLName;
function  PESelectPrinter;                        external DLLName;
function  PEGetSelectedPrinter;                   external DLLName;
function  PEOutputToPrinter;                      external DLLName;
function  PESetNDetailCopies;                     external DLLName;
function  PEGetNDetailCopies;                     external DLLName;
function  PESetPrintOptions;                      external DLLName;
function  PEGetPrintOptions;                      external DLLName;
function  PEGetExportOptions;                     external DLLName;
function  PEExportTo;                             external DLLName;
function  PESetMargins;                           external DLLName;
function  PEGetMargins;                           external DLLName;
function  PEGetReportSummaryInfo;                 external DLLName;
function  PESetReportSummaryInfo;                 external DLLName;
function  PESetSectionHeight;                     external DLLName;
function  PEGetSectionHeight;       	      	  external DLLName;
function  PESetSectionFormat;                     external DLLName;
function  PEGetSectionFormatFormula;              external DLLName;
function  PESetSectionFormatFormula;              external DLLName;
function  PEGetSectionFormat;                     external DLLName;
function  PESetAreaFormat;                        external DLLName;
function  PEGetAreaFormatFormula;                 external DLLName;
function  PESetAreaFormatFormula;                 external DLLName;
function  PEGetAreaFormat;                        external DLLName;
function  PESetFont;                              external DLLName;
function  PEGetNSubreportsInSection;              external DLLName;
function  PEGetNthSubreportInSection;             external DLLName;
function  PEGetSubreportInfo;                     external DLLName;
function  PEOpenSubreport;                        external DLLName;
function  PECloseSubreport;                       external DLLName;
function  PEGetNParameterFields;                  external DLLName;
function  PEGetNthParameterField;                 external DLLName;
function  PESetNthParameterField;                 external DLLName;
function  PEGetNSections;                         external DLLName;
function  PEGetSectionCode;                       external DLLName;
function  PEGetNPages;                            external DLLName;
function  PEShowNthPage;                          external DLLName;
function  PESetDialogParentWindow;                external DLLName;
function  PEEnableProgressDialog;                 external DLLName;
function  PEGetAllowPromptDialog;                 external DLLName;
function  PESetAllowPromptDialog;                 external DLLName;
function  PEConvertPFInfoToVInfo;                 external DLLName;
function  PEConvertVInfoToPFInfo;                 external DLLName;
function  PEGetNParameterDefaultValues;           external DLLName;
function  PEGetNthParameterDefaultValue;          external DLLName;
function  PESetNthParameterDefaultValue;          external DLLName;
function  PEAddParameterDefaultValue;             external DLLName;
function  PEDeleteNthParameterDefaultValue;       external DLLName;
function  PEGetParameterMinMaxValue;              external DLLName;
function  PESetParameterMinMaxValue;              external DLLName;
function  PEGetParameterValueInfo;                external DLLName;
function  PESetParameterValueInfo;                external DLLName;
function  PEGetNParameterCurrentValues;           external DLLName;
function  PEGetNthParameterCurrentValue;          external DLLName;
function  PEAddParameterCurrentValue;             external DLLName;
function  PEGetNParameterCurrentRanges;           external DLLName;
function  PEGetNthParameterCurrentRange;          external DLLName;
function  PEAddParameterCurrentRange;             external DLLName;
function  PEGetNthParameterType;                  external DLLName;
function  PEClearParameterCurrentValuesAndRanges; external DLLName;
function  PESetFieldMappingType;                  external DLLName;
function  PEGetFieldMappingType;                  external DLLName;
function  PEEnableEvent;                          external DLLName;
function  PEGetEnableEventInfo;                   external DLLName;
function  PESetEventCallback;                     external DLLName;
function  PESetTrackCursorInfo;                   external DLLName;
function  PEGetTrackCursorInfo;                   external DLLName;
{ New Graph Calls }
function  PEGetGraphTypeInfo;                     external DLLName;
function  PESetGraphTypeInfo;                     external DLLName;
function  PEGetGraphTextInfo;                     external DLLName;
function  PESetGraphTextInfo;                     external DLLName;
function  PEGetGraphOptionInfo;                   external DLLName;
function  PESetGraphOptionInfo;                   external DLLName;
function  PEGetGraphFontInfo;                     external DLLName;
function  PESetGraphFontInfo;                     external DLLName;
function  PEGetGraphAxisInfo;                     external DLLName;
function  PESetGraphAxisInfo;                     external DLLName;
{ Other New Calls }
function PEGetNthParameterValueDescription;       external DLLName;
function PESetNthParameterValueDescription;       external DLLName;
function PEGetParameterPickListOption;            external DLLName;
function PESetParameterPickListOption;            external DLLName;
function PECheckNthTableDifferences;              external DLLName;
{ New Calls for SCR8 }
function PEGetNSectionsInArea;                    external DLLName;
function PEGetGraphTextDefaultOption;             external DLLName;
function PESetGraphTextDefaultOption;             external DLLName;
function PEFreeDevMode;                           external DLLName;
function PEGetFormulaSyntax;                      external DLLName;
function PESetFormulaSyntax;                      external DLLName;
function PEReimportSubreport;                     external DLLName;

end.

