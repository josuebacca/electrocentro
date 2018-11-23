Attribute VB_Name = "crwrap"
'
'               Visual Basic Declarations of crwrap32.DLL
'               =====================================
'
'       File:         crwrap.BAS
'
'       Author:       Seagate Software Information Management Group, Inc.
'
'
'       Language:     Visual Basic for Windows
'
'       Copyright (c) 1992 - 1999 Seagate Software Information Management Group, Inc.
'
'       Revisions:
'       SM- Nov 99, removed obsolete exporting formats
'           Excel 2.1
'            Global Const crUXFXls2Type% = 0
'           Excel 3.0
'            Global Const crUXFXls3Type% = 1
'           Excel 4.0
'            Global Const crUXFXls4Type% = 2
'           HTML 3.0
'            Global Const crUXFHTML3Type% = 0   ' Draft HTML 3.0 tags
'
'
'================================================================================
' Constants and Function Definitions for Exporting
'================================================================================
'******************
'** Format Types **
'******************

'Separated Values           (DLL: "uxfsepv.dll")
Global Const crUXFCharSeparatedType% = 100
Global Const crUXFCommaSeparatedType% = 200
Global Const crUXFTabSeparatedType% = 300

'Data Interchange Format    (DLL: "uxfdif.dll")
Global Const crUXFDIFType% = 400

'Record Style Format        (DLL: "uxfrec.dll")
Global Const crUXFRecordType% = 500


'Crystal Report             (DLL: "uxfcr.dll")
Global Const crUXFCrystalReportType% = 0

'RTF                        (DLL: "uxfrtf.dll")
Global Const crUXFRichTextFormatType% = 0

'Text                       (DLL: "uxftext.dll")
Global Const crUXFTextType% = 0
Global Const crUXFTabbedTextType% = 1
Global Const crUXFPaginatedTextType% = 600

'Lotus                      (DLL: "uxfwks.dll")
Global Const crUXFLotusWksType% = 0 ' Lotus 123 (WKS)
Global Const crUXFLotusWk1Type% = 1 ' Lotus 123(WK1)
Global Const crUXFLotusWk3Type% = 2

'Word for Windows           (DLL: "uxfwordw.dll")
Global Const crUXFWordWinType% = 0

'Excel                      (DLL: "uxfxls.dll")
Global Const crUXFXls5Type% = 3
Global Const crUXFXls5TypeTab% = 700

'Report Definition          (DLL: "uxfrdef.dll")
Global Const crUXFReportDefinitionType = 0

#If Win16 Then
'Word for DOS & WordPerfect (DLL: "uxfdoc.dll")
Global Const crUXFWordDosType% = 1
Global Const crUXFWordPerfectType% = 2

'Quattro Pro                (DLL: "uxfqp.dll")
Global Const crUXFQP5Type% = 0

#End If

'****************
'** HTML Types **
'****************

'HTML
Global Const crUXFExplorer2Type% = 1             ' Include MS Explorer 2.0 tags
Global Const crUXFNetscape2Type% = 2             ' Include Netscape 2.0 tags
Global Const crUXFHTML32ExtType% = 1             ' HTML 3.2 tags + bg color extensions
Global Const crUXFHTML32StdType% = 2             ' HTML 3.2 tags


'DateTime structure
Type crDateTime
    crDate As Long
    crTime As Long
End Type


'Devmode structure
Type crDEVMODE
    dmDriverVersion As Integer  ' printer driver version number (usually not required)
#If Win16 Then                  ' add padding so it aligns the same way under both 16-bit and 32-bit environments
    pad1 As Integer
#End If
    dmFields As Long           'flags indicating fields to modify (required)
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
End Type

'/* field selection bits */
Global Const DM_ORIENTATION = &H1
Global Const DM_PAPERSIZE = &H2
Global Const DM_PAPERLENGTH = &H4
Global Const DM_PAPERWIDTH = &H8
Global Const DM_SCALE = &H10
Global Const DM_COPIES = &H100
Global Const DM_DEFAULTSOURCE = &H200
Global Const DM_PRINTQUALITY = &H400
Global Const DM_COLOR = &H800
Global Const DM_DUPLEX = &H1000
Global Const DM_YRESOLUTION = &H2000
Global Const DM_TTOPTION = &H4000

'/* orientation selections */
Global Const DMORIENT_PORTRAIT = 1
Global Const DMORIENT_LANDSCAPE = 2

'/* paper selections */
'/*  Warning: The PostScript driver mistakingly uses DMPAPER_ values between
' *  50 and 56.  Don't use this range when defining new paper sizes.
' */
Global Const DMPAPER_FIRST = 1
Global Const DMPAPER_LETTER = 1               '/* Letter 8 1/2 x 11 in               */
Global Const DMPAPER_LETTERSMALL = 2          '/* Letter Small 8 1/2 x 11 in         */
Global Const DMPAPER_TABLOID = 3              '/* Tabloid 11 x 17 in                 */
Global Const DMPAPER_LEDGER = 4               '/* Ledger 17 x 11 in                  */
Global Const DMPAPER_LEGAL = 5                '/* Legal 8 1/2 x 14 in                */
Global Const DMPAPER_STATEMENT = 6            '/* Statement 5 1/2 x 8 1/2 in         */
Global Const DMPAPER_EXECUTIVE = 7            '/* Executive 7 1/4 x 10 1/2 in        */
Global Const DMPAPER_A3 = 8                   '/* A3 297 x 420 mm                    */
Global Const DMPAPER_A4 = 9                   '/* A4 210 x 297 mm                    */
Global Const DMPAPER_A4SMALL = 10             '/* A4 Small 210 x 297 mm              */
Global Const DMPAPER_A5 = 11                  '/* A5 148 x 210 mm                    */
Global Const DMPAPER_B4 = 12                  '/* B4 250 x 354                       */
Global Const DMPAPER_B5 = 13                  '/* B5 182 x 257 mm                    */
Global Const DMPAPER_FOLIO = 14               '/* Folio 8 1/2 x 13 in                */
Global Const DMPAPER_QUARTO = 15              '/* Quarto 215 x 275 mm                */
Global Const DMPAPER_10X14 = 16               '/* 10x14 in                           */
Global Const DMPAPER_11X17 = 17               '/* 11x17 in                           */
Global Const DMPAPER_NOTE = 18                '/* Note 8 1/2 x 11 in                 */
Global Const DMPAPER_ENV_9 = 19               '/* Envelope #9 3 7/8 x 8 7/8          */
Global Const DMPAPER_ENV_10 = 20              '/* Envelope #10 4 1/8 x 9 1/2         */
Global Const DMPAPER_ENV_11 = 21              '/* Envelope #11 4 1/2 x 10 3/8        */
Global Const DMPAPER_ENV_12 = 22              '/* Envelope #12 4 \276 x 11           */
Global Const DMPAPER_ENV_14 = 23              '/* Envelope #14 5 x 11 1/2            */
Global Const DMPAPER_CSHEET = 24              '/* C size sheet                       */
Global Const DMPAPER_DSHEET = 25              '/* D size sheet                       */
Global Const DMPAPER_ESHEET = 26              '/* E size sheet                       */
Global Const DMPAPER_ENV_DL = 27              '/* Envelope DL 110 x 220mm            */
Global Const DMPAPER_ENV_C5 = 28              '/* Envelope C5 162 x 229 mm           */
Global Const DMPAPER_ENV_C3 = 29              '/* Envelope C3  324 x 458 mm          */
Global Const DMPAPER_ENV_C4 = 30              '/* Envelope C4  229 x 324 mm          */
Global Const DMPAPER_ENV_C6 = 31              '/* Envelope C6  114 x 162 mm          */
Global Const DMPAPER_ENV_C65 = 32             '/* Envelope C65 114 x 229 mm          */
Global Const DMPAPER_ENV_B4 = 33              '/* Envelope B4  250 x 353 mm          */
Global Const DMPAPER_ENV_B5 = 34              '/* Envelope B5  176 x 250 mm          */
Global Const DMPAPER_ENV_B6 = 35              '/* Envelope B6  176 x 125 mm          */
Global Const DMPAPER_ENV_ITALY = 36           '/* Envelope 110 x 230 mm              */
Global Const DMPAPER_ENV_MONARCH = 37         '/* Envelope Monarch 3.875 x 7.5 in    */
Global Const DMPAPER_ENV_PERSONAL = 38        '/* 6 3/4 Envelope 3 5/8 x 6 1/2 in    */
Global Const DMPAPER_FANFOLD_US = 39          '/* US Std Fanfold 14 7/8 x 11 in      */
Global Const DMPAPER_FANFOLD_STD_GERMAN = 40  '/* German Std Fanfold 8 1/2 x 12 in   */
Global Const DMPAPER_FANFOLD_LGL_GERMAN = 41  '/* German Legal Fanfold 8 1/2 x 13 in */

Global Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN

Global Const DMPAPER_USER = 256

'/* bin selections */
Global Const DMBIN_FIRST = 1
Global Const DMBIN_UPPER = 1
Global Const DMBIN_ONLYONE = 1
Global Const DMBIN_LOWER = 2
Global Const DMBIN_MIDDLE = 3
Global Const DMBIN_MANUAL = 4
Global Const DMBIN_ENVELOPE = 5
Global Const DMBIN_ENVMANUAL = 6
Global Const DMBIN_AUTO = 7
Global Const DMBIN_TRACTOR = 8
Global Const DMBIN_SMALLFMT = 9
Global Const DMBIN_LARGEFMT = 10
Global Const DMBIN_LARGECAPACITY = 11
Global Const DMBIN_CASSETTE = 14
Global Const DMBIN_LAST = DMBIN_CASSETTE

Global Const DMBIN_USER = 256             '/* device specific bins start here */

'/* print qualities */
Global Const DMRES_DRAFT = -1
Global Const DMRES_LOW = -2
Global Const DMRES_MEDIUM = -3
Global Const DMRES_HIGH = -4

'/* color enable/disable for color printers */
Global Const DMCOLOR_MONOCHROME = 1
Global Const DMCOLOR_COLOR = 2

'/* duplex enable */
Global Const DMDUP_SIMPLEX = 1
Global Const DMDUP_VERTICAL = 2
Global Const DMDUP_HORIZONTAL = 3

'/* TrueType options */
Global Const DMTT_BITMAP = 1          '/* print TT fonts as graphics */
Global Const DMTT_DOWNLOAD = 2        '/* download TT fonts as soft fonts */
Global Const DMTT_SUBDEV = 3          '/* substitute device fonts for TT fonts */



'================================================================================
' Type Declarations for 4-byte aligned functions
'================================================================================
#If Win32 Then
    Type PEJobInfo4
        StructSize As Integer  ' initialize to PE_SIZEOF_JOB_INFO
    
        NumRecordsRead As Long
        NumRecordsSelected As Long
        NumRecordsPrinted As Long
    
        DisplayPageN As Integer
        LatestPageN As Integer
        StartPageN As Integer
    
        PrintEnded As Long
    End Type
    Global Const PE_SIZEOF_JOB_INFO4 = 28
    
    Type PETableType4
        StructSize As Integer   ' initialize to # bytes in PETableType
    
        DLLName As String * PE_DLL_NAME_LEN
        DescriptiveName  As String * PE_FULL_NAME_LEN
    
        DBType As Integer
    End Type
    Global Const PE_SIZEOF_TABLE_TYPE4 = 644

    Type PELogOnInfo4
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
    Global Const PE_SIZEOF_LOGON_INFO4 = 514  ' # bytes in PELogOnInfo


    Type PESessionInfo4
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
    Global Const PE_SIZEOF_SESSION_INFO4 = 262  ' # bytes in PESessionInfo


    Type PEGraphOptions4
        StructSize     As Integer  ' initialize to # bytes in PEGraphOptions
        GraphMaxValue  As Double
        GraphMinValue  As Double
        ShowDataValue  As Long  ' Show data values on risers.
        ShowGridLine   As Long
        VerticalBars   As Long
        ShowLegend     As Long
        FontFaceName   As String * PE_GRAPH_TEXT_LEN
    End Type
    Global Const PE_SIZEOF_GRAPH_OPTIONS4 = PE_WORD_LEN + 2 * 8 + 4 * 4 + PE_GRAPH_TEXT_LEN

#End If

'================================================================================
' Function Declarations
'================================================================================
#If Win16 Then
    '** new CRPE wrapper Functions
    Declare Function crPELogOnServer Lib "crwrap16.dll" (ByVal DLLName$, ByVal ServerName$, ByVal DBName$, ByVal UserID$, ByVal Password$) As Integer
    Declare Function crPELogOffServer Lib "crwrap16.dll" (ByVal DLLName$, ByVal ServerName$, ByVal DBName$, ByVal UserID$, ByVal Password$) As Integer
    Declare Function crPESetNthTableLogOnInfo Lib "crwrap16.dll" (ByVal printJob%, ByVal TableN%, ByVal ServerName$, ByVal DBName$, ByVal UserID$, ByVal Password$, ByVal PropagateAcrossTables%) As Integer
    Declare Function crPEGetNthTableLogOnInfo Lib "crwrap16.dll" Alias "crvbPEGetNthTableLogOnInfo" (ByVal printJob%, ByVal TableN%, ByRef ServerName$, ByRef DBName$, ByRef UserID$) As Integer
    Declare Function crPESetNthTableLocation Lib "crwrap16.dll" (ByVal printJob%, ByVal TableN%, ByVal tableLocation$) As Integer
    Declare Function crPEGetNthTableLocation Lib "crwrap16.dll" Alias "crvbPEGetNthTableLocation" (ByVal printJob%, ByVal TableN%, ByRef tableLocation$) As Integer
    Declare Function crPEGetSelectedPrinter Lib "crwrap16.dll" Alias "crvbPEGetSelectedPrinter" (ByVal printJob%, ByRef driverName$, ByRef PrinterName$, ByRef PortName$, crmode As crDEVMODE) As Integer
    Declare Function crPESelectPrinter Lib "crwrap16.dll" (ByVal printJob%, ByVal driverName$, ByVal PrinterName$, ByVal PortName$, crmode As crDEVMODE) As Integer
    Declare Function crPESetNthTableSessionInfo Lib "crwrap16.dll" (ByVal printJob%, ByVal TableN%, ByVal UserID$, ByVal Password$, ByVal SessionHandle As Long, ByVal PropagateAcrossTables%) As Integer
    Declare Function crPEGetNthTableSessionInfo Lib "crwrap16.dll" Alias "crvbPEGetNthTableSessionInfo" (ByVal printJob%, ByVal TableN%, ByRef UserID$, ByRef SessionHandle As Long) As Integer
    Declare Function crPESetGraphOptions Lib "crwrap16.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, ByVal GraphMaxValue As Double, ByVal GraphMinValue As Double, ByVal ShowDataValue%, ByVal ShowGridLine%, ByVal VerticalBars%, ByVal ShowLegend%, ByRef FontFaceName$) As Integer
    Declare Function crPEGetGraphOptions Lib "crwrap16.dll" Alias "crvbPEGetGraphOptions" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, ByRef GraphMaxValue As Double, ByRef GraphMinValue As Double, ByRef ShowDataValue%, ByRef ShowGridLine%, ByRef VerticalBars%, ByRef ShowLegend%, ByRef FontFaceName$) As Integer
    Declare Function crPEGetNthParameterType Lib "crwrap16.dll" (ByVal printJob%, ByVal parameterN%) As Integer
    Declare Function crPEGetNthParameterField Lib "crwrap16.dll" Alias "crvbPEGetNthParameterField" (ByVal printJob%, ByVal parameterN%, ByRef valueType%, ByRef DefaultValueSet%, ByRef CurrentValueSet%, ByRef Name$, ByRef Prompt$, ByRef DefaultValue As Any, ByRef currentValue As Any) As Integer
    Declare Function crPESetNthParameterField Lib "crwrap16.dll" (ByVal printJob%, ByVal parameterN%, ByVal valueType%, ByVal DefaultValueSet%, ByVal CurrentValueSet%, ByVal Name$, ByVal Prompt$, ByRef DefaultValue As Any, ByRef currentValue As Any) As Integer

    Declare Function crvbHandleToBStr Lib "crwrap16.dll" (ByRef BString$, ByVal strHandle%, ByVal strLength%) As Integer
    Declare Function crPEYearMonthDayToDate Lib "crwrap16.dll" (ByVal nYear%, ByVal nMonth%, ByVal nDay%) As Long
    Declare Function crPEDateToYearMonthDay Lib "crwrap16.dll" (ByVal nDate As Long, ByRef nYear%, ByRef nMonth%, ByRef nDay%) As Integer
    Declare Function crPEHourMinuteSecondToTime Lib "crwrap16.dll" (ByVal nHour%, ByVal nMinute%, ByVal nSecond%) As Long
    Declare Function crTimeToHourMinuteSecond Lib "crwrap16.dll" (ByVal nTime As Long, ByRef nHour%, ByRef nMinute%, ByRef nSecond%) As Integer

    '** Export Functions **
    Declare Function crPEExportToApp Lib "crwrap16.dll" (ByVal printJob%, ByVal fileName$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToDisk Lib "crwrap16.dll" (ByVal printJob%, ByVal fileName$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToMapi Lib "crwrap16.dll" (ByVal printJob%, ByVal toList$, ByVal ccList$, ByVal subject$, ByVal message$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToExch Lib "crwrap16.dll" (ByVal printJob%, ByVal profile$, ByVal Password$, ByVal folderPath$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToODBC Lib "crwrap16.dll" (ByVal printJob%, ByVal dataSourceName$, ByVal dataSourceUserID$, ByVal dataSourcePassword$, ByVal exportTableName$) As Integer
    Declare Function crPEExportToHTML Lib "crwrap16.dll" (ByVal printJob%, ByVal htmlType%, ByVal fileName$) As Integer
    Declare Function crPEExportToVIM Lib "crwrap16.dll" (ByVal printJob%, ByVal toList$, ByVal ccList$, ByVal bccList$, ByVal subject$, ByVal message$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToNotes Lib "crwrap16.dll" (ByVal printJob%, ByVal DBName$, ByVal FormName$, ByVal comments$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer

#ElseIf Win32 Then
    '** new CRPE wrapper Functions
    Declare Function crPELogOnServer Lib "crwrap32.dll" (ByVal DLLName$, ByVal ServerName$, ByVal DBName$, ByVal UserID$, ByVal Password$) As Integer
    Declare Function crPELogOffServer Lib "crwrap32.dll" (ByVal DLLName$, ByVal ServerName$, ByVal DBName$, ByVal UserID$, ByVal Password$) As Integer
    Declare Function crPESetNthTableLogOnInfo Lib "crwrap32.dll" (ByVal printJob%, ByVal TableN%, ByVal ServerName$, ByVal DBName$, ByVal UserID$, ByVal Password$, ByVal PropagateAcrossTables As Long) As Integer
    Declare Function crPEGetNthTableLogOnInfo Lib "crwrap32.dll" Alias "crvbPEGetNthTableLogOnInfo" (ByVal printJob%, ByVal TableN%, ByRef ServerName$, ByRef DBName$, ByRef UserID$) As Integer
    Declare Function crPESetNthTableLocation Lib "crwrap32.dll" (ByVal printJob%, ByVal TableN%, ByVal tableLocation$) As Integer
    Declare Function crPEGetNthTableLocation Lib "crwrap32.dll" Alias "crvbPEGetNthTableLocation" (ByVal printJob%, ByVal TableN%, ByRef tableLocation$) As Integer
    Declare Function crPEGetSelectedPrinter Lib "crwrap32.dll" Alias "crvbPEGetSelectedPrinter" (ByVal printJob%, ByRef driverName$, ByRef PrinterName$, ByRef PortName$, crmode As crDEVMODE) As Integer
    Declare Function crPESelectPrinter Lib "crwrap32.dll" (ByVal printJob%, ByVal driverName$, ByVal PrinterName$, ByVal PortName$, crmode As crDEVMODE) As Integer
    Declare Function crPESetNthTableSessionInfo Lib "crwrap32.dll" (ByVal printJob%, ByVal TableN%, ByVal UserID$, ByVal Password$, ByVal SessionHandle As Long, ByVal PropagateAcrossTables As Long) As Integer
    Declare Function crPEGetNthTableSessionInfo Lib "crwrap32.dll" Alias "crvbPEGetNthTableSessionInfo" (ByVal printJob%, ByVal TableN%, ByRef UserID$, ByRef SessionHandle As Long) As Integer
    Declare Function crPESetGraphOptions Lib "crwrap32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, ByVal GraphMaxValue As Double, ByVal GraphMinValue As Double, ByVal ShowDataValue As Long, ByVal ShowGridLine As Long, ByVal VerticalBars As Long, ByVal ShowLegend As Long, ByVal FontFaceName$) As Integer
    Declare Function crPEGetGraphOptions Lib "crwrap32.dll" Alias "crvbPEGetGraphOptions" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, ByRef GraphMaxValue As Double, ByRef GraphMinValue As Double, ByRef ShowDataValue As Long, ByRef ShowGridLine As Long, ByRef VerticalBars As Long, ByRef ShowLegend As Long, ByRef FontFaceName$) As Integer
    Declare Function crPEGetNthParameterType Lib "crwrap32.dll" (ByVal printJob%, ByVal parameterN%) As Integer
    Declare Function crPEGetNthParameterField Lib "crwrap32.dll" Alias "crvbPEGetNthParameterField" (ByVal printJob%, ByVal parameterN%, ByRef valueType%, ByRef DefaultValueSet%, ByRef CurrentValueSet%, ByRef Name$, ByRef Prompt$, ByRef DefaultValue As Any, ByRef currentValue As Any) As Integer
    Declare Function crPESetNthParameterField Lib "crwrap32.dll" (ByVal printJob%, ByVal parameterN%, ByVal valueType%, ByVal DefaultValueSet%, ByVal CurrentValueSet%, ByVal Name$, ByVal Prompt$, ByRef DefaultValue As Any, ByRef currentValue As Any) As Integer

    Declare Function crvbHandleToBStr Lib "crwrap32.dll" (ByRef BString$, ByVal strHandle As Long, ByVal strLength%) As Integer
    Declare Function crPEYearMonthDayToDate Lib "crwrap32.dll" (ByVal nYear%, ByVal nMonth%, ByVal nDay%) As Long
    Declare Function crPEDateToYearMonthDay Lib "crwrap32.dll" (ByVal nDate As Long, ByRef nYear%, ByRef nMonth%, ByRef nDay%) As Integer
    Declare Function crPEHourMinuteSecondToTime Lib "crwrap32.dll" (ByVal nHour%, ByVal nMinute%, ByVal nSecond%) As Long
    Declare Function crTimeToHourMinuteSecond Lib "crwrap32.dll" (ByVal nTime As Long, ByRef nHour%, ByRef nMinute%, ByRef nSecond%) As Integer

    '** Export Functions **
    Declare Function crPEExportToApp Lib "crwrap32.dll" (ByVal printJob%, ByVal fileName$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToDisk Lib "crwrap32.dll" (ByVal printJob%, ByVal fileName$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToMapi Lib "crwrap32.dll" (ByVal printJob%, ByVal toList$, ByVal ccList$, ByVal subject$, ByVal message$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToExch Lib "crwrap32.dll" (ByVal printJob%, ByVal profile$, ByVal Password$, ByVal folderPath$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToODBC Lib "crwrap32.dll" (ByVal printJob%, ByVal dataSourceName$, ByVal dataSourceUserID$, ByVal dataSourcePassword$, ByVal exportTableName$) As Integer
    Declare Function crPEExportToHTML Lib "crwrap32.dll" (ByVal printJob%, ByVal htmlType%, ByVal fileName$) As Integer
    Declare Function crPEExportToVIM Lib "crwrap32.dll" (ByVal printJob%, ByVal toList$, ByVal ccList$, ByVal bccList$, ByVal subject$, ByVal message$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
    Declare Function crPEExportToNotes Lib "crwrap32.dll" (ByVal printJob%, ByVal DBName$, ByVal FormName$, ByVal comments$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer
  
    '** Functions which accept VB's native 4-byte aligned "Types"
    Declare Function PEGetJobStatus4 Lib "crwrap32.dll" (ByVal printJob%, JobInfo As PEJobInfo4) As Integer
    Declare Function PEGetNthTableType4 Lib "crwrap32.dll" (ByVal printJob%, ByVal TableN%, TableType As PETableType4) As Integer
    Declare Function PELogOnServer4 Lib "crwrap32.dll" (ByVal DLLName$, LogOnInfo As PELogOnInfo4) As Integer
    Declare Function PELogOffServer4 Lib "crwrap32.dll" (ByVal DLLName$, LogOnInfo As PELogOnInfo4) As Integer
    Declare Function PESetNthTableSessionInfo4 Lib "crwrap32.dll" (ByVal printJob%, ByVal TableN%, SessionInfo As PESessionInfo4, ByVal PropagateAcrossTables As Long) As Integer
    Declare Function PEGetNthTableSessionInfo4 Lib "crwrap32.dll" (ByVal printJob%, ByVal TableN%, SessionInfo As PESessionInfo4) As Integer
    Declare Function PESetGraphOptions4 Lib "crwrap32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphOptions As PEGraphOptions4) As Integer
    Declare Function PEGetGraphOptions4 Lib "crwrap32.dll" (ByVal printJob%, ByVal sectionCode%, ByVal GraphN%, GraphOptions As PEGraphOptions4) As Integer

#End If

