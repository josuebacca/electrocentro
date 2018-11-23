/* //////////////////////////////////////////////////////////////////////////////
   Purpose: Display a report generated entirely at run-time using the
          new (for version 8.0) Crystal Reports Creation API functions
          This form takes the selections made in the Selections form and translates
          them into a report that is placed in the Smart Viewer.  The report
          uses the Creation API to manipulate the report into one of the hundreds
          of possible layouts for the Pro Athlete Salary Report.
*/

#include "stdafx.h"
#include "ProSalaries.h"
#include "ReportView.h"

#import "C:\Program Files\Common Files\System\ado\msado15.dll" rename( "EOF", "adoEOF" )

// The ADO connection used in AddDataSource()
ADODB::_ConnectionPtr  Conn1 = NULL;
ADODB::_RecordsetPtr   RS1;

// The constants needed to create the Application and Report Objects COM objects
const CLSID CLSID_Application = {0xb4741fd0,0x45a6,0x11d1,{0xab,0xec,0x00,0xa0,0xc9,0x27,0x4b,0x91}};
const IID IID_IApplication = {0x0bac5cf2,0x44c9,0x11d1,{0xab,0xec,0x00,0xa0,0xc9,0x27,0x4b,0x91}};

const CLSID CLSID_ReportObjects = {0xb4741e60,0x45a6,0x11d1,{0xab,0xec,0x00,0xa0,0xc9,0x27,0x4b,0x91}};
const IID IID_IReportObjects = {0x0bac59b2,0x44c9,0x11d1,{0xab,0xec,0x00,0xa0,0xc9,0x27,0x4b,0x91}};

// The variables that we need to share between Selection.cpp and ReportView.cpp
extern CString g_LeagueAbbrev;
extern CString g_SQLString;
extern BOOL g_GroupByTeam;
extern CString g_SortOrder;
extern long g_TopCount;
extern int g_TypeofCount;

extern BOOL g_TeamName;
extern BOOL g_PlayerName;
extern BOOL g_Salary;
extern BOOL g_YearName;

extern int vcTopN;
extern int vcCompareSalary;

// The application icon used in the two forms
extern HICON g_Icon;

/////////////////////////////////////////////////////////////////////////////
// ReportView dialog

ReportView::ReportView(CWnd* pParent /*=NULL*/)
	: CDialog(ReportView::IDD, pParent)
{
	//{{AFX_DATA_INIT(ReportView)
	//}}AFX_DATA_INIT
}


void ReportView::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(ReportView)
	DDX_Control(pDX, IDC_CRVIEWER1, m_CRViewer1);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(ReportView, CDialog)
	//{{AFX_MSG_MAP(ReportView)
	ON_WM_SHOWWINDOW()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// ReportView message handlers

void ReportView::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	try
	{
		CDialog::OnShowWindow(bShow, nStatus);
		InitReport();
	}
	catch (_com_error& comerr)
	{
		IErrorInfo* pErrorInfo = comerr.ErrorInfo();
		HRESULT hr = comerr.Error();
		if (pErrorInfo)
		{
			BSTR bsDesc = NULL;
			pErrorInfo->GetDescription( &bsDesc );
			_bstr_t sDesc( bsDesc, false );

			MessageBox( sDesc.operator LPCTSTR() );

			pErrorInfo->Release();
		}
		
	}
}

/////////////////////////////////////////////////////////////////////////////
// Reset all module variables before use or re-use

void ReportView::ResetVariables()
{
    m_Application = 0;
    m_Report = 0;
    m_Sections = 0;

	// Variables used in generating the Details
    pPlayerSalariesTitle = 0;
    pSalaryTitle = 0;
    PlayerFormula = 0;
    SalaryFormula = 0;
    TeamFormula = 0;
    YearFormula = 0;
    m_maxLogoHeight = 0;

	// Variables used in the Grouping Procedures
    vOffset = 0;
    LocX = 0;
    groupNum = 0;
    sectionNum = 2;
    pTeamGroupObj = 0;
    pTeamSummary = 0;

	// A dummy variant used throughout the class
    VariantInit (&dummy);
    dummy.vt = VT_EMPTY;
}

/////////////////////////////////////////////////////////////////////////////
// The "central control station" for the report generation.

void ReportView::InitReport()
{
    ResetVariables();
    // Initialize our Application instance
    HRESULT hr = CoCreateInstance(CLSID_Application, NULL, CLSCTX_INPROC_SERVER , IID_IApplication, (void **) &m_Application);
    ASSERT(SUCCEEDED(hr));

    // Create the new report and assign it to the viewer
    m_Report = m_Application->NewReport();
    m_CRViewer1.SetReportSource(m_Report);
    hr = m_Report->get_Sections((ISections**) &m_Sections);
    ASSERT(SUCCEEDED(hr));

	// Do the dynamic report generating for each report section
    DoReportSizing();
    DoReportHeader();
    DoPageHeader();
    DoDetails();
    DoGroupHeader();
    DoFooters();
    AddDataSource();
}

/////////////////////////////////////////////////////////////////////////////
// Do the Report Header Formatting

void ReportView::DoReportHeader()
{
    ISectionPtr pSection = GetReportSection(1);
    // Suppress this section
    pSection->put_Suppress(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// For the Page Header formatting, add the appropriate league logos.  A logo will
// only appear if the user has chosen players from that particular league.

void ReportView::AddPageFieldObjects()
{
    ISectionPtr pSection = GetReportSection(2);
    ICROleObjectPtr pLogo = 0;
    ITextObjectPtr pTextObj = 0;    
    LocX = 0;

    // Get the Application Path
    CString str ("");
    DWORD nBufSize = 100;
    LPTSTR lpStr = str.GetBuffer (nBufSize);
    GetCurrentDirectory(nBufSize, lpStr);
    CString appPath(lpStr);
    
    m_maxLogoHeight = 0;
    long logoHeight = 0;
    CString logoPath = "";
    // Add the NBA Logo only if an NBA team or player is included in the report
    if (g_LeagueAbbrev.Find(_T("NBA")) >= 0)
    {
        logoPath = appPath + _T("\\res\\NBA.bmp");
        pLogo = pSection->AddPictureObject((_bstr_t)logoPath, LocX, 0);
        pLogo->get_Height(&m_maxLogoHeight);
        LocX += 1400;
    }

    // Add the NHL Logo if an NHL team or player is included in the report
    if (g_LeagueAbbrev.Find(_T("NHL")) >= 0)
    {
        logoPath = appPath + _T("\\res\\NHL.bmp");
        pLogo = pSection->AddPictureObject((_bstr_t)logoPath, LocX, 0);
        pLogo->get_Height(&logoHeight);
        if (logoHeight > m_maxLogoHeight)
            m_maxLogoHeight = logoHeight;
        LocX += 1500;
    }

    // Add the MLB Logo if an MLB team or player is included in the report
    if (g_LeagueAbbrev.Find(_T("MLB")) >= 0)
    {
        logoPath = appPath + _T("\\res\\MLB.bmp");
        pLogo = pSection->AddPictureObject((_bstr_t)logoPath, LocX, 0);
        pLogo->get_Height(&logoHeight);
        if (logoHeight > m_maxLogoHeight)
            m_maxLogoHeight = logoHeight;
        LocX += 2000;
    }

    // Add the NFL Logo if an NFL team or player is included in the report
    if (g_LeagueAbbrev.Find(_T("NFL")) >= 0)
    {
        logoPath = appPath + _T("\\res\\NFL.bmp");
        pLogo = pSection->AddPictureObject((_bstr_t)logoPath, LocX, 0);
        pLogo->get_Height(&logoHeight);
        if (logoHeight > m_maxLogoHeight)
            m_maxLogoHeight = logoHeight;
    }
  
	// Add the title to the header
    BSTR objName = SysAllocString(L"Pro Athlete Salary Report");
    pPlayerSalariesTitle = pSection->AddTextObject(objName, 10, 0, dummy);

    SysReAllocString(&objName, L"Title");
    pPlayerSalariesTitle->put_Name(objName);

	// Garbage Collection
    SysFreeString(objName);
    str.ReleaseBuffer ();
}

/////////////////////////////////////////////////////////////////////////////
// In the Page header, move the logos and text around.  Note how the header 
// is formatted differently if one, two, three or four leagues are chosen.
// The Report Creation API functions allow this to easily occur

void ReportView::FormatPageObjects()
{
    ISectionPtr pSection = GetReportSection(2);
    IReportObjectsPtr pReportObjects = 0;
    IReportObjectPtr pReportObject = 0;
    IReportObjectPtr pReportObject2 = 0;

    pPlayerSalariesTitle->put_Width(5000);
    pPlayerSalariesTitle->put_Height(600);
    pPlayerSalariesTitle->put_Top(500);
    pPlayerSalariesTitle->put_Left(1400);

    // Change the Font Settings
    IFontDispPtr pFontDisp = 0;
    IFontPtr pFont = 0;

    // Set the Font Size for the Title
    CY fontSize;
    HRESULT hr = pPlayerSalariesTitle->get_Font((IFontDisp**) &pFontDisp);
    ASSERT(SUCCEEDED(hr));
    hr = pFontDisp->QueryInterface(IID_IFont, (void **) &pFont);
    ASSERT(SUCCEEDED(hr));
    // A back-handed way of setting the font size (16000 roughly translates to 16 point font). 
    pFont->SetRatio(60, 22000);

    // Get the Report Objects in our Report
    hr = pSection->get_ReportObjects((IReportObjects**) &pReportObjects);
    ASSERT(SUCCEEDED(hr));

    VARIANT var;
    VariantInit (&var);
    var.vt = VT_I2;         // short
    long objCount = 0;
    pReportObjects->get_Count(&objCount);

    long objLeft = 0, objWidth = 0;
    switch (objCount)
    {
    case 1:
        // Nothing extra to do
        break;
    case 2:
        // Nothing extra to do
        break;
    case 3:
		// Center the title on the page
        pFont->SetRatio(60, 20000);
        pPlayerSalariesTitle->put_Width(3500);
        pPlayerSalariesTitle->put_Height(1000);
        pPlayerSalariesTitle->put_Left(1500);
        pPlayerSalariesTitle->put_HorAlignment(crHorCenterAlign);

        var.lVal = 2;
        pReportObjects->get_Item(var, (IDispatch**) &pReportObject);
        pReportObject->put_Left(5250);
        break;
    case 4:
        var.lVal = 1;
        pReportObjects->get_Item(var, (IDispatch**) &pReportObject);
        var.lVal = 2;
        pReportObjects->get_Item(var, (IDispatch**) &pReportObject2);

        // Change the property for whichever logo is the second object in the section
        // based on the width and left position of the first logo.
        pReportObject->get_Left(&objLeft);
        pReportObject->get_Width(&objWidth);
        pReportObject2->put_Left(objLeft + objWidth + 1500);

        // Now change the location of the third object in the section
        var.lVal = 3;
        pReportObjects->get_Item(var, (IDispatch**) &pReportObject2);
        pReportObject2->put_Left(5250);
        pPlayerSalariesTitle->put_Top(1400);
        break;
    case 5:
		// All the logos are present, so move the title a bit lower.
        pPlayerSalariesTitle->put_Top(1400);
        break;
    }
    
    // Save the font changes to the Page Title
    fontSize.Lo = 1;
    pFont->put_Size(fontSize);
    pFont->put_Bold(TRUE);
    pPlayerSalariesTitle->putref_Font(pFontDisp);

    VariantClear (&var);
}

/////////////////////////////////////////////////////////////////////////////
void ReportView::DoPageHeader()
{
    AddPageFieldObjects();
    FormatPageObjects();
}

/////////////////////////////////////////////////////////////////////////////
// Add the groups.  The group headers are only added if the user has
// clicked on the "Group by Year and Team" check box in the Selections dialog

void ReportView::AddGroups()
{
    ISectionPtr pSection = 0;
    groupNum = 0;
    sectionNum = 2;

    if (g_GroupByTeam)
    {
        // Add a hidden group by Year
        if (g_YearName)
        {
            m_Report->AddGroup(groupNum, (IDispatch*) YearFormula,
                              crGCAnyValue, crAscendingOrder);
            sectionNum++;
            groupNum++;
            
            // Hide the Year Group, or at least set the height to 1 to make the grouping 
			// look better.  This "invisible" grouping will not occur if the user is 
			// not displaying the League year.
            pSection = GetReportSection(sectionNum);
            pSection->put_Height(1);
        }
        // Add the group by Team        
        if (g_TeamName)
        {
            m_Report->AddGroup(groupNum, (IDispatch*) TeamFormula,
                              crGCAnyValue, crAscendingOrder);
            sectionNum++;
        }
        // Since we're viewing groups, show the nice group display tree
        m_CRViewer1.SetEnableGroupTree(TRUE);
		m_CRViewer1.SetDisplayGroupTree(TRUE);
        vOffset = 770;
    }
    else
    {
		// Ensure that the vertical offset for the field headings never overlaps
		// the logos.
		if (m_maxLogoHeight == 2085)  // NBA logo height
            vOffset = m_maxLogoHeight + 50;
        else
            vOffset = m_maxLogoHeight + 550;
    }
}

/////////////////////////////////////////////////////////////////////////////
// Add the title text objects depending on the user's choices in the Selection dialog.

void ReportView::AddGroupTitles()
{
    ITextObjectPtr pTextObj = 0;
    // Get the Group Section that we're going to work with
    ISectionPtr pSection = GetReportSection(sectionNum);
    
    // Initalize our horizontal offset
	// Since we don't know what fields are being added until runtime, we use an X offset variable
    // to align the selected objects correctly.
    LocX = 10;
    BSTR objName = NULL;

    // Add the actual title text objects depending on the user's choices in the Selection dialog.
    if (g_PlayerName)
    {  
        objName = SysAllocString(L"Player");
        pTextObj = pSection->AddTextObject(objName, LocX, vOffset, dummy);
        SysFreeString(objName);
        LocX += 2400;
    }
    if (g_TeamName)
    {
        objName = SysAllocString(L"Team");
        pTextObj = pSection->AddTextObject(objName, LocX, vOffset, dummy);
        SysFreeString(objName);
        LocX += 2500;
    }
    if (g_Salary)
    {
        objName = SysAllocString(L"Salary");
        pTextObj = pSection->AddTextObject(objName, LocX, vOffset, dummy);
        SysFreeString(objName);
        LocX += 1900;
    }
}

/////////////////////////////////////////////////////////////////////////////
// If we're grouping by Team, add a feature that shows the total displayed 
// payroll for each team.  Also format the Year Object to be a heading if
// the user is grouping by year.

void ReportView::AddGroupFieldObjects()
{
    ISectionPtr pSection = GetReportSection(sectionNum);
    IFieldObjectPtr pFieldObj = 0;

    // Get the Group Section that we're going to work with
    pSection = GetReportSection(sectionNum);

    // The variant will eventually hold our IDispatch reference that we need for adding
	// a summary field object
    VARIANT var;
    VariantInit (&var);

    BSTR objName = NULL;    
    if (g_GroupByTeam)
    {
        var.vt = VT_DISPATCH;
        if (g_TeamName)
        {
			// Add the payroll summary object
            var.pdispVal = TeamFormula;
            pTeamGroupObj = pSection->AddFieldObject(var, 10, 0);

            var.pdispVal = SalaryFormula;
            pTeamSummary = pSection->AddSummaryFieldObject(var, crSTSum, 0, 450, dummy);
            
            objName = SysAllocString(L"Total Payroll Shown:");
            pSalaryTitle = pSection->AddTextObject(objName, 0, 450, dummy);
            SysFreeString(objName);

            // Do some formatting for the payroll total.
            pTeamSummary->put_Left(2350);
            pTeamSummary->put_CurrencySymbolType(crCSTFloatingSymbol);
            pTeamSummary->put_DecimalPlaces(0);
            pTeamGroupObj->put_TextColor(0xFF0000);
            pTeamGroupObj->put_HasDropShadow(TRUE);
        }
    
        if (g_YearName)
        {
            var.pdispVal = YearFormula;
            pFieldObj = pSection->AddFieldObject(var, 0, 0);

            // Do some formatting for the Year that the team payroll is listed for
			// if the user is grouping by year
            pFieldObj->put_Left(4000);
            pFieldObj->put_HorAlignment(crRightAlign);
        }
    }
    else
    {
        if (g_YearName)
        {
			// If the user is not grouping by year, include the year heading along with
			// player, team and salary as a field heading
            objName = SysAllocString(L"Year");
            ITextObjectPtr pTextObj = pSection->AddTextObject(objName, LocX, vOffset, dummy);
            SysFreeString(objName);
        }
    }
}

/////////////////////////////////////////////////////////////////////////////
// Format the field headings to stand out by shading them Teal and making them
// slightly larger

void ReportView::FormatGroupObjects()
{
    ISectionPtr pSection = GetReportSection(sectionNum);
    IReportObjectsPtr pReportObjects = 0;
    IReportObjectPtr pReportObject = 0;

    // Get the Report Objects in our Report
    HRESULT hr = pSection->get_ReportObjects((IReportObjects**) &pReportObjects);
    ASSERT(SUCCEEDED(hr));

    VARIANT var;
    VariantInit (&var);
    var.vt = VT_I2;         // short
    long objCount = 0;
    pReportObjects->get_Count(&objCount);

    IFontDispPtr pFontDisp = 0;
    IFontPtr pFont = 0;
    CY fontSize;

    CRObjectKind objKind;
    for (long i = 1; i <= objCount; i++)
    {
        // Get the ith item
        var.lVal = i;
        pReportObjects->get_Item(var, (IDispatch**) &pReportObject);
        pReportObject->get_Kind(&objKind);

        // Get the name and object kind of the ith item
        BSTR objNameBSTR;
        pReportObject->get_Name(&objNameBSTR);
        CString objName(objNameBSTR);
        SysFreeString(objNameBSTR);

        // We only want to format some objects (the field headings for Player, Team and Salary)
        if (!((objKind != crOleObject) && (objName != "Title")))
            continue;

        // Resize the object
        pReportObject->put_Width(2500);
        pReportObject->put_Height(350);

        if (objKind == crFieldObject)
        {
            IFieldObjectPtr pObj(pReportObject);

            // Set the Font Name and Text Size
            hr = pObj->get_Font((IFontDisp**) &pFontDisp);
            ASSERT(SUCCEEDED(hr));
            hr = pFontDisp->QueryInterface(IID_IFont, (void **) &pFont);
            ASSERT(SUCCEEDED(hr));
            pFont->put_Name(_bstr_t("Arial"));
            // A back-handed way of setting the font size.
            pFont->SetRatio(60, 14000);
            fontSize.Lo = 1;
            pFont->put_Size(fontSize);
            pFont->put_Bold(TRUE);
            pObj->putref_Font(pFontDisp);
        }
        else
        {
            ITextObjectPtr pObj(pReportObject);
        
            // Set the Font Name and Text Size
            hr = pObj->get_Font((IFontDisp**) &pFontDisp);
            ASSERT(SUCCEEDED(hr));
            hr = pFontDisp->QueryInterface(IID_IFont, (void **) &pFont);
            ASSERT(SUCCEEDED(hr));
            pFont->put_Name(_bstr_t("Arial"));
            // A back-handed way of setting the font size.
            pFont->SetRatio(60, 14000);
            fontSize.Lo = 1;
            pFont->put_Size(fontSize);

            if (pObj->GetLeft() != 0)
            {
                pObj->put_TextColor(0x808000);
                pFont->put_Underline(TRUE);
            }
            else
                pFont->put_Bold(TRUE);
            pObj->putref_Font(pFontDisp);
        }
    }

    if ((g_TeamName) && (g_GroupByTeam))
    {
        pSalaryTitle->put_Width(3700);
        pTeamGroupObj->put_Width(3700);

        // Set the Font Name and Text Size
        hr = pTeamGroupObj->get_Font((IFontDisp**) &pFontDisp);
        ASSERT(SUCCEEDED(hr));
        hr = pFontDisp->QueryInterface(IID_IFont, (void **) &pFont);
        ASSERT(SUCCEEDED(hr));
        // A back-handed way of setting the font size.
        pFont->SetRatio(60, 16000);
        fontSize.Lo = 1;
        pFont->put_Size(fontSize);
        pTeamGroupObj->putref_Font(pFontDisp);
    }
}

/////////////////////////////////////////////////////////////////////////////
void ReportView::DoGroupHeader()
{
    AddGroups();
    AddGroupTitles();
    AddGroupFieldObjects();
    FormatGroupObjects();
}

/////////////////////////////////////////////////////////////////////////////
// Add the database fields to the report.  This particular method
// creates unbound fields and binds those fields to the database at
// runtime.  Unbound fields are formula fields in disguise, so the
// techniques for manipulating unbound fields also apply to formulas
// In this procedure we initialize the formulas to be added to the details 
// section in the procedure AddDetailFieldObjects() 

void ReportView::AddDetailFormulas()
{
    ISectionPtr pSection = GetReportSection(sectionNum);
    IFormulaFieldDefinitionsPtr pFormulaFields = 0;

    m_Report->get_FormulaFields((IFormulaFieldDefinitions**) &pFormulaFields);

    // A field must be recurring if you want to group on a field or
    // automatically bind it to a data source.
    // Putting "WhileReadingRecords" into the formula text makes the formula
    // into a recurring field.
    CString Recur = CString(L"WhileReadingRecords;" + 
                           (CString)__toascii(13) + 
                           (CString)__toascii(10));

	// Add the formulas to the formula fields collection
    CString CStrText = Recur + CString(L"Space(10)");
    BSTR strText = CStrText.AllocSysString();
    BSTR objName = SysAllocString(L"Player");
    PlayerFormula = pFormulaFields->Add(objName, strText);
    
    SysReAllocString(&objName, L"Team");
    TeamFormula = pFormulaFields->Add(objName, strText);

    SysReAllocString(&objName, L"Year");
    YearFormula = pFormulaFields->Add(objName, strText);

    SysFreeString(strText);
    CStrText = Recur + CString(L"$0.0");
    strText = CStrText.AllocSysString();
    SysReAllocString(&objName, L"Salary");

    SalaryFormula = pFormulaFields->Add(objName, strText);
    SysFreeString(objName);
}

/////////////////////////////////////////////////////////////////////////////
// Add the field objects to the Report that the user wants to use.  The 
// field objects are based on the formulas created in AddDetailFormulas()

void ReportView::AddDetailFieldObjects()
{
    // Get the Report Details Section
    ISectionPtr pSection = GetReportSection(3);
    IFieldObjectPtr pFieldObj = 0;
   
    LocX = 10;   // Horizontal offset
    VARIANT var;
    VariantInit (&var);
    var.vt = VT_DISPATCH;
    BSTR objName = NULL;

	// Player
    if (g_PlayerName)
    {
        var.pdispVal = PlayerFormula;
        pFieldObj = pSection->AddFieldObject(var, LocX, 0);
        SysReAllocString(&objName, L"Player");
        pFieldObj->put_Name(objName);
        pFieldObj->put_Width(2600);
        LocX += 2100;
    }
    
	// Add the name of the team that the player belongs to
    if (g_TeamName)
    {
        var.pdispVal = TeamFormula;
        pFieldObj = pSection->AddFieldObject(var, LocX, 0);

        SysReAllocString(&objName, L"Team");
        pFieldObj->put_Name(objName);
        pFieldObj->put_Width(2800);
        LocX += 2000;
    }
    
	// Add the year that the player made this amount of money
    if (g_Salary)
    {
        var.pdispVal = SalaryFormula;
        pFieldObj = pSection->AddFieldObject(var, LocX, 0);

        SysReAllocString(&objName, L"Salary");
        pFieldObj->put_Name(objName);
        pFieldObj->put_Width(1500);
        pFieldObj->put_CurrencySymbolType(crCSTFloatingSymbol);
        pFieldObj->put_DecimalPlaces(0);
        LocX += 1600;
    }

	// Add the year that the player made this amount of money
    if (g_YearName)
    {
        var.pdispVal = YearFormula;
        pFieldObj = pSection->AddFieldObject(var, LocX, 0);

        SysReAllocString(&objName, L"Year");
        pFieldObj->put_Name(objName);
        pFieldObj->put_Width(2000);
        pFieldObj->put_HorAlignment(crRightAlign);
        if (g_GroupByTeam) 
            pFieldObj->put_TextColor(0xFFFFFF);
        LocX += 2000;
    }
}

/////////////////////////////////////////////////////////////////////////////
// Format the details to look more visually appealing.  Formatting is done
// by treated the objects in the details section as a group and then formatting
// all those objects in the same manner

void ReportView::FormatFieldObjects()
{
    ISectionPtr pSection = GetReportSection(3);  // Section 3 is the details section
    IReportObjectsPtr pReportObjects = 0;
    IReportObjectPtr pReportObject = 0;
    
    HRESULT hr = CoCreateInstance(CLSID_ReportObjects, NULL, CLSCTX_INPROC_SERVER,
                          IID_IReportObjects, (void **) &pReportObjects);
    pSection->get_ReportObjects((IReportObjects**) &pReportObjects);
    
    VARIANT var;
    VariantInit (&var);
    var.vt = VT_I2;

    long objCount = 0;
    CY fontSize;
    pReportObjects->get_Count(&objCount);
	// Format each field object in the details section
    for (long i = 1; i <= objCount; i++)
    {
        var.lVal = i;
        pReportObjects->get_Item(var, (IDispatch**) &pReportObject);

        IFontPtr pFont;
        IFontDispPtr pFontDisp;
        // Set the Font Name and Text Size
        IFieldObjectPtr pObj(pReportObject);
        hr = pObj->get_Font((IFontDisp**) &pFontDisp);
        ASSERT(SUCCEEDED(hr));
        hr = pFontDisp->QueryInterface(IID_IFont, (void **) &pFont);
        ASSERT(SUCCEEDED(hr));
        pFont->put_Name(_bstr_t("Arial"));
        // A back-handed way of setting the font size.
        pFont->SetRatio(60, 11000);
        fontSize.Lo = 1;
        pFont->put_Size(fontSize);
        pObj->putref_Font(pFontDisp);
        pReportObject->put_Height(250);
    }
}

/////////////////////////////////////////////////////////////////////////////
// Add a sort object that causes the report to internally sort the records
// ascending or descending depending on the user's selection.  If the user
// selects "Show me the highest paid 'n' players", then the report will be
// sort ascending.

void ReportView::AddSortObject()
{
    ISectionPtr pSection = GetReportSection(3);
    ISortFieldsPtr pSortFields = 0;

    m_Report->get_RecordSortFields((ISortFields**) &pSortFields);
	// Add a sorting field to the report sort fields collection
    if (g_SortOrder.Compare("ASC") == 0)
        pSortFields->Add((IFieldDefinitionPtr)SalaryFormula, crAscendingOrder);
    else
        pSortFields->Add((IFieldDefinitionPtr)SalaryFormula, crDescendingOrder);
}

/////////////////////////////////////////////////////////////////////////////
void ReportView::DoDetails()
{    
    AddDetailFormulas();
    AddDetailFieldObjects();
    FormatFieldObjects();
    AddSortObject();
}

/////////////////////////////////////////////////////////////////////////////
// Add a footer to the bottom of each page.  The section number for the
// page varies according to how many grouping fields were added.
//
void ReportView::DoFooters()
{
    ISectionPtr pSection = 0;
    IFieldObjectPtr pPageNumber = 0;
    ILineObjectPtr pLn = 0;
    int curSection;
    long lineRight;
    long PageX;
    VARIANT var;

    VariantInit (&var);
    var.vt = VT_I2;         // short
    
    curSection = 5;
	// The page width for the report is set to a different value if there are no
    // grouping options in the report
    if (g_GroupByTeam)
    {
        if ((g_TeamName) && (g_YearName))
            curSection++;
        lineRight = 6500;
        PageX = 2700;
        var.lVal = curSection;
        m_Sections->get_Item(var, (ISection**) &pSection);
        pSection->put_NewPageAfter(TRUE);
    }
    else
    {
        lineRight = 8500;
        PageX = 4000;
        var.lVal = curSection;
        m_Sections->get_Item(var, (ISection**) &pSection);
    }
    VariantClear(&var);            
    pSection->put_Height(700);
   
	// Add a line directly above the page number
    pLn = pSection->AddLineObject(10, 100, lineRight, 100, dummy);
    pLn->put_LineColor(0xFF0000);
    pLn->put_LineThickness(4);

	// Add the page number to the footer
    // SpecialVar Object is really just a specialized FieldObject
    pPageNumber = pSection->AddSpecialVarFieldObject(crSVTPageNofM, PageX, 0);
    pPageNumber->put_Top(200);

    // Set the Font Name
    IFontDispPtr pFontDisp = 0;
    IFontPtr pFont = 0;

	// Format the page number to be consistent with the rest of the report
    HRESULT hr = pPageNumber->get_Font((IFontDisp**) &pFontDisp);
    ASSERT(SUCCEEDED(hr));
    hr = pFontDisp->QueryInterface(IID_IFont, (void **) &pFont);
    ASSERT(SUCCEEDED(hr));
    pFont->put_Name(_bstr_t("Arial"));
    pPageNumber->putref_Font(pFontDisp);
}

/////////////////////////////////////////////////////////////////////////////
// Change the paper size of the report to better center the report
// fields and group headers.

void ReportView::DoReportSizing()
{
    long pLeftMargin;
    
    if (g_GroupByTeam)
        m_Report->put_PaperSize(crPaperEnvelope14);
    else
        m_Report->put_PaperSize(crPaperEnvelopeB5);
    
    m_Report->get_LeftMargin(&pLeftMargin);
    m_Report->put_RightMargin(pLeftMargin);
    m_Report->put_TopMargin(100);
}

/////////////////////////////////////////////////////////////////////////////
// Add an ADO data source to the report.  The PlayerSalaries.mdb
// database must be located in the same directory as this application
// The ADO data set is generated at run-time and does not persist after
// the report viewing is completed.

void ReportView::AddDataSource()
{
    IDatabasePtr pDatabase = 0;
    IDatabaseTablesPtr pTables = 0;
    IDatabaseTablePtr pTable = 0;
    ASSERT(g_SQLString != "");

    m_Report->get_Database((IDatabase**) &pDatabase);
    pDatabase->get_Tables((IDatabaseTables**) &pTables);
	
    // Get all the records that the user wants -- this runtime dataset is assigned to the report
    HRESULT hr = Conn1.CreateInstance( __uuidof( ADODB::Connection ) );
	ASSERT(SUCCEEDED(hr));

	// Connection String
	_bstr_t bstrAccessConnect ( L"DRIVER={Microsoft Access Driver (*.mdb)};"
		                        L"DBQ=PlayerSalaries.MDB;"                         
			                    L"DefaultDir=;"                             
				                L"UID=admin;PWD=;"
								L"Mode=Read;");
    Conn1->ConnectionString = bstrAccessConnect;
	// An empty string used as a dummy value
	_bstr_t bstrEmpty(L"");

    Conn1->Open( bstrEmpty, bstrEmpty, bstrEmpty, -1 );

    // Open Recordset Object
	RS1.CreateInstance(__uuidof(ADODB::Recordset));
    RS1->Open((_bstr_t)g_SQLString, _variant_t((IDispatch *)Conn1, true), ADODB::adOpenStatic, ADODB::adLockReadOnly,
               ADODB::adCmdText);

    VARIANT var, var2;
    VariantInit(&var);
	VariantInit(&var2);
	var.vt = VT_DISPATCH;
	var.pdispVal = (IDispatch*)Conn1;
	var2.vt = VT_DISPATCH;
	var2.pdispVal = (IDispatch*)RS1->GetActiveCommand();
	
    // Add a new table to the report based on the ADO Record set
	
	hr = pDatabase->AddADOCommand(var, var2);
	ASSERT(SUCCEEDED(hr));

	// Automatically bind the unbound fields to the open data set.  The name of the
    // unbound field must be the same (case insensitive) as the database field.  The
    // name of the unbound field can easily be changed at runtime to allow binding to
    // a different database field.  Alternatively, the FieldName.SetUnboundFieldSource "Database Field Name" 
	// function could be used to manually bind the unbound field.
    hr = m_Report->AutoSetUnboundFieldSource(crBMTName, dummy);
	ASSERT(SUCCEEDED(hr));

    if (g_TopCount > 0)
    {
        // Add the "Top N" or "Greater/Less" than
        pTables->get_Item(1, (IDatabaseTable**) &pTable);
        BSTR tblNameBSTR;
        pTable->get_Name(&tblNameBSTR);
        CString TableName(tblNameBSTR);
        SysFreeString(tblNameBSTR);
        
        if (g_TypeofCount == vcTopN)
        {
			// Get the salary of the Nth Highest or Lowest Paid Player
            RS1->MoveFirst();
			RS1->Move(g_TopCount - 1);
	
			// Ensure we have a valid record
			if (RS1->GetadoEOF())
				RS1->MoveLast();
			
            COleVariant varField;
            VariantInit(&varField);
            varField.vt = VT_CY;
			varField = RS1->GetFields()->GetItem(_T("Salary"))->GetValue();
            if (!RS1->GetadoEOF())
            {
                COleCurrency tmpCur = varField.cyVal;
				CString tmpStr = tmpCur.Format();
                CString tmpSelection;
                if (g_SortOrder == "DESC")
                    tmpSelection = _T("{") + TableName + _T(".Salary} >= $") + tmpStr;
                else
                    tmpSelection = _T("{") + TableName + _T(".Salary} <= $") + tmpStr;
                
                BSTR selectionFormula = tmpSelection.AllocSysString();
                m_Report->put_RecordSelectionFormula(selectionFormula);
                SysFreeString(selectionFormula);
            }
            VariantClear(&varField);
        }
        else if (g_TypeofCount == vcCompareSalary)
        {
            CString tmp;
            char buffer[20];
            _ltoa(g_TopCount, buffer, 10 );
            if (g_SortOrder == "DESC")
                tmp = _T("{") + TableName + _T(".Salary} >= $") + (CString)buffer;
            else
                tmp = _T("{") + TableName + _T(".Salary} <= $") + (CString)buffer;
            BSTR selectionFormula = tmp.AllocSysString();
            m_Report->put_RecordSelectionFormula(selectionFormula);
            SysFreeString(selectionFormula);
        }
    }
	// Finally, view the report in the Smart Viewer!
    m_CRViewer1.ViewReport();
}

/////////////////////////////////////////////////////////////////////////////
// A helper function that gets a requested section

ISectionPtr ReportView::GetReportSection(int sectionNum)
{
    ISectionPtr pSection = 0;
    VARIANT var;
    VariantInit (&var);

    var.vt = VT_I2;
    var.lVal = sectionNum;
    HRESULT hr = m_Sections->get_Item(var, (ISection**) &pSection);
    ASSERT(SUCCEEDED(hr));
    VariantClear (&var);
    return pSection;
}

/////////////////////////////////////////////////////////////////////////////
BOOL ReportView::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
    SetIcon(g_Icon, TRUE);	
	return TRUE;
}

BEGIN_EVENTSINK_MAP(ReportView, CDialog)
    //{{AFX_EVENTSINK_MAP(ReportView)
	ON_EVENT(ReportView, IDC_CRVIEWER1, 411 /* ZoomLevelChanged */, OnZoomLevelChangedCrviewer1, VTS_I2)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

/////////////////////////////////////////////////////////////////////////////
// Force the viewer to always stay zoomed to 100%

void ReportView::OnZoomLevelChangedCrviewer1(short ZoomLevel) 
{
    if (ZoomLevel != 100) 
		m_CRViewer1.Zoom(100);	
}
