/*  //////////////////////////////////////////////////////////////////////////////
  Purpose: Provide an interface that allows the user to compare
           the salaries of any players from any teams from any
           of the major North American sports leagues.
*/

#include "stdafx.h"
#include "ProSalaries.h"
#include "Selection.h"
#include "ReportView.h"
#include "AboutProSalaries.h"

// The variables that we need to share between Selection.cpp and ReportView.cpp
CString g_LeagueAbbrev;
CString g_SQLString;
BOOL g_GroupByTeam;
CString g_SortOrder;
long g_TopCount;
int g_TypeofCount;

BOOL g_TeamName;
BOOL g_PlayerName;
BOOL g_Salary;
BOOL g_YearName;

int vcTopN = 1;
int vcCompareSalary = 2;

// The application icon used in the two forms
extern HICON g_Icon;

/////////////////////////////////////////////////////////////////////////////
// CSelection dialog

CSelection::CSelection(CWnd* pParent /*=NULL*/)
	: CDialog(CSelection::IDD, pParent)
{
	//{{AFX_DATA_INIT(CSelection)
	m_ShowPlayers = 0;
	m_chkPlayer = TRUE;
	m_chkSalary = TRUE;
	m_chkTeam = TRUE;
	m_chkYear = TRUE;
	m_chkGroupTeam = TRUE;
	m_txtLow = 100;
	m_txtHigh = 100;
	m_txtSalaryHigher = 1000000;
	m_txtSalaryLower = 1000000;
	//}}AFX_DATA_INIT
    iLeagueNum = 0;
    iLeagueYearNum = 0;
    hRoot = NULL;
}

CSelection::~CSelection()
{
    m_Database.Close();
}

void CSelection::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CSelection)
	DDX_Control(pDX, IDC_TVWSELECT, m_tvwSelection);
	DDX_Radio(pDX, IDC_SHOW_ALL, m_ShowPlayers);
	DDX_Check(pDX, IDC_CHK_PLAYER, m_chkPlayer);
	DDX_Check(pDX, IDC_CHK_SALARY, m_chkSalary);
	DDX_Check(pDX, IDC_CHK_TEAM, m_chkTeam);
	DDX_Check(pDX, IDC_CHK_YEAR, m_chkYear);
	DDX_Check(pDX, IDC_GROUP_BY_TEAM_YEAR, m_chkGroupTeam);
	DDX_Text(pDX, IDC_TXTLOW, m_txtLow);
	DDV_MinMaxLong(pDX, m_txtLow, 0, 10000);
	DDX_Text(pDX, IDC_TXTHIGH, m_txtHigh);
	DDV_MinMaxLong(pDX, m_txtHigh, 0, 10000);
	DDX_Text(pDX, IDC_TXTSALARYHIGHER, m_txtSalaryHigher);
	DDV_MinMaxLong(pDX, m_txtSalaryHigher, 0, 100000000);
	DDX_Text(pDX, IDC_TXTSALARYLOWER, m_txtSalaryLower);
	DDV_MinMaxLong(pDX, m_txtSalaryLower, 0, 100000000);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CSelection, CDialog)
	//{{AFX_MSG_MAP(CSelection)
	ON_NOTIFY(TVN_ITEMEXPANDING, IDC_TVWSELECT, OnItemexpandingTvwselect)
	ON_WM_SHOWWINDOW()
	ON_BN_CLICKED(IDC_SHOW_REPORT, OnShowReport)
	ON_BN_CLICKED(IDC_ABOUT, OnAbout)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSelection message handlers

/////////////////////////////////////////////////////////////////////////////
// Note that it is only possible to type numbers into the edit boxes.

void CSelection::OnShowReport() 
{
	// Check that we have valid data
    if (UpdateData(TRUE) != 0)
    {
		SaveSettings();
		GenerateSQLStatement();

		// Check that the user actually selected something to report on
		if (g_SQLString.Left(6) == "SELECT")
		{
			ReportView dlg;
			dlg.DoModal();
		}
		else
			MessageBox(_T("Please select a League, Year, Team or Player to use as a data source.")
						, _T("Correction Required"), MB_ICONINFORMATION);
	}
}

/////////////////////////////////////////////////////////////////////////////
void CSelection::OnShowWindow(BOOL bShow, UINT nStatus) 
{
	CDialog::OnShowWindow(bShow, nStatus);

    InitVisuals();
    InitTreeView();
}

/////////////////////////////////////////////////////////////////////////////
// Set the initial default settings for the checkboxes, radio buttons
// and text boxes

void CSelection::InitVisuals()
{
    CTreeCtrl* pCtrl = (CTreeCtrl*) GetDlgItem(IDC_TVWSELECT);
    ASSERT(pCtrl != NULL);

    // Treeview Checkboxes
	m_imageState.Create(IDB_STATE, 14, 1, RGB(255,255,255) );
    ASSERT(pCtrl->GetImageList(TVSIL_NORMAL) == NULL);
    pCtrl->SetImageList( &(m_imageState), TVSIL_STATE );

    // Labels
    CStatic* pEditCtrl = (CStatic*) GetDlgItem(IDC_PLAYERREPORTLABEL);
    ASSERT(pEditCtrl != NULL);
    CFont editFont;
    editFont.CreatePointFont(14, _T("Arial"));
    pEditCtrl->SetFont(&editFont);
    
    pEditCtrl = (CStatic*) GetDlgItem(IDC_FIELDSLABEL);
    ASSERT(pEditCtrl != NULL);
    pEditCtrl->SetFont(&editFont);

}

/////////////////////////////////////////////////////////////////////////////
// Initialize the treeview.  There is one node for each table
// in the PlayerSalaries.mdb database. AddNodes is a helper function
// that adds all the teams in a league as child nodes to the league
// node.  Note that the nodes for the players are only added to the
// treeview when they need to be viewed.  This on demand treeview node
// creation greatly decreases the amount of time it takes to populate
// the treeview.

void CSelection::InitTreeView()
{
    // Get the Application Path
    CString str ("");
    DWORD nBufSize = 100;
    LPTSTR lpStr = str.GetBuffer (nBufSize);
    GetCurrentDirectory(nBufSize, lpStr);

    CString appPath(lpStr);
    m_Database.Open(appPath + (_T("\\PlayerSalaries.mdb")), FALSE, TRUE, _T(""));
    ASSERT(m_Database.GetTableDefCount() > 0);

    str.ReleaseBuffer ();
    CDaoRecordset m_RS(&m_Database);
    
    CString SortKey(_T(" Order By Team, Player\0"));
    // Get all the records from our data tables
    m_RS.Open(dbOpenSnapshot, _T("Select * From NBA9697") + SortKey, dbReadOnly);
    AddNodes(&m_RS, &(_T("NBA\0")), &(_T("1996/1997\0")), &(_T("NBA9697\0")), TRUE);
    m_RS.Close();

    m_RS.Open(dbOpenSnapshot, _T("Select * From NBA9798") + SortKey, dbReadOnly);
    AddNodes(&m_RS, &(_T("NBA\0")), &(_T("1997/1998\0")), &(_T("NBA9798\0")), FALSE);
    m_RS.Close();

    m_RS.Open(dbOpenSnapshot, _T("Select * From NHL9697") + SortKey, dbReadOnly);
    AddNodes(&m_RS, &(_T("NHL\0")), &(_T("1996/1997\0")), &(_T("NHL9697\0")), TRUE);
    m_RS.Close();

    m_RS.Open(dbOpenSnapshot, _T("Select * From NHL9798") + SortKey, dbReadOnly);
    AddNodes(&m_RS, &(_T("NHL\0")), &(_T("1997/1998\0")), &(_T("NHL9798\0")), FALSE);
    m_RS.Close();

    m_RS.Open(dbOpenSnapshot, _T("Select * From NFL1997") + SortKey, dbReadOnly);
    AddNodes(&m_RS, &(_T("NFL\0")), &(_T("1997\0")), &(_T("NFL1997\0")), TRUE);
    m_RS.Close();

    m_RS.Open(dbOpenSnapshot, _T("Select * From MLB1998") + SortKey, dbReadOnly);
    AddNodes(&m_RS, &(_T("MLB\0")), &(_T("1998\0")), &(_T("MLB1998\0")), TRUE);
    m_RS.Close();

    m_RS.Open(dbOpenSnapshot, _T("Select * From MLB1999") + SortKey, dbReadOnly);
    AddNodes(&m_RS, &(_T("MLB\0")), &(_T("1999\0")), &(_T("MLB1999\0")), FALSE);
    m_RS.Close();
}

/////////////////////////////////////////////////////////////////////////////
// This helper function for InitTreeview takes a recordset and adds
// a group of nodes in the format:
//   Root Level = <Name of League>
//   1st Child = <Year>
//   2nd Child = <Teams>
//   3rd Child = <Players>
// The 3rd level nodes are not added until the user chooses to expand
// the team node to view and individually select players

void CSelection::AddNodes(CDaoRecordset* RS, char* League, 
                                  char* Year, char* TableName, BOOL NewLeague)
{
    TVINSERTSTRUCT tvInsert;
    HTREEITEM nodx;
    COleVariant var;

    CTreeCtrl* pCtrl = (CTreeCtrl*) GetDlgItem(IDC_TVWSELECT);
    ASSERT(pCtrl != NULL);

    pCtrl->SetIndent(50);
    // Add Node objects.
    // Root Level = <Name of League>
    // 1st Child = <Year>
    // 2nd Child = <Teams>
    // 3rd Child = <Players>

    // Root node
	// Add a new root node if this is a new league
    if (NewLeague)
    {
        tvInsert.hParent = NULL;
        tvInsert.hInsertAfter = NULL;
        tvInsert.item.mask = TVIF_TEXT | TVIF_PARAM;
        tvInsert.item.pszText = League;
        RootTag[iLeagueNum] = CString(_T("Root")) + League;
        tvInsert.item.lParam = (long)&RootTag[iLeagueNum];
        hRoot = pCtrl->InsertItem(&tvInsert);
        iLeagueNum++;;
    }
    
    // 2nd Level Node
	// Add a child node that represents the year for the salary information
    tvInsert.hParent = hRoot;
    tvInsert.hInsertAfter = NULL;
    tvInsert.item.mask = TVIF_TEXT | TVIF_PARAM;
    tvInsert.item.pszText = Year;
    SecondLevelTag[iLeagueYearNum] = CString(TableName);
    tvInsert.item.lParam = (long)&SecondLevelTag[iLeagueYearNum++];
    HTREEITEM hSecondLevel = pCtrl->InsertItem(&tvInsert);

    pCtrl->EnsureVisible(hSecondLevel);

    VariantInit(&var);
    var.vt = VT_BSTR;
    CString sTmp;

    while (!RS->IsEOF())
    {
        // Team
        RS->GetFieldValue(_T("Team"), var);
        if ((CString(var.pbVal)) != sTmp)
        {
            tvInsert.hParent = hSecondLevel;
            tvInsert.hInsertAfter = NULL;
            // Force the plus sign to appear.
            tvInsert.item.mask = TVIF_TEXT | TVIF_CHILDREN;;
            tvInsert.item.pszText = (char*)_T(var.pbVal);
            tvInsert.item.cChildren = 1;

            nodx = pCtrl->InsertItem(&tvInsert);
            RS->GetFieldValue(_T("Team"), var);
            sTmp = CString(var.pbVal);
        }
        RS->MoveNext();
    }

    VariantClear(&var);
}

/////////////////////////////////////////////////////////////////////////////
// Expand the selected node on-demand.  Since nodes are only
// loaded as the viewer wants to view them, the program can load
// much quicker.  This user interface enhancement is done only
// to make the program appear quicker.

void CSelection::OnItemexpandingTvwselect(NMHDR* pNMHDR, LRESULT* pResult) 
{
    NM_TREEVIEW* pNMTreeView = (NM_TREEVIEW*)pNMHDR;
    CTreeCtrl* pCtrl = (CTreeCtrl*) GetDlgItem(IDC_TVWSELECT);
    ASSERT(pCtrl != NULL);

    TVITEM parent;
    int iNodeLevel = 0;

    // Traverse to the Root node to discover what level of the tree we're on
    HTREEITEM RootNode;
    TVITEMA nodx = pNMTreeView->itemNew;
    HTREEITEM nodh = nodx.hItem;
    RootNode = nodh;
    while (pCtrl->GetParentItem(RootNode) != 0)
    {
        iNodeLevel++;
        RootNode = pCtrl->GetParentItem(RootNode);
    } 

    if (pNMTreeView->action == TVE_EXPAND)
    {

        TVINSERTSTRUCT tvInsert;
        COleVariant var;

        // Ensure this node has not already been expanded
        if ((iNodeLevel == 2) && (pCtrl->GetChildItem(nodh) == 0))
        {
            CString Team = pCtrl->GetItemText(nodh);
            parent.hItem = pCtrl->GetParentItem(nodh);
            parent.mask = TVIF_PARAM | TVIF_HANDLE;
            pCtrl->GetItem(&parent);
            CString* pTableName = (CString*)parent.lParam;

            CString SQL = (CString)_T("Select * From ") + *pTableName + (CString)_T(" Where Team = '")
                            + Team + (CString)_T("' Order By Team, Player");

			// Load the players for this team from the database
            CDaoRecordset RS(&m_Database);
            RS.Open(dbOpenSnapshot, SQL, dbReadOnly);
            tvInsert.hParent = nodh;
            tvInsert.hInsertAfter = NULL;
            tvInsert.item.mask = TVIF_TEXT;
            
            VariantInit(&var);
            var.vt = VT_BSTR;
            RS.MoveFirst();
            while (!RS.IsEOF())
            {        
                RS.GetFieldValue(_T("Player"), var);
                tvInsert.item.pszText = (char*)_T(var.pbVal);
                HTREEITEM playerNode = pCtrl->InsertItem(&tvInsert);
                // Set the child check box to the same state as the parent.                                
                pCtrl->SetItemState(playerNode, pCtrl->GetItemState(nodh, TVIS_STATEIMAGEMASK), TVIS_STATEIMAGEMASK);
                RS.MoveNext();
            }
            RS.Close();
            VariantClear(&var);        
        }
    *pResult = FALSE;;
    }
    else if (pNMTreeView->action == TVE_COLLAPSE)
    {
        *pResult = iNodeLevel < 1 ? TRUE : FALSE;
    }
}

/////////////////////////////////////////////////////////////////////////////
// Recursively check if the treeview node for the team or any node for the players
// on the team is checked.  If zero is returned, then the team needs to be
// excluded in the SQL statement.  A non-zero value means that this team
// needs to be included in the data set for the report

long CSelection::NeedTeam(HTREEITEM nodh)
{
    CTreeCtrl* pCtrl = (CTreeCtrl*) GetDlgItem(IDC_TVWSELECT);

    HTREEITEM curNode;
    long includeTeam = 0;

    if (pCtrl->GetCheck(nodh))
        includeTeam++;
    else
    {
        curNode = pCtrl->GetChildItem(nodh);
        while (!curNode == 0)
        {
            includeTeam += NeedTeam(curNode);
            curNode = pCtrl->GetNextSiblingItem(curNode);
        }
    }
    return includeTeam;
}

/////////////////////////////////////////////////////////////////////////////
// Recursively check through all the child nodes of this league node
// If the return value is zero, then the table doesn//t need to be included
// If the return value is non-zero, then the table needs to be included
// in the SQL statement for the report.

long CSelection::NeedTable(HTREEITEM nodh)
{
    CTreeCtrl* pCtrl = (CTreeCtrl*) GetDlgItem(IDC_TVWSELECT);

    HTREEITEM curNode;
    long includeTable = 0;

    if (pCtrl->GetCheck(nodh))
        includeTable++;
    else
    {
        curNode = pCtrl->GetChildItem(nodh);
        while (!curNode == 0)
        {
            includeTable += NeedTable(curNode);
            curNode = pCtrl->GetNextSiblingItem(curNode);
        }
    }
    return includeTable;
}

/////////////////////////////////////////////////////////////////////////////
// Save the user's settings into a format that can be easily read
// later when generating the report in ReportView.cpp

void CSelection::SaveSettings()
{
    HTREEITEM firstLevelNode;
    HTREEITEM secondLevelNode;
    HTREEITEM teamNode;
    HTREEITEM curNode;
    TVITEM nodx;

    CTreeCtrl* pCtrl = (CTreeCtrl*) GetDlgItem(IDC_TVWSELECT);
    // Clear all the strings
    for (int i = 0; i <= numLeagues; i++)
    {
        m_Players[i] = "";
        m_Teams[i] = ""; 
        m_LeagueName[i] = "";
    }
    g_LeagueAbbrev = "";
    
    // Now save the user's settings to dynamically generate the report
    g_GroupByTeam = m_chkGroupTeam;
    g_TeamName = m_chkTeam;
    g_PlayerName = m_chkPlayer;
    g_Salary = m_chkSalary;
    g_YearName = m_chkYear;

    // Root Level
    int LeagueCount = 0;
    firstLevelNode = pCtrl->GetRootItem();
    while (firstLevelNode != 0)
    {
        // Table Name
        secondLevelNode = pCtrl->GetChildItem(firstLevelNode);
        while (secondLevelNode != 0)
        {
            LeagueCount++;
            // Check if we need to include this table
            if (NeedTable(secondLevelNode) > 0)
            {
                nodx.hItem = secondLevelNode;
                nodx.mask = TVIF_PARAM | TVIF_HANDLE;
                pCtrl->GetItem(&nodx);

                CString *sTmp = (CString*)nodx.lParam;
                m_LeagueName[LeagueCount] = *sTmp;
                g_LeagueAbbrev += *sTmp;

                // Do we need to specially exclude any teams?
                teamNode = pCtrl->GetChildItem(secondLevelNode);
                while (teamNode != 0)
                {
                    if (NeedTeam(teamNode) == 0) 
                        m_Teams[LeagueCount] += _T("|") + pCtrl->GetItemText(teamNode);
                    else if (pCtrl->GetChildItem(teamNode) != 0)
                    {
                        curNode = pCtrl->GetChildItem(teamNode);
                        while (curNode != 0)
                        {
                            if (!pCtrl->GetCheck(curNode))
                                m_Players[LeagueCount] += _T("|") + pCtrl->GetItemText(curNode);
                            curNode = pCtrl->GetNextSiblingItem(curNode);                        
                        }
                    } // end if
                    teamNode = pCtrl->GetNextSiblingItem(teamNode); 
                } // while
            } // end if
            secondLevelNode = pCtrl->GetNextSiblingItem(secondLevelNode);
        }
        firstLevelNode = pCtrl->GetNextSiblingItem(firstLevelNode);
    }
}

/////////////////////////////////////////////////////////////////////////////
// Now that the user's preferences have been saved in SaveSettings(),
// generate the SQL statement that will be used in the report.
// Also save the other relevant settings, such as the number of leagues
// that are being included and whether the user wants ascending or
// descending sorting.  This information is used to rearrange the
// report headings and details in frmMain

void CSelection::GenerateSQLStatement()
{
    int viewValue = m_ShowPlayers;
    // Add parameter to only select the topN players from each league
    switch (viewValue)
    {
    case 1:  // Highest n paid
        g_SortOrder = "DESC";
        g_TopCount = m_txtHigh;
        g_TypeofCount = vcTopN;
        break;
    case 2: // Lowest n paid
        g_TopCount = m_txtLow;
        g_SortOrder = "ASC";
        g_TypeofCount = vcTopN;
        break;
    case 3: // Salary Higher than $n
        g_TopCount = m_txtSalaryHigher;
        g_SortOrder = "DESC";
        g_TypeofCount = vcCompareSalary;
        break;
    case 4: // Salary Lower than $n
        g_TopCount = m_txtSalaryLower;
        g_SortOrder = "ASC";
        g_TypeofCount = vcCompareSalary;
        break;
    default:        // optAll.Value = True
        g_SortOrder = "DESC";
        g_TopCount = 0;
        g_TypeofCount = vcTopN;
    }

    CString strTemp = "";

    // Select the tables
    for (int i = 0; i <= numLeagues; i++)
    {
        if (m_LeagueName[i] != "")
        {
            if (strTemp != "")
                strTemp += _T(" UNION ");
            strTemp = strTemp + _T("SELECT Team, Salary, Year, Player FROM ") + m_LeagueName[i];
            
            // Exclude individual teams
            if (m_Teams[i] != "")
            {
                m_Teams[i].Delete(0,1); // Delete the first "|"
                m_Teams[i].Replace(_T("|"), _T("' AND Team <> '"));
                strTemp = strTemp + _T(" WHERE (Team <> '")
                                  + m_Teams[i]
                                  + _T("')");
            }
    
            // Exclude individual players
            if (m_Players[i] != "")
            {
                m_Players[i].Delete(0,1); // Delete the first "|"
                m_Players[i].Replace(_T("|"), _T("' AND Player <> '"));
                if (strTemp.Right(1) == _T(")"))
                    strTemp = strTemp + _T(" AND (Player <> '")
                                      + m_Players[i]
                                      + _T("')");
                else
                    strTemp = strTemp + _T(" WHERE (Player <> '")
                                      + m_Players[i]
                                      + _T("')");
            }
        }
    }

	// Add the sorting
    strTemp += _T(" Order by Salary ") + g_SortOrder;
    g_SQLString = strTemp;
}

/////////////////////////////////////////////////////////////////////////////
BOOL CSelection::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
    SetIcon(g_Icon, TRUE);
	return TRUE;  
}

/////////////////////////////////////////////////////////////////////////////
// Show the About box

void CSelection::OnAbout() 
{
	CAboutProSalaries dlg;
	dlg.DoModal();
}

