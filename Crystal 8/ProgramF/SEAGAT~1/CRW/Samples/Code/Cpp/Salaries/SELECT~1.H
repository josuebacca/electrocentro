#if !defined(AFX_SELECTION_H__12007A5C_4C47_11D3_A685_00902741EC66__INCLUDED_)
#define AFX_SELECTION_H__12007A5C_4C47_11D3_A685_00902741EC66__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "TreeCtrlCheckBox.h"
#include <afxdao.h>
const int numLeagues = 7;  // The number of leagues that are available in the database

/////////////////////////////////////////////////////////////////////////////
// CSelection dialog

class CSelection : public CDialog
{
// Construction
public:
	CSelection(CWnd* pParent = NULL);   // standard constructor
    ~CSelection();
// Dialog Data
	//{{AFX_DATA(CSelection)
	enum { IDD = IDD_SELECTIONS };
	CStatic	m_Header;
	CTreeCtrlCheckBox	m_tvwSelection;
	int		m_ShowPlayers;
	BOOL	m_chkPlayer;
	BOOL	m_chkSalary;
	BOOL	m_chkTeam;
	BOOL	m_chkYear;
	BOOL	m_chkGroupTeam;
	long	m_txtLow;
	long	m_txtHigh;
	long	m_txtSalaryHigher;
	long	m_txtSalaryLower;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSelection)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
    CDaoDatabase m_Database;
    CString RootTag[numLeagues];
    CString SecondLevelTag[25];
    int iLeagueNum;
    int iLeagueYearNum;
    HTREEITEM hRoot;
	CImageList m_imageState;
    CString m_Players[numLeagues + 1];
    CString m_Teams[numLeagues + 1]; 
    CString m_LeagueName[numLeagues + 1];
    
protected:
    void InitTreeView();
    void InitVisuals();
    void AddNodes(CDaoRecordset* RS, char* League, char* Year, char* TableName, BOOL NewLeague);
    long NeedTable(HTREEITEM nodh);
    long NeedTeam(HTREEITEM nodh);
	void GenerateSQLStatement();
    void SaveSettings();
    // Generated message map functions
	//{{AFX_MSG(CSelection)
	afx_msg void OnItemexpandingTvwselect(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	afx_msg void OnShowReport();
	virtual BOOL OnInitDialog();
	afx_msg void OnAbout();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SELECTION_H__12007A5C_4C47_11D3_A685_00902741EC66__INCLUDED_)
