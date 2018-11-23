//{{AFX_INCLUDES()
#include "crystalreportviewer3.h"
//}}AFX_INCLUDES
#if !defined(REPORTVIEW_H)
#define REPORTVIEW_H

#include <afxdao.h>
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#if _MSC_VER >= 1100        // VC 5.0/6.0 specific -- support for 
    #include <comdef.h>     // #import and _com_error exception class
#endif

#import "craxdrt.tlb" no_namespace

// ReportView.h : header file
//
/////////////////////////////////////////////////////////////////////////////
// ReportView dialog

class ReportView : public CDialog
{
// Construction
public:
	ReportView(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(ReportView)
	enum { IDD = IDD_REPORT_VIEW };
	CCrystalReportViewer3	m_CRViewer1;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(ReportView)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
    IApplicationPtr m_Application;
    IReportPtr m_Report;
    ISectionsPtr m_Sections;

    // Variable Used in the Page Header Procedures
    ITextObjectPtr pPlayerSalariesTitle;

    // Variables used in the Grouping Procedures
    IFieldObjectPtr pTeamGroupObj;
    IFieldObjectPtr pTeamSummary;
    ITextObjectPtr pSalaryTitle;
    long vOffset;
    long m_maxLogoHeight;
    long LocX;
    short groupNum;
    int sectionNum;

    // Variables used in generating the Details
    IFormulaFieldDefinitionPtr PlayerFormula, SalaryFormula, TeamFormula, YearFormula;

    // A dummy variable used to represent NULL variant values throughout the report
    VARIANT dummy;

protected:
    // Report Initialization
    void ResetVariables();
    void InitReport();
    void DoReportHeader();
    void DoReportSizing();

    // Page Header
    void DoPageHeader();
    void AddPageFieldObjects();
    void FormatPageObjects();

    // Grouping
    void DoGroupHeader();
    void AddGroups();
    void AddGroupTitles();
    void AddGroupFieldObjects();
    void FormatGroupObjects();

    // Details
    void DoDetails();
    void AddDetailFormulas();
    void AddDetailFieldObjects();
    void FormatFieldObjects();
    void AddSortObject();

    // The rest of the Report Generating code
    void DoFooters();
    void AddDataSource();

    // Helper Function
    ISectionPtr GetReportSection(int);
	// Generated message map functions
	//{{AFX_MSG(ReportView)
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	virtual BOOL OnInitDialog();
	afx_msg void OnZoomLevelChangedCrviewer1(short ZoomLevel);
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_REPORTVIEW_H__12007A59_4C47_11D3_A685_00902741EC66__INCLUDED_)
