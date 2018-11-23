#if !defined(AFX_ABOUTPROSALARIES_H__727B4A10_76EE_11D3_A701_00902741EC66__INCLUDED_)
#define AFX_ABOUTPROSALARIES_H__727B4A10_76EE_11D3_A701_00902741EC66__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AboutProSalaries.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAboutProSalaries dialog

class CAboutProSalaries : public CDialog
{
// Construction
public:
	CAboutProSalaries(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAboutProSalaries)
	enum { IDD = IDD_ABOUTBOX };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutProSalaries)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAboutProSalaries)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ABOUTPROSALARIES_H__727B4A10_76EE_11D3_A701_00902741EC66__INCLUDED_)
