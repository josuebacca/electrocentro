#if !defined(AFX_TREECTRLCHECKBOX_H__12007A58_4C47_11D3_A685_00902741EC66__INCLUDED_)
#define AFX_TREECTRLCHECKBOX_H__12007A58_4C47_11D3_A685_00902741EC66__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// TreeCtrlCheckBox.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CTreeCtrlCheckBox window

class CTreeCtrlCheckBox : public CTreeCtrl
{
// Construction
public:
	CTreeCtrlCheckBox();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTreeCtrlCheckBox)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CTreeCtrlCheckBox();

	// Generated message map functions
protected:
    void CheckNode(HTREEITEM nodh, UINT checkState);
    //{{AFX_MSG(CTreeCtrlCheckBox)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TREECTRLCHECKBOX_H__12007A58_4C47_11D3_A685_00902741EC66__INCLUDED_)
