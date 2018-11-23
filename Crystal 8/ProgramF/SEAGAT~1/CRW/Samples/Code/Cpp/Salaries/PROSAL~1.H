// ProSalaries.h : main header file for the PROSALARIES application
//

#if !defined(AFX_PROSALARIES_H__97E788C4_7610_11D3_A6FE_00902741EC66__INCLUDED_)
#define AFX_PROSALARIES_H__97E788C4_7610_11D3_A6FE_00902741EC66__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CProSalariesApp:
// See ProSalaries.cpp for the implementation of this class
//

class CProSalariesApp : public CWinApp
{
public:
	CProSalariesApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CProSalariesApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CProSalariesApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROSALARIES_H__97E788C4_7610_11D3_A6FE_00902741EC66__INCLUDED_)
