// AboutProSalaries.cpp : implementation file
//

#include "stdafx.h"
#include "ProSalaries.h"
#include "AboutProSalaries.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutProSalaries dialog


CAboutProSalaries::CAboutProSalaries(CWnd* pParent /*=NULL*/)
	: CDialog(CAboutProSalaries::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAboutProSalaries)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CAboutProSalaries::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutProSalaries)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAboutProSalaries, CDialog)
	//{{AFX_MSG_MAP(CAboutProSalaries)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

