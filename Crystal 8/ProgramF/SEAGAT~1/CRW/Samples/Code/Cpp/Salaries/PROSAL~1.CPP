// ProSalaries.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "ProSalaries.h"
#include "Selection.h"

// Initialization of OLE
struct InitOle {
  InitOle()  { ::CoInitialize(NULL); }
  ~InitOle() { ::CoUninitialize();   }
} _init_InitOle_;       // Global Instance to force load/unload of OLE

HICON g_Icon;

/////////////////////////////////////////////////////////////////////////////
// CProSalariesApp

BEGIN_MESSAGE_MAP(CProSalariesApp, CWinApp)
	//{{AFX_MSG_MAP(CProSalariesApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProSalariesApp construction

CProSalariesApp::CProSalariesApp()
{
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CProSalariesApp object

CProSalariesApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CProSalariesApp initialization

BOOL CProSalariesApp::InitInstance()
{
	AfxEnableControlContainer();

	// Standard initialization
#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	g_Icon = LoadIcon(MAKEINTRESOURCE(IDR_MAINFRAME));
	CSelection dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// No Code necessary -- just exit the program
	}
	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
