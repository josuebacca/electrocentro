; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CSelection
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "ProSalaries.h"

ClassCount=5
Class1=CProSalariesApp
Class2=ReportView

ResourceCount=4
Resource1=IDD_ABOUTBOX
Resource2=IDR_MAINFRAME
Class4=CSelection
Resource3=IDD_REPORT_VIEW
Class3=CTreeCtrlCheckBox
Class5=CAboutProSalaries
Resource4=IDD_SELECTIONS

[CLS:CProSalariesApp]
Type=0
HeaderFile=ProSalaries.h
ImplementationFile=ProSalaries.cpp
Filter=N
LastObject=CProSalariesApp

[DLG:IDD_ABOUTBOX]
Type=1
Class=CAboutProSalaries
ControlCount=9
Control1=IDC_STATIC,static,1342308480
Control2=IDC_STATIC,static,1342308352
Control3=IDOK,button,1342373889
Control4=IDC_STATIC,static,1342308352
Control5=IDC_STATIC,static,1342308352
Control6=IDC_STATIC,static,1342308352
Control7=IDC_STATIC,static,1342308352
Control8=IDC_STATIC,static,1342177283
Control9=IDC_STATIC,static,1342308352

[DLG:IDD_REPORT_VIEW]
Type=1
Class=CProSalariesDlg
ControlCount=2
Control1=IDOK,button,1342242817
Control2=IDC_CRVIEWER1,{C4847596-972C-11D0-9567-00A0C9273C2A},1342242816

[CLS:CTreeCtrlCheckBox]
Type=0
HeaderFile=TreeCtrlCheckBox.h
ImplementationFile=TreeCtrlCheckBox.cpp
BaseClass=CTreeCtrl
Filter=W
LastObject=CTreeCtrlCheckBox

[DLG:IDD_SELECTIONS]
Type=1
Class=CSelection
ControlCount=22
Control1=IDC_SHOW_REPORT,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_TVWSELECT,SysTreeView32,1350631811
Control4=IDC_GROUP_BY_TEAM_YEAR,button,1342242819
Control5=IDC_SHOW_ALL,button,1342308361
Control6=IDC_HIGHEST_PAID,button,1342177289
Control7=IDC_TXTHIGH,edit,1350639744
Control8=IDC_LOWEST_PAID,button,1342177289
Control9=IDC_TXTLOW,edit,1350639744
Control10=IDC_SALARY_HIGHER,button,1342177289
Control11=IDC_TXTSALARYHIGHER,edit,1350639744
Control12=IDC_SALARY_LOWER,button,1342177289
Control13=IDC_TXTSALARYLOWER,edit,1350639744
Control14=IDC_CHK_PLAYER,button,1342242819
Control15=IDC_CHK_TEAM,button,1342242819
Control16=IDC_CHK_SALARY,button,1342242819
Control17=IDC_CHK_YEAR,button,1342242819
Control18=IDC_FIELDSLABEL,static,1342308352
Control19=IDC_PLAYERREPORTLABEL,static,1342308352
Control20=IDC_STATIC,static,1342177294
Control21=IDC_STATIC,static,1342177294
Control22=IDC_ABOUT,button,1342242816

[CLS:CSelection]
Type=0
HeaderFile=Selection.h
ImplementationFile=Selection.cpp
BaseClass=CDialog
Filter=D
LastObject=CSelection
VirtualFilter=dWC

[CLS:ReportView]
Type=0
HeaderFile=ReportView.h
ImplementationFile=ReportView.cpp
BaseClass=CDialog
Filter=D
LastObject=IDC_CRVIEWER1
VirtualFilter=dWC

[CLS:CAboutProSalaries]
Type=0
HeaderFile=AboutProSalaries.h
ImplementationFile=AboutProSalaries.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CAboutProSalaries

