; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CCwordDlg
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "cword.h"

ClassCount=3
Class1=CCwordApp
Class2=CCwordDlg
Class3=CAboutDlg

ResourceCount=3
Resource1=IDD_ABOUTBOX
Resource2=IDR_MAINFRAME
Resource3=IDD_CWORD_DIALOG

[CLS:CCwordApp]
Type=0
HeaderFile=cword.h
ImplementationFile=cword.cpp
Filter=N

[CLS:CCwordDlg]
Type=0
HeaderFile=cwordDlg.h
ImplementationFile=cwordDlg.cpp
Filter=D
BaseClass=CDialog
VirtualFilter=dWC
LastObject=CCwordDlg

[CLS:CAboutDlg]
Type=0
HeaderFile=cwordDlg.h
ImplementationFile=cwordDlg.cpp
Filter=D

[DLG:IDD_ABOUTBOX]
Type=1
Class=CAboutDlg
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308480
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889

[DLG:IDD_CWORD_DIALOG]
Type=1
Class=CCwordDlg
ControlCount=6
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_STATIC,static,1342312449
Control4=IDC_NEW_WORD,button,1342242816
Control5=IDC_PICTURE,button,1342242816
Control6=IDC_BIAOGE,button,1342242816

