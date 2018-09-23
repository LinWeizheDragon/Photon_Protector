; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=Shit
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "prmdlg.h"
LastPage=0

ClassCount=6

ResourceCount=4
Resource1=IDD_PROPPAGE_LARGE (English (U.S.))
Class1=Shit
Class2=Control
Class3=MyDialog
Class4=CMyDialog
Class5=CTextOut
Resource2=IDD_DIALOGBAR (English (U.S.))
Resource3=IDD_DIALOG1
Class6=CMyAboutBox
Resource4=IDR_ACCELERATOR1

[DLG:IDD_DIALOG1]
Type=1
Class=Shit
ControlCount=11
Control1=IDC_STATIC,button,1342177287
Control2=IDC_STATIC,button,1342177287
Control3=IDC_BTN_PMON,button,1342275584
Control4=IDC_BTN_PMOFF,button,1476493312
Control5=IDC_BTN_RMON,button,1342275584
Control6=IDC_BTN_RMOFF,button,1476493312
Control7=IDC_SA,static,1342312460
Control8=IDC_STATIC,button,1342177287
Control9=IDC_BTN_MMON,button,1342275584
Control10=IDC_BTN_MMOFF,button,1476493312
Control11=IDC_BUTTON1,button,1342242816

[ACL:IDR_ACCELERATOR1]
Type=1
Class=?
Command1=ID_ACCEL40001
CommandCount=1

[CLS:Shit]
Type=0
HeaderFile=Shit.h
ImplementationFile=Shit.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=ID_ACCEL40001

[CLS:MyDialog]
Type=0
HeaderFile=MyDialog.h
ImplementationFile=MyDialog.cpp
BaseClass=CDialog
Filter=D
LastObject=IDOK

[CLS:CMyDialog]
Type=0
HeaderFile=MyDialog1.h
ImplementationFile=MyDialog1.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CMyDialog

[CLS:CTextOut]
Type=0
HeaderFile=TextOut.h
ImplementationFile=TextOut.cpp
BaseClass=CDialog
Filter=D
LastObject=IDOK

[DLG:IDD_DIALOGBAR (English (U.S.))]
Type=1
Class=?
ControlCount=1
Control1=IDC_STATIC,static,1342308352

[CLS:CMyAboutBox]
Type=0
HeaderFile=MyAboutBox.h
ImplementationFile=MyAboutBox.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=IDOK

[DLG:IDD_PROPPAGE_LARGE (English (U.S.))]
Type=1
Class=?
ControlCount=1
Control1=IDC_STATIC,static,1342308352

