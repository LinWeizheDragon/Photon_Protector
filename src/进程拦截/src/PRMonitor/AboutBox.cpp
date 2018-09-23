// AboutBox.cpp : implementation file
//

#include "stdafx.h"

#include "AboutBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutBox dialog


CAboutBox::CAboutBox(CWnd* pParent /*=NULL*/)
	: CDialog(CAboutBox::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAboutBox)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CAboutBox::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutBox)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAboutBox, CDialog)
	//{{AFX_MSG_MAP(CAboutBox)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAboutBox message handlers

void CAboutBox::OnOK() 
{
	// TODO: Add extra validation here
	CResult=1;
	if (CResult==1)
		AfxMessageBox("A");
	CDialog::OnOK();
}

void CAboutBox::OnCancel() 
{
	// TODO: Add extra cleanup here
		CResult=2;
	if (CResult==1)
		AfxMessageBox("A");
	
	CDialog::OnCancel();
}
