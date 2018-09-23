// MyAboutBox.cpp : implementation file
//

#include "stdafx.h"

#include "MyAboutBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMyAboutBox dialog


CMyAboutBox::CMyAboutBox(CWnd* pParent /*=NULL*/)
	: CDialog(CMyAboutBox::IDD, pParent)
{
	//{{AFX_DATA_INIT(CMyAboutBox)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CMyAboutBox::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CMyAboutBox)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CMyAboutBox, CDialog)
	//{{AFX_MSG_MAP(CMyAboutBox)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMyAboutBox message handlers

void CMyAboutBox::OnOK() 
{
	// TODO: Add extra validation here
	
	CDialog::OnOK();
}
