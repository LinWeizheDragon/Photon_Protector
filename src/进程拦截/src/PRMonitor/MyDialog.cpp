// MyDialog.cpp : implementation file
//

#include "stdafx.h"
#include "prmdlg.h"
#include "MyDialog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// MyDialog dialog


MyDialog::MyDialog(CWnd* pParent /*=NULL*/)
	: CDialog(MyDialog::IDD, pParent)
{
	//{{AFX_DATA_INIT(MyDialog)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void MyDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(MyDialog)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(MyDialog, CDialog)
	//{{AFX_MSG_MAP(MyDialog)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// MyDialog message handlers
