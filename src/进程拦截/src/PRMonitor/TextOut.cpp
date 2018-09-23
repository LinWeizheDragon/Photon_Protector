// TextOut.cpp : implementation file
//

#include "stdafx.h"
#include "prmdlg.h"
#include "TextOut.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTextOut dialog


CTextOut::CTextOut(CWnd* pParent /*=NULL*/)
	: CDialog(CTextOut::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTextOut)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CTextOut::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTextOut)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CTextOut, CDialog)
	//{{AFX_MSG_MAP(CTextOut)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTextOut message handlers
