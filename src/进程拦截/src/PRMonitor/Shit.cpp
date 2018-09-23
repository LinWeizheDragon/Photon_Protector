// Shit.cpp : implementation file
//

#include "stdafx.h"
#include "prmdlg.h"
#include "Shit.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// Shit dialog


Shit::Shit(CWnd* pParent /*=NULL*/)
	: CDialog(Shit::IDD, pParent)
{
	//{{AFX_DATA_INIT(Shit)
	//}}AFX_DATA_INIT
}


void Shit::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(Shit)
	DDX_Control(pDX, IDC_BUTTON1, m_bitmapbutton);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(Shit, CDialog)
	//{{AFX_MSG_MAP(Shit)
	ON_WM_LBUTTONDOWN()
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	ON_BN_CLICKED(IDC_BUTTON2, OnButton2)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Shit message handlers

void Shit::OnLButtonDown(UINT nFlags, CPoint point) 
{
	// TODO: Add your message handler code here and/or call default
	
	AfxMessageBox("1");


	CDialog::OnLButtonDown(nFlags, point);
}

void Shit::OnButton1() 
{

}

void Shit::OnButton2() 
{
AfxMessageBox("1");	
}
