#if !defined(AFX_TEXTOUT_H__E5DCD745_9FEB_4C39_BBE6_534301AE97A3__INCLUDED_)
#define AFX_TEXTOUT_H__E5DCD745_9FEB_4C39_BBE6_534301AE97A3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// TextOut.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CTextOut dialog

class CTextOut : public CDialog
{
// Construction
public:
	CTextOut(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CTextOut)
	enum { IDD = IDD_DIALOG2 };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTextOut)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CTextOut)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TEXTOUT_H__E5DCD745_9FEB_4C39_BBE6_534301AE97A3__INCLUDED_)
