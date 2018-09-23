#if !defined(AFX_MYABOUTBOX_H__F670C9B6_FB83_49F2_8CB1_8D1098D863D4__INCLUDED_)
#define AFX_MYABOUTBOX_H__F670C9B6_FB83_49F2_8CB1_8D1098D863D4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// MyAboutBox.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CMyAboutBox dialog

class CMyAboutBox : public CDialog
{
// Construction
public:
	CMyAboutBox(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CMyAboutBox)
//	enum { IDD = IDD_ABOUTBOX };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMyAboutBox)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CMyAboutBox)
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MYABOUTBOX_H__F670C9B6_FB83_49F2_8CB1_8D1098D863D4__INCLUDED_)
