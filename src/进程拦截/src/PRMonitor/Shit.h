#if !defined(AFX_SHIT_H__9A325785_AE88_4B8B_BCAC_463011ADAB43__INCLUDED_)
#define AFX_SHIT_H__9A325785_AE88_4B8B_BCAC_463011ADAB43__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Shit.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// Shit dialog

class Shit : public CDialog
{
// Construction
public:
	Shit(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(Shit)
	enum { IDD = IDD_DIALOG1 };
	CButton	m_bitmapbutton;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(Shit)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(Shit)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnButton1();
	afx_msg void OnButton2();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SHIT_H__9A325785_AE88_4B8B_BCAC_463011ADAB43__INCLUDED_)
