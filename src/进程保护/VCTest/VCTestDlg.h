// VCTestDlg.h : 头文件
//

#pragma once

#include "ProcProtectCtrl.h"
// CVCTestDlg 对话框
class CVCTestDlg : public CDialog
{
// 构造
public:
	
	CVCTestDlg(CWnd* pParent = NULL);	// 标准构造函数
	~CVCTestDlg();

// Attributes
protected:
	
	
public:
// Operations
protected:
public:

// 对话框数据
	enum { IDD = IDD_VCTEST_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;
	
	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	long m_lPid;
	
	afx_msg void OnBnClickedBtnDisprotect();
	afx_msg void OnBnClickedBtnProtect();
};
