// VCTestDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "VCTest.h"
#include "VCTestDlg.h"
#include "ProcProtectCtrl_i.c"
#include ".\vctestdlg.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif
	HANDLE hMapping;
	LPSTR lpData;
	CString MyStr;
		DWORD		dwResult = -1;
			IProcProtect*				m_pProcProtect;
// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持
	VOID   CALLBACK   TimerProc(HWND   hwnd,UINT   uMsg,UINT   idEvent,DWORD   dwTime);
// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CVCTestDlg 对话框



CVCTestDlg::CVCTestDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CVCTestDlg::IDD, pParent)
	, m_lPid(-1)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	
	HRESULT				hResult;

	::CoInitialize(NULL);
	hResult = ::CoCreateInstance(CLSID_ProcProtect, NULL, CLSCTX_INPROC_SERVER, IID_IProcProtect, (void**)&m_pProcProtect);
	if(!SUCCEEDED(hResult))
	{
		m_pProcProtect->Register(_T("YitProcProtectCtrlSample"));
		::AfxMessageBox(_T("ProcProtect component create failed!"));
	}


}

CVCTestDlg::~CVCTestDlg()
{
	if(m_pProcProtect)
	{
		m_pProcProtect->Release();
	}
}


void CVCTestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_PID, m_lPid);
}

BEGIN_MESSAGE_MAP(CVCTestDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BTN_DISPROTECT, OnBnClickedBtnDisprotect)
	ON_BN_CLICKED(IDC_BTN_PROTECT, OnBnClickedBtnProtect)
END_MESSAGE_MAP()


// CVCTestDlg 消息处理程序

BOOL CVCTestDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将\“关于...\”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	
	return TRUE;  // 除非设置了控件的焦点，否则返回 TRUE
}

void CVCTestDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CVCTestDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
		// TODO: 在此添加控件通知处理程序代码
	CString MyString;
	ShowWindow(SW_HIDE);
	hMapping=CreateFileMapping((HANDLE)0xFFFFFFFF,NULL,PAGE_READWRITE,0,0x100,"PPProtectSelf");
	if (hMapping==NULL)
	{
		AfxMessageBox("创建文件映像失败");
		return;
	}
	lpData=(LPSTR)MapViewOfFile(hMapping,FILE_MAP_ALL_ACCESS,0,0,0);
	if (lpData==NULL)
	{
		AfxMessageBox("映射文件视图失败！");
		return;
	}
	//AfxMessageBox("设定计时器");
	//SetTimer(1,1000,TimerProc);
Back:
	if (strcmp("Unload",lpData)==0)
	{
		ShowWindow(SW_SHOW);
		if(m_pProcProtect)
		{
			m_pProcProtect->Release();
		}
		exit(0);
	}

	if (strcmp("Wait",lpData)==0)
	{
		Sleep (1000);
		goto Back;
	}
	else
	{
		MyString.Format("%s",lpData);
		dwResult=_ttoi(MyString);
		m_pProcProtect->Protect(dwResult, TRUE, &dwResult);
		if(dwResult != -1)
		{
			::MessageBox(NULL,"成功保护进程！","光子防御网-进程保护",MB_ICONWARNING);
		}
		else
		{
			::MessageBox(NULL,"保护进程失败！","光子防御网-进程保护",MB_ICONWARNING);
		}
		if (!strcmp("Wait",lpData)==0)
		{
				Sleep (1500);
		}
		
		goto Back;
		
		
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标显示。
HCURSOR CVCTestDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CVCTestDlg::OnBnClickedBtnDisprotect()
{
	// TODO: 在此添加控件通知处理程序代码
	//UpdateData();
	
	m_pProcProtect->Protect(m_lPid, FALSE, &dwResult);
	if(dwResult != -1)
	{
		::AfxMessageBox(_T("成功取消保护!"));
	}
	else
	{
		::AfxMessageBox(_T("取消保护失败!"));
	}

}



VOID   CALLBACK   TimerProc(HWND   hwnd,UINT   uMsg,UINT   idEvent,DWORD   dwTime)   
{   

}  

void CVCTestDlg::OnBnClickedBtnProtect()
{
	// TODO: 在此添加控件通知处理程序代码
	CString MyString;
	ShowWindow(SW_HIDE);
	hMapping=CreateFileMapping((HANDLE)0xFFFFFFFF,NULL,PAGE_READWRITE,0,0x100,"PPProtectSelf");
	if (hMapping==NULL)
	{
		AfxMessageBox("创建文件映像失败");
		return;
	}
	lpData=(LPSTR)MapViewOfFile(hMapping,FILE_MAP_ALL_ACCESS,0,0,0);
	if (lpData==NULL)
	{
		AfxMessageBox("映射文件视图失败！");
		return;
	}
	//AfxMessageBox("设定计时器");
	//SetTimer(1,1000,TimerProc);
Back:
	if (strcmp("Unload",lpData)==0)
	{
		ShowWindow(SW_SHOW);
		exit(0);
		return;
	}

	if (strcmp("Wait",lpData)==0)
	{
		Sleep (1000);
		goto Back;
	}
	else
	{
		MyString.Format("%s",lpData);
		dwResult=_ttoi(MyString);
		m_pProcProtect->Protect(dwResult, TRUE, &dwResult);
		if(dwResult != -1)
		{
			::MessageBox(NULL,"成功保护进程！","光子防御网-进程保护",MB_ICONWARNING);
		}
		else
		{
			::MessageBox(NULL,"保护进程失败！","光子防御网-进程保护",MB_ICONWARNING);
		}
		if (!strcmp("Wait",lpData)==0)
		{
				Sleep (1500);
		}
		
		goto Back;
		
		
	}



	/*
	DWORD MyPro=-1;
	UpdateData();
	HANDLE myhProcess;
PROCESSENTRY32 mype;
mype.dwSize = sizeof(PROCESSENTRY32); 
BOOL mybRet;
//进行进程快照
myhProcess=CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS,0); //TH32CS_SNAPPROCESS快照所有进程
//开始进程查找
mybRet=Process32First(myhProcess,&mype);
//循环比较，得出ProcessID
while(mybRet)
{
if(strcmp("DragonShield.exe",mype.szExeFile)==0)
{
MyPro = mype.th32ProcessID;
m_pProcProtect->Protect(MyPro, TRUE, &dwResult);
goto Next;
}
else
mybRet=Process32Next(myhProcess,&mype);
}
Next:

//进行进程快照
myhProcess=CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS,0); //TH32CS_SNAPPROCESS快照所有进程
//开始进程查找
mybRet=Process32First(myhProcess,&mype);
//循环比较，得出ProcessID
while(mybRet)
{
if(strcmp("PhotonProtect.exe",mype.szExeFile)==0)
{
MyPro = mype.th32ProcessID;
m_pProcProtect->Protect(MyPro, TRUE, &dwResult);
goto Next1;
}
else
mybRet=Process32Next(myhProcess,&mype);
}
Next1:
//进行进程快照
myhProcess=CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS,0); //TH32CS_SNAPPROCESS快照所有进程
//开始进程查找
mybRet=Process32First(myhProcess,&mype);
//循环比较，得出ProcessID
while(mybRet)
{
if(strcmp("ProcessRTA.exe",mype.szExeFile)==0)
{
MyPro = mype.th32ProcessID;
m_pProcProtect->Protect(MyPro, TRUE, &dwResult);
goto Next2;
}
else
mybRet=Process32Next(myhProcess,&mype);
}
Next2:*/
return;
}
