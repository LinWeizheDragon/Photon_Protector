// kcldscantest.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
//包含炫彩界面库文件


#include <iostream>
#include <cctype>
#include <conio.h>
#include <locale>
#include <string>
#include "testcustomscan.h"

using namespace std;


//-------------------------------------------------------------------------

HRESULT TestClient();
HELE hStatic;
HWINDOW hWindow;
//-------------------------------------------------------------------------
BOOL CALLBACK My_EventBtnClick(HELE hEle,HELE hEventEle)
{
	XStatic_SetText(hStatic,L"加载服务中，如果加载失败将无法开始扫描。");
	//SetText(hStatic);
    system("koemsec1.exe -Service");
	system("koemsec1.exe -Start");
    ::TestCustomScan(TRUE);
    return false;
}

int _tmain(int argc, _TCHAR* argv[])
{
	XInitXCGUI(); //初始化
    
    hWindow=XWnd_CreateWindow(0,0,700,500,L"光子防御网杀毒模块");//创建窗口
    if(hWindow)
    {
        HELE hButton=XBtn_Create(330,115,80,25,L"开始扫描",hWindow);//创建按钮
        XEle_RegisterEvent(hButton,XE_BNCLICK,My_EventBtnClick);//注册按钮点击事件
		XWnd_EnableMaxButton(hWindow,FALSE,1);
		
		//创建进度条
		HELE hProgBar1=XProgBar_Create(61,70,560,20,true,hWindow);
		XProgBar_SetPos(hProgBar1,50); //设置进度
		XProgBar_EnablePercent(hProgBar1,true);
		XWnd_ShowWindow(hWindow,SW_SHOW);//显示窗口

		hStatic=XStatic_Create(10,10,668,57,L"欢迎使用光子防御网扫描程序！\n本程序基于金山云安全开放平台的API制作，略显粗糙，需联网使用。一切解释权归金山公司所有。",hWindow);
        XEle_SetBkTransparent(hStatic,true); //设置背景透明
        XRunXCGUI(); //运行
    }
	/*TCHAR szDir[_MAX_PATH];
	ios::sync_with_stdio(true);
	::setlocale(LC_ALL, "CHS");
	cout<<"加载服务中，如果出现错误提示，服务启动失败，扫描可能无法继续。\n";
	system("koemsec1.exe -Service");
	system("koemsec1.exe -Start");
		cout<<"\n\n\n欢迎使用光子防御网扫描程序！\n本程序基于金山云安全开放平台的API制作，略显粗糙，需联网使用。"
			<<"\n一切解释权归金山公司所有。";

	HRESULT	hRet = TestClient();
	cout<<"终止服务中，请稍候。\n";
	system("koemsec1.exe -Stop");
	system("koemsec1.exe -UnRegServer");
	////::printf("按下任意键退出扫描\n");
	::_getch();*/

	return 0;
}

HRESULT TestClient()
{
	
	HRESULT	hRet	= E_FAIL;
	char	cChoice	= 0;
	do 
	{

		cout << "\n----------------------------------\n"
			<< "[1]. 全盘扫描\n"
			//<< "[2]. 自定义扫描\n"
			<< "[2]. 退出\n"
			<< "请输入选择:\n";

		cChoice = ::_getch();
		switch (cChoice)
		{
		case '1':
			//cout << "Test full scan...\n";
			hRet = ::TestCustomScan(TRUE);
			break;
		/*case '2':
			//cout << "Test custom scan...\n";
			hRet = ::TestCustomScan(FALSE);
			break;*/
		case '2':
			return S_OK;
		//default:
			//cout << "\nIncorrect input...\n";
		}
		
	} while (cChoice < '1' || cChoice > '4');

	return hRet;
}
