// ***************************************************************
//  PRMonitor   version:  1.0   ? date: 11/17/2007
//  -------------------------------------------------------------
//  Author:X-STAR
//  E-MAIL:qqshow@live.com
//  BLOG  :http://hi.baidu.com/heartdbg
//  -------------------------------------------------------------
//  Copyright (C) 2007 - All Rights Reserved
// ***************************************************************
// 
// ***************************************************************
#include <iostream.h>
//#include <Afx.h>
#include <windows.h>
#include <stdio.h>
#include <string.h>
#include "resource.h"
HANDLE hMapping;   //创建内存映像对象
LPSTR lpData;   
BOOL retz;
HWND hExe=0;
BOOL asking;
BOOL askresult;
HWND hDialog;
HANDLE device;
char outputbuff[256]; 
char * strings[256]; 
DWORD stringcount;
DWORD controlbuff[64];DWORD dw;
NOTIFYICONDATA nid;
HWND hwnd;
BOOL bdrv;
BOOL CALLBACK DialogProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam);
void CALLBACK TimerProc(HWND hWnd,UINT nMsg,UINT nTimerid,DWORD dwTime);
void steup();
void CloseSrv();
void Begin();
void thread();
#define WM_PRM_NOTIFY WM_USER+1
#define WM_MYMESSAGE WM_USER+100
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	MSG msg;
	HACCEL hAccel;

	hAccel = LoadAccelerators (hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));
	hDialog = CreateDialog(hInstance, MAKEINTRESOURCE(IDD_DIALOG1), NULL, DialogProc);
	hwnd = hDialog;
	nid.cbSize = sizeof(NOTIFYICONDATA);
	nid.hWnd	= hDialog;
	nid.hIcon  = LoadIcon(hInstance,MAKEINTRESOURCE(ID_ACCEL40001));
//	sprintf(nid.szTip,"PRMonitor BY:X-STAR(heartdbg)");
	nid.uCallbackMessage = WM_PRM_NOTIFY;
	nid.uFlags = NIF_ICON|NIF_TIP|NIF_MESSAGE;
	nid.uID = 100001;

//	Shell_NotifyIcon(NIM_ADD, &nid  );  
    
    //建立内存映射

	hMapping=CreateFileMapping((HANDLE)0xFFFFFFFF,NULL,PAGE_READWRITE,0,0x100,"PPProcessRTAChat");   
	if(hMapping==NULL)   
	{   
			::MessageBox(NULL, "创建内存文件映像失败！" , "龙盾 实时防护",   MB_OK);
		    CloseSrv();
			exit(0);//加载失败后退出
	}
	//将文件的视图映射到一个进程的地址空间上，返回LPVOID类型的内存指针
	lpData=(LPSTR)MapViewOfFile(hMapping,FILE_MAP_ALL_ACCESS,0,0,0);   
	if(lpData==NULL)   
	{   
			::MessageBox(NULL, "映射文件视图失败！" , "龙盾 实时防护",   MB_OK);
		    CloseSrv();
			exit(0);//加载失败后退出
	}

	RECT   rect;   
	int   ScreenWidth= GetSystemMetrics(SM_CXSCREEN);   
	int   ScreenHeight=  GetSystemMetrics(SM_CYSCREEN); 
	GetWindowRect(hDialog,&rect);
	int width = rect.right-rect.left;
	int height = rect.bottom-rect.top;
	MoveWindow(hDialog,ScreenWidth/2-width/2,ScreenHeight/2-height/2,width,height,TRUE);
    SetTimer(hDialog,1,1000,TimerProc);
	ShowWindow(hDialog, nCmdShow);
	ShowWindow(hDialog,SW_HIDE);
	UpdateWindow(hDialog);
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(hDialog, hAccel, &msg))
		{
			if(!IsDialogMessage(hDialog, &msg)) 
			{
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}			
		}
	}
	
	return msg.wParam;
}
BOOL CALLBACK DialogProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
			char* x="C:\\Documents and Settings\\Administrator\\桌面\\新建文件夹 (3)\\VCtest.exe|C:\\Windows\\System32\\Explorer.exe|PRO";
			char strPathName[1024];
    switch (message)
    {
    case WM_SHOWWINDOW:
		if (bdrv)
		{
			SetWindowText(GetDlgItem(hwnd,IDC_SA),"内核模块已成功加载");
		}
		else
		{
			SetWindowText(GetDlgItem(hwnd,IDC_SA),"内核模块加载失败");
			::MessageBox(NULL, "内核模块加载失败，请尝试重新开启，如果反复出现此提示，请重新安装本软件！" , "龙盾 实时防护",   MB_OK);
			CloseSrv();
			exit(0);//加载失败后退出
		}
		
        DeviceIoControl(device,
								1000,
								controlbuff,
								256,
								controlbuff,
								256,
								&dw,
								0);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_PMON),FALSE);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_PMOFF),TRUE);
		DeviceIoControl(device,
								 1004,
								 controlbuff,
								 256,
								 controlbuff,
								 256,
								 &dw,
								 0);
				 EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_MMON),FALSE);
				 EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_MMOFF),TRUE);
        
		return TRUE;
	case WM_INITDIALOG:
		Begin();
        return (TRUE);
	case WM_PRM_NOTIFY:
		if (lParam == WM_LBUTTONDOWN )
		{
			
			ShowWindow(hwndDlg,SW_RESTORE);
			SetForegroundWindow(hwndDlg);
		}
		else if (lParam == WM_RBUTTONDOWN)
		{

		}
		return TRUE ;
	case WM_SIZE :
		if (wParam == SIZE_MINIMIZED)
		{
			ShowWindow(hwndDlg,SW_HIDE);
		}
		else
		{
			return TRUE;
		}
		return TRUE;
    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_BTN_PMON:		
				DeviceIoControl(device,
								1000,
								controlbuff,
								256,
								controlbuff,
								256,
								&dw,
								0);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_PMON),FALSE);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_PMOFF),TRUE);
			return TRUE;
        case IDC_BTN_PMOFF:
				DeviceIoControl(device,
								1001,
								NULL,
								NULL,
								NULL,
								NULL,
								&dw,
								0);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_PMON),TRUE);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_PMOFF),FALSE);
			return TRUE;
		case IDC_BTN_RMON:
				DeviceIoControl(device,
								1002,
								controlbuff,
								256,
								controlbuff,
								256,
								&dw,
								0);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_RMON),FALSE);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_RMOFF),TRUE);
			return TRUE;
		case IDC_BTN_RMOFF:
				DeviceIoControl(device,
								1003,
								NULL,
								NULL,
								NULL,
								NULL,
								&dw,
								0);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_RMON),TRUE);
				EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_RMOFF),FALSE);
			return TRUE;
		case IDC_BTN_MMON:
				 DeviceIoControl(device,
								 1004,
								 controlbuff,
								 256,
								 controlbuff,
								 256,
								 &dw,
								 0);
				 EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_MMON),FALSE);
				 EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_MMOFF),TRUE);
				 return TRUE;
		case IDC_BTN_MMOFF:
				 DeviceIoControl(device,
								 1005,
								 NULL,
								 NULL,
								 NULL,
								 NULL,
								 &dw,
								 0);
				 EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_MMON),TRUE);
				 EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_MMOFF),FALSE);
				 return TRUE;
		case IDC_BUTTON1:
			

			// ::MessageBox(NULL, p , "Caption ",   MB_OK);
			//::WritePrivateProfileString ("RTA","Message",TEXT(p),".\\chat.ini");
		    sprintf(lpData,x);   //给这段映像内存写数据
			//UnmapViewOfFile(lpData);   //释放映像内存	
			::MessageBox(NULL, "开始" , "Caption ",   MB_OK);
			hExe = ::FindWindow(NULL, "MYRECEIVER");
			::SendMessage(hExe, WM_MYMESSAGE, 1,1);
				//::MessageBox(NULL, "成功发送" , "Caption ",   MB_OK);
			// sprintf(strPathName,lpData);	
			//UnmapViewOfFile(lpData);//释放映像内存
			strcpy(strPathName,lpData);
			::MessageBox(NULL, strPathName , "Caption ",   MB_OK);
			if (strcmp(strPathName,"Disallow")==0)
				::MessageBox(NULL, "1成功" , "Caption ",   MB_OK);
			else
				::MessageBox(NULL, "2失败" , "Caption ",   MB_OK);

       //     hExe = ::FindWindow(NULL, "MYRECEIVER");	
		//	::SendMessage(hExe, WM_MYMESSAGE, 1,1);
	return TRUE;
		case IDC_STATIC_BLOG:
				 ShellExecute(hwndDlg,
							  "open",
							  "http://hi.baidu.com/heartdbg",
							  NULL,							  
							  NULL,
							  SW_SHOWNORMAL);
				 return TRUE;
        default:
            break;
        }
        return (FALSE);
		
    case WM_CLOSE:
		CloseSrv();
		DestroyWindow(hwndDlg);
		PostQuitMessage(0);
		break;
    case WM_DESTROY:
		Shell_NotifyIcon(NIM_DELETE,&nid);
		CloseSrv();
        EndDialog (hwndDlg, 0);
        return (TRUE);
    }
    return (FALSE);
}

void steup()
{	
	int i;
	char namebuff[256];
	SC_HANDLE sch,scm;

	GetModuleFileName(NULL,namebuff,256);
	i = strlen(namebuff);
	while (namebuff[i] != '\\')
	{
		i--;
	}
	i++;

	strcpy(&namebuff[i],"PRMonitor.sys");

	sch = OpenSCManager(NULL,NULL,SC_MANAGER_ALL_ACCESS);
	scm = CreateService(sch,
						"PRMonitor",
						"PRMonitor",
						SERVICE_START|SERVICE_STOP,
						SERVICE_KERNEL_DRIVER,
						SERVICE_DEMAND_START,
						SERVICE_ERROR_NORMAL,
						namebuff,
						0,
						0,
						0,
						0,
						0);

	StartService(scm,NULL,NULL);

	CloseServiceHandle(scm);
	
}

void CloseSrv()
{
	SC_HANDLE sch ;
	SERVICE_STATUS ss;
	sch = OpenSCManager(NULL,0,0);
	SC_HANDLE scm ;
	scm = OpenService(sch,"PRMonitor",SERVICE_ALL_ACCESS);
	ControlService(scm,SERVICE_CONTROL_STOP,&ss);
	DeleteService(scm);
}

void Begin()
{
	steup();

	
	Sleep(100);
	//create processing thread
	CreateThread(0,0,(LPTHREAD_START_ROUTINE)thread,0,0,&dw);
	
	//open device
	device=CreateFile("\\\\.\\PRMONITOR",
						GENERIC_READ|GENERIC_WRITE,
						0,
						0,
						OPEN_EXISTING,
						FILE_ATTRIBUTE_SYSTEM,
						0);
	
	if (device == INVALID_HANDLE_VALUE)
	{
		bdrv = FALSE;
		
	}
	else
	{
		bdrv = TRUE;
	}

	DWORD * addr=(DWORD *)(1+(DWORD)GetProcAddress(GetModuleHandle("ntdll.dll"),"NtCreateProcess"));
	ZeroMemory(outputbuff,256);
	controlbuff[0]=addr[0];
	controlbuff[1]=(DWORD)&outputbuff[0];
	

}

void thread()
{
	DWORD a,x; char msgbuff[512];
	char *pdest;
	int  result;
	char text[512];
	while(1)
	{
		memmove(&a,&outputbuff[0],4);
		
		
		if(!a){Sleep(10);continue;}
		
		
		
		char*name=(char*)&outputbuff[8];
		for(x=0;x<stringcount;x++)
		{
			if(!stricmp(name,strings[x])){a=1;goto skip;}
		}
		
		
		char* p;
		pdest = strstr(name,"##");
		if (pdest != NULL)
		{
			
            p = strtok(&outputbuff[8], "##");//切分字符串，类型为：“Notepad.exe”
           p = strtok(NULL, "##");//再次切分，类型为：“C:\Windows\System32\Notepad.exe”
			//p="PRO";
			//这个据说是用来查看信息的。
			//::MessageBox(NULL, p , "Caption ",   MB_OK);
			result = pdest-name;
			strcpy(msgbuff, "|");
			strncat(msgbuff,&outputbuff[8],result);
			strcat(msgbuff,"|PRO");
			//strcat(msgbuff,&outputbuff[result+10]);
			
		}
		else if((pdest=strstr(name,"$$")) != NULL)
		{	
			
			p = strtok(&outputbuff[8], "$$");//切分字符串，类型为：“Notepad.exe”
            p = strtok(NULL, "$$");//再次切分，类型为：“C:\Windows\System32\Notepad.exe”
		//	::MessageBox(NULL, p , "Caption ",   MB_OK);
			p="REG";
			result = pdest-name;
			strcpy(msgbuff, "|");
			strncat(msgbuff,&outputbuff[8],result);
			strcat(msgbuff,"设置注册表");
			strcat(msgbuff,&outputbuff[result+10]);
		}
		else
		{
			
			p = strtok(&outputbuff[8], "&&");//切分字符串，类型为：“Notepad.exe”
            p = strtok(NULL, "&&");//再次切分，类型为：“C:\Windows\System32\Notepad.exe”
			//p="DRV";
			//这个据说是用来查看信息的。
		//	::MessageBox(NULL, p , "Caption ",   MB_OK);
			pdest = strstr(name,"&&");
			result = pdest-name;
			strcpy(msgbuff,"|");
			strncat(msgbuff,&outputbuff[8],result);
			strcat(msgbuff,"|DRV");
			//strcat(msgbuff,&outputbuff[result]+10);
		}
		//---------------------------------------------------
/*char* strPath;
char* strInit;
//WritePrivateProfileString("RTA","Message",strPath,".\\Chat.ini");
GetPrivateProfileString("RTA", "Message", NULL, strPath, sizeof(strPath),".\\Chat.ini"); 
::MessageBox(NULL,strPath,"..",MB_OK);

*/

/*			CString Path,Create,Init,Type;
	CString strKeyName="Message";
	CString strClassName="RTA";
	CString Show;*/
 /*   BOOL Rec;
	Rec=FALSE;
	Type="Create";
	Path=Type+"|C:\\Windows\\Regedit.exe";
    Init=Path;
	WritePrivateProfileString (strClassName,strKeyName,Path,".\\Chat.ini");//写入

    Init=Path;
	//ShowWindow(SW_HIDE);
    while (Rec!=TRUE)//直到Rec不为FALSE
	{
		GetPrivateProfileString(strClassName,strKeyName,NULL,
		Path.GetBuffer(50),50,".\\Chat.ini");//获取配置文件
		//Path.ReleaseBuffer;//释放
			//AfxMessageBox(Path);
	    if (Path!=Init)//如果变化了
		{
		//	Path=Path.Right(1);
			//Show.Format ("获取的消息内容是%d",Path);
			//AfxMessageBox(Show);
		   if (Path=="1")
		   {
			   AfxMessageBox("同意放行");
			   a=1;
			   strings[stringcount]=_strdup(name);
			stringcount++;
		   }
		   else
		   {
			   AfxMessageBox("不同意放行");
			   a=0
		   }
		   Rec=TRUE;//变为TRUE
		   //ShowWindow(SW_RESTORE);
		}
		Sleep(1000);
	}*/
	//---------------------------------------------------------------
    /*    do while(asking=TRUE)
		{
	 COleDateTime  start_time = COleDateTime::GetCurrentTime();  
     COleDateTimeSpan  end_time= COleDateTime::GetCurrentTime()-start_time;   
	 while(end_time.GetTotalSeconds()< 10) //   
	 {              
		 MSG   msg;        
		 GetMessage(&msg,NULL,0,0);    
		 TranslateMessage(&msg);  
		 DispatchMessage(&msg);          
         end_time = COleDateTime::GetCurrentTime()-start_time;   
	 }
		}
		if (asking=TRUE)
			a=0;
		else
		{
		    asking=TRUE;
			askresult=NULL;
			do while (askresult!=NULL)
			{
					 COleDateTime  start_time = COleDateTime::GetCurrentTime();  
     COleDateTimeSpan  end_time= COleDateTime::GetCurrentTime()-start_time;   
	 while(end_time.GetTotalSeconds()< 10) //   
	 {              
		 MSG   msg;        
		 GetMessage(&msg,NULL,0,0);    
		 TranslateMessage(&msg);  
		 DispatchMessage(&msg);          
         end_time = COleDateTime::GetCurrentTime()-start_time;   
	 }
			}
		if (askresult=TRUE)
		{*/
	//	HWND hWnd = ::GetDlgItem(GetSafeHwnd(), IDC_STATIC);

   strcat(p,msgbuff);
   char strPathName[1024];
  // ::MessageBox(NULL, p , "Caption ",   MB_OK);
   //::WritePrivateProfileString ("RTA","Message",TEXT(p),".\\chat.ini");
            sprintf(lpData,p);   //给这段映像内存写数据
			//UnmapViewOfFile(lpData);   //释放映像内存	
			//::MessageBox(NULL, "开始" , "Caption ",   MB_OK);
			hExe = ::FindWindow(NULL, "MYRECEIVER");
			::SendMessage(hExe, WM_MYMESSAGE, 1,1);
			//::MessageBox(NULL, "成功发送" , "Caption ",   MB_OK);
			// sprintf(strPathName,lpData);	
			//UnmapViewOfFile(lpData);//释放映像内存
			strcpy(strPathName,lpData);
		//	::MessageBox(NULL, strPathName , "Caption ",   MB_OK);
		if (strcmp(strPathName,"Disallow")==0)
		a=0;
		else 
		{
			a=1;
			strings[stringcount]=_strdup(name);
			stringcount++;
		}
		
		
skip:memmove(&outputbuff[4],&a,4);
	 
	 
	 a=0;
	 memmove(&outputbuff[0],&a,4);
	}
	
}

void CALLBACK TimerProc(HWND hWnd,UINT nMsg,UINT nTimerid,DWORD dwTime)
{
    if (strcmp(lpData,"ProcessRTA.Close")==0)
	{
		CloseSrv();
			exit(0);//加载失败后退出
	}

}

