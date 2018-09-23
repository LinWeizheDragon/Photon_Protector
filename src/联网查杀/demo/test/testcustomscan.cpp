//////////////////////////////////////////////////////////////////////
//
//  @ File		:	testcustomscan.cpp
//  @ Version	:	1.0
//  @ Author	:	EasyLogic <liangguangcai@kingsoft.com>
//  @ Datetime	:	[2010-6-2, 17:19:52]
//  @ Brief		:	测试自定义扫描功能
//
//////////////////////////////////////////////////////////////////////
#include "stdafx.h"

#include "testcustomscan.h"
#include "public_def.h"
#include "kscomdll.h"
#include "shlobj.h"

#include <string>
#include <vector>
HELE NewText;
HWINDOW MyWindow;
//-------------------------------------------------------------------------

#define QUERY_SESSION_STATUS_TIME_OUT		500

BOOL g_bFullScan				= FALSE;	// 为真表示全盘扫描，否则为自定义扫描
BOOL g_bCustomScanCompleted		= FALSE;	// 标识全盘扫描完成

BOOL g_bCustomScanNeedToStop	= FALSE;	// 须停止扫描
BOOL g_bNeedToPauseScanning		= FALSE;	// 暂停扫描
BOOL g_bNeedToResumeScanning	= FALSE;	// 继续扫描

//-------------------------------------------------------------------------

bool SetText(HELE TextControl)
{
	NewText=TextControl;
	return true;
}
bool trims( const std::wstring& str, std::vector <std::wstring>& vcResult, char c)
{
	size_t fst = str.find_first_not_of( c );
	size_t lst = str.find_last_not_of( c );

	if( fst != std::wstring::npos )
		vcResult.push_back(str.substr(fst, lst - fst + 1));

	return true;
}

bool SplitString( 
				 /*[in]*/  const std::wstring& str, 
				 /*[out]*/ std::vector <std::wstring>& vcResult,
				 /*[in]*/  char delim
				 )
{
	//ad,dd,dfd,sdf

	size_t nIter = 0;
	size_t nLast = 0;
	std::wstring v;

	while( true )
	{
		nIter = str.find(delim, nIter); 
		trims(str.substr(nLast, nIter - nLast), vcResult, delim);
		if( nIter == std::wstring::npos )
			break;

		nLast = ++nIter;
	}
	return true;
}


//-------------------------------------------------------------------------

DWORD WINAPI CustomScanThreadProc(LPVOID lpParameter)
{
	IKCldCustomScanClient* pCustomScan =
		reinterpret_cast<IKCldCustomScanClient*>(lpParameter);

	HRESULT hResult = E_FAIL;

	IKCldGetFullScanStatusInfo*		pGetFullInfo			= NULL;
	IKCldGetCustomScanStatusInfo*	pGetCustomInfo			= NULL;
	IKCldGetFileVirusThreatsInfo*	pGetThreatsInfo			= NULL;
	IKCldGetProcessResultInfo*		pGetProcessResultInfo	= NULL;

	S_KCLD_FULL_SCAN_STATUS*		pFullScanStatus;
	S_KCLD_CUSTOM_SCAN_STATUS*		pCustomScanStatus;

	DWORD							dwProcess			= 0;	// 进度
	BOOL							bCompleted			= FALSE;// 标识扫描完成
	DWORD							dwNewThreatCount	= 0;	// 发现新威胁数目
	DWORD							dwTotalThreatCount	= 0;	// 总威胁数
	DWORD							dwDisplayInterval	= 0;	// 显示间隔
	std::vector<std::wstring>		vecScanTargetFiles;			// 正在扫描目标
	std::vector<std::wstring>::iterator itervecScanTargetFiles;
	KCLD_FILEVIRUS_THREAT* pFileVirus = NULL;
	std::vector<DWORD> vecThreatIDs;
	wchar_t* DesText;
	/*
	* 查询扫描状态直到扫描过程完成
	*/
	MyWindow=XWnd_CreateWindow(0,0,800,500,L"光子防御网杀毒模块");//创建窗口
    if(MyWindow)
    {
		//创建进度条
		HELE hProgBar1=XProgBar_Create(61,70,560,20,true,MyWindow);
		XProgBar_SetPos(hProgBar1,25); //设置进度
		XProgBar_EnablePercent(hProgBar1,true);
		XWnd_ShowWindow(MyWindow,SW_SHOW);//显示窗口

		NewText=XStatic_Create(10,10,668,57,L"正在扫描：",MyWindow);
        XEle_SetBkTransparent(NewText,true); //设置背景透明
    }
	while (true)
	{
		if (g_bCustomScanNeedToStop)
		{
			//::printf("****用户停止扫描****\n");
			hResult = pCustomScan->StopScan();
			PRINT_FUNCTION_CALL_ERR_MSG("StopScan", hResult);

			g_bCustomScanCompleted = TRUE;

			::printf("按下任意键退出......\n");
			::_getch();
			goto Exit0;
		}

		if (g_bNeedToPauseScanning)
		{
			hResult = pCustomScan->PauseScan();
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"PauseScan",
				hResult
				);

			g_bNeedToPauseScanning = FALSE;

			while (!g_bNeedToResumeScanning)
			{
				::Sleep(200);
			}
		}

		if (g_bNeedToResumeScanning)
		{
			hResult = pCustomScan->ResumeScan();
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"ResumeScan",
				hResult
				);

			g_bNeedToResumeScanning = FALSE;
		}

		vecScanTargetFiles.clear();

		if (g_bFullScan)
		{
			hResult = pCustomScan->QueryFullScanStatus(
				&pGetFullInfo
				);
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"pCustomScan->QueryFullScanStatus",
				hResult
				);

			hResult = pGetFullInfo->GetInfo(&pFullScanStatus);
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"pGetFullInfo->GetInfo",
				hResult
				);

			bCompleted = (pFullScanStatus->scanStatus.emSessionStatus == em_KCLD_CUSTOM_ScanStatus_Complete);

			// 扫描进度（大于100则保持与上次相同）
			if (pFullScanStatus->scanStatus.progress.dwStreamProgress <= 100)
			{
				dwProcess = pFullScanStatus->scanStatus.progress.dwStreamProgress;
			}
			
			// 正在扫描目标
			if (NULL != pFullScanStatus->scanStatus.currentTarget.pszTargetName)
			{
				SplitString(
					std::wstring(pFullScanStatus->scanStatus.currentTarget.pszTargetName),
					vecScanTargetFiles,
					'\n'
					);
			}
		}
		else
		{
			hResult = pCustomScan->QueryCustomScanStatus(
				&pGetCustomInfo
				);
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"pCustomScan->QueryCustomScanStatus",
				hResult
				);

			hResult = pGetCustomInfo->GetInfo(&pCustomScanStatus);
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"pGetCustomInfo->GetInfo",
				hResult
				);

			bCompleted = (pCustomScanStatus->emSessionStatus == em_KCLD_CUSTOM_ScanStatus_Complete);

			// 扫描进度（大于100则保持与上次相同）
			if (pCustomScanStatus->progress.dwStreamProgress <= 100)
			{
				dwProcess = pCustomScanStatus->progress.dwStreamProgress;
			}

			// 正在扫描目标
			if (NULL != pCustomScanStatus->currentTarget.pszTargetName)
			{
				SplitString(
					std::wstring(pCustomScanStatus->currentTarget.pszTargetName),
					vecScanTargetFiles,
					'\n'
					);
			}
		}

		if (dwProcess > 100)
		{
			dwProcess = 0;
		}

		if (vecScanTargetFiles.size() > 0)
		{
			dwDisplayInterval = static_cast<DWORD>(QUERY_SESSION_STATUS_TIME_OUT / vecScanTargetFiles.size());
		}
		
		for (itervecScanTargetFiles = vecScanTargetFiles.begin();
			itervecScanTargetFiles != vecScanTargetFiles.end();
			++itervecScanTargetFiles)
		{
			 system("cls");
			 
			::wsprintf(
				DesText,
				L"扫描进度:\t%%%-3u \n正在扫描:\n\t%ls\n",
				dwProcess,
				itervecScanTargetFiles->c_str()
				);
			//XStatic_SetText(NewText,DesText);
			/*for (i=1;i+10;i<dwProcess)
			{
				::printf("▉");
			}*/
			::Sleep(dwDisplayInterval);
		}

		// 检查是否扫描完毕
		if (bCompleted)
		{
			system("cls");
			::printf("\n\n********扫描完成********\n\n");
			::Sleep(500);
			break;
		}

		if (NULL != pGetFullInfo)
		{
			pGetFullInfo->Release();
			pGetFullInfo = NULL;
		}

		if (NULL != pGetCustomInfo)
		{
			pGetCustomInfo->Release();
			pGetCustomInfo = NULL;
		}
	}

	/*
	* 处理
	*/
	KCLD_QUERY_THREAT querySetting;
	querySetting.dwStartIndex = 0;
	if (g_bFullScan)
	{
		querySetting.dwTotalCount = pFullScanStatus->scanStatus.dwFindThread;
	}
	else
	{
		querySetting.dwTotalCount = pCustomScanStatus->dwFindThread;
	}

	hResult = pCustomScan->QueryFileVirusThreats(
		&querySetting,
		&pGetThreatsInfo
		);
	PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
		"pCustomScan->QueryFileVirusThreats",
		hResult
		);

	hResult = pGetThreatsInfo->GetCount(&dwTotalThreatCount);
	PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
		"pGetThreatsInfo->GetCount",
		hResult
		);

	if (dwTotalThreatCount == 0)
	{
		::printf("没有发现威胁，您的电脑很安全！\n");
	}
	else
	{
		::printf("发现威胁\n");
		::printf("-----------------------------------------------------\n");
		::printf("序号\t|\t名称\t\t|\t文件路径\n");
		::printf("-----------------------------------------------------\n");

		for (int j = 0; j != dwTotalThreatCount; ++j)
		{
			hResult = pGetThreatsInfo->GetInfo(j, &pFileVirus);
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"pGetThreatsInfo->GetInfo",
				hResult
				);

			vecThreatIDs.push_back(pFileVirus->dwThreatID);
			::wprintf(
				L"%6d\t\t|\t%15d\t\t|\t%ls\n",
				pFileVirus->dwFoundThreatIndex,
				pFileVirus->dwThreatID,
				pFileVirus->pszFileFullPath
				);
		}

		::printf("输入Y修复威胁，输入N不修复威胁：[Y/N]\n");
		if (::toupper(::_getch()) != L'Y')
		{
			goto Exit0;
		}
		::printf("开始修复，请等待......\n");

		KCLD_PROCESS_SCAN_TARGET processTargets;
		processTargets.bClearAllThread = TRUE;
		processTargets.dwThreatIndexCount = dwTotalThreatCount;
		processTargets.pdwThreatIndexList = &(vecThreatIDs[0]);

		hResult = pCustomScan->ProcessScanResult(&processTargets);
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"pCustomScan->ProcessScanResult",
			hResult
			);

		::printf("修复完毕\n");

		hResult = pCustomScan->QueryScanThreatProcessResult(
			static_cast<unsigned int>(vecThreatIDs.size()),
			&(vecThreatIDs[0]),
			&pGetProcessResultInfo
			);
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"QueryScanThreatProcessResult",
			hResult
			);

		DWORD dwResultCount = 0;
		hResult = pGetProcessResultInfo->GetCount(&dwResultCount);
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"pGetProcessResultInfo->GetCount",
			hResult
			);

		if (0 == dwResultCount)
		{
			::printf("没有威胁正在运行\n");
		}
		else
		{
			::printf("-----------------------------------------------------\n");
			::printf("\t序号\t|\t进程标识\n");
			::printf("-----------------------------------------------------\n");
			for (int i = 0; i != dwResultCount; ++i)
			{
				S_KCLD_THREAT_PROCESS_RESULT*	pProcessResult	= NULL;
				hResult = pGetProcessResultInfo->GetInfo(i, &pProcessResult);
				PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
					"pGetProcessResultInfo->GetInfo",
					hResult
					);

				::printf("\t%d\t\t|\t%#010x\n",
					pProcessResult->dwThreadIndex,
					pProcessResult->eResult
					);
			}
			::printf("-----------------------------------------------------\n");
		}


		//::printf("是否需要重启电脑\n");

		BOOL bNeedToReboot = FALSE;
		hResult = pCustomScan->QueryNeedReboot(&bNeedToReboot);
		if (!bNeedToReboot)
		{
			::printf("不需要重启电脑就能完成本次处理\n");
		}
		else
		{
			::printf("需要重启电脑完成本次处理\n");
		}
	}

	g_bCustomScanCompleted = TRUE;
	::printf("按下任意键退出扫描！\n");

Exit0:
	if (NULL != pGetFullInfo)
	{
		pGetFullInfo->Release();
		pGetFullInfo = NULL;
	}

	if (NULL != pGetCustomInfo)
	{
		pGetCustomInfo->Release();
		pGetCustomInfo = NULL;
	}

	return hResult;
}

//-------------------------------------------------------------------------

HRESULT TestCustomScan(BOOL bFullScanning)
{
	

	g_bFullScan		= bFullScanning;

	HRESULT hResult = E_FAIL;
	IKCldCustomScanClient* pCustomScan = NULL;
	KSCOMDll scom_dll;

	hResult = scom_dll.Open(L"kcldscan.dll");
	PROCESS_COM_ERROR(hResult);

	scom_dll.GetClassObject(
		CLSID_KCldCustomScanClient,
		IID_IKCldCustomScanClient,
		(void**)&pCustomScan
		);
	if (NULL == pCustomScan)
		return 0;
	/*
	system("Cls");
	::printf("***************************提示***************************\n");
	::printf("**\t扫描开始后，您可以：\n");
	::printf("**\t  1.按空格键暂停\n");
	::printf("**\t  2.按下任意键停止扫描\n");
	::printf("**********************************************************\n");
	*/

	if (g_bFullScan)
	{	// 全盘
		/*::printf("现在按下任意键开始扫描\n");*/
		/*::_getch();*/
		
		hResult = pCustomScan->StartFullScan();
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"StartScan",
			hResult
			);
	}
	else
	{	// 自定义
		/*
		* 添加路径
		*/
		/*wchar_t pwszDesktop[MAX_PATH] = {0};
		if (::SHGetSpecialFolderPathW(
			NULL,
			pwszDesktop,
			CSIDL_DESKTOP,
			FALSE))
		{
			hResult = pCustomScan->AppendScanTargetPath(
				pwszDesktop,
				TRUE	// 扫子目录
				);
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"AppendScanTargetPath",
				hResult
				);
		}
		
		*/
		/*hResult = pCustomScan->AppendScanTargetPath(
			WStr,
			TRUE	// 不扫子目录
			);
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"AppendScanTargetPath",
			hResult
			);
		*/
		/*
		* 开始扫描
		*/
		//::printf("扫描开始\n");
		/*::wprintf(
			L"[1]. \"%ls\" (扫描子目录)\n[2]. \"%ls\" (不扫描子目录)\n",
			pwszDesktop,
			L"D:\\test"
			);*/
		//::printf("现在按下任意键开始！\n");
		//::_getch();

		hResult = pCustomScan->StartCustomScan();
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"StartCustomScan",
			hResult
			);
	}


	/*
	* 创建线程以查询扫描状态
	*/
	HANDLE hQueryThread = ::CreateThread(
		NULL,
		0,
		CustomScanThreadProc,
		pCustomScan,
		0,
		NULL
		);
	if (NULL == hQueryThread)
	{
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"CreateThread",
			(hResult = HRESULT_FROM_WIN32(::GetLastError()))
			);
	}

	BOOL bPauseScanning = FALSE;
	char cChoice = 0;
	do 
	{
		cChoice = ::_getch();
		if (g_bCustomScanCompleted)
		{
			break;
		}

		switch (cChoice)
		{
		case ' ':
			bPauseScanning = !bPauseScanning;

			if (bPauseScanning)
			{
				g_bNeedToPauseScanning = TRUE;
			}
			else
			{
				g_bNeedToResumeScanning = TRUE;
			}

			break;
		default:
			g_bCustomScanNeedToStop = TRUE;
			break;
		}
	} while (!g_bCustomScanNeedToStop && !g_bCustomScanCompleted);


	::WaitForSingleObject(
		hQueryThread,
		INFINITE
		);

Exit0:

	if (NULL != pCustomScan)
	{
		pCustomScan->Release();
	}
	return hResult;
}
