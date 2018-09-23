//////////////////////////////////////////////////////////////////////
//
//  @ File		:	testquickscan.cpp
//  @ Version	:	1.0
//  @ Author	:	EasyLogic <liangguangcai@kingsoft.com>
//  @ Datetime	:	[2010-6-2, 17:18:04]
//  @ Brief		:	测试快速扫描功能
//
//////////////////////////////////////////////////////////////////////
#include "stdafx.h"

#include "testquickscan.h"
#include "public_def.h"
#include "kscomdll.h"
#include <locale>

//-------------------------------------------------------------------------

#define QUERY_SESSION_STATUS_TIME_OUT		500

BOOL g_bQuickScanNeedToStop;

//-------------------------------------------------------------------------

DWORD WINAPI QuickScanThreadProc(LPVOID lpParameter)
{
	IKCldQuickScanClient* pQuickScan =
		reinterpret_cast<IKCldQuickScanClient*>(lpParameter);

	HRESULT hResult = E_FAIL;
	IKCldGetQuickScanStatusInfo*		pGetScanStatusInfo	= NULL;
	IKCldGetFixItemsInfo*		pGetFixItemsInfo	= NULL;
	S_KCLD_QUICK_SCAN_STATUS*	pScanStatus			= NULL;
	S_KCLD_FIX_ITEM_EX*			pFixItem			= NULL;
	DWORD						dwProcess			= 0;	// 进度
	DWORD						dwNewFixItemCount	= 0;	// 发现新威胁数目
	DWORD						dwTotalFixItemCount	= 0;	// 总威胁数
	DWORD						dwDisplayInterval	= 0;	// 显示间隔

	/*
	* 查询扫描状态直到扫描过程完成
	*/
	while (true)
	{
		if (g_bQuickScanNeedToStop)
		{
			::printf("****User stopped scanning process.****\n");
			hResult = pQuickScan->StopScan();
			PRINT_FUNCTION_CALL_ERR_MSG("StopScan", hResult);
			::printf("Press any key to quit.\n");
			::_getch();
			goto Exit0;
		}

		hResult = pQuickScan->QuerySessionStatus(
			&pGetScanStatusInfo,
			&pGetFixItemsInfo
			);
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"pQuickScan->QuerySessionStatus",
			hResult
			);

		if (NULL == pGetScanStatusInfo)
		{
			goto Exit0;
		}

		hResult = pGetScanStatusInfo->GetInfo(&pScanStatus);
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"pGetScanStatusInfo->GetInfo",
			hResult
			);

		// 显示进度
		if (0 != pScanStatus->dwTotalQuantity)
		{
			dwProcess = (100 *pScanStatus->dwFinishedQuantity) / pScanStatus->dwTotalQuantity;
		}

		// 显示扫描目标
		if (0 != pScanStatus->dwCurrentTargetCount)
		{
			dwDisplayInterval =
				QUERY_SESSION_STATUS_TIME_OUT / pScanStatus->dwCurrentTargetCount;
		}

		for (int i = 0; i != pScanStatus->dwCurrentTargetCount; ++i)
		{
			if (NULL != pScanStatus->ppCurrentTargets[i]->pszLocation
				&& L'\0' != pScanStatus->ppCurrentTargets[i]->pszLocation[0])
			{
				::wprintf(
					L"[Scan Progress:\t%%%-3u] [Scan Target:\t%ls]\n",
					dwProcess,
					pScanStatus->ppCurrentTargets[i]->pszLocation
					);
				::Sleep(dwDisplayInterval);
			}
			else if (NULL != pScanStatus->ppCurrentTargets[i]->pszFile
				&& L'\0' != pScanStatus->ppCurrentTargets[i]->pszFile[0])
			{
				::wprintf(
					L"[Scan Progress:\t%%%-3u] [Scan Target:\t%ls]\n",
					dwProcess,
					pScanStatus->ppCurrentTargets[i]->pszFile
					);
				::Sleep(dwDisplayInterval);
			}
		}

		// 显示发现的新威胁：
		if (NULL != pGetFixItemsInfo)
		{
			dwNewFixItemCount = 0;
			hResult = pGetFixItemsInfo->GetCount(&dwNewFixItemCount);
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"pGetFixItemsInfo->GetCount",
				hResult
				);

			dwTotalFixItemCount += dwNewFixItemCount;

			if (dwNewFixItemCount > 0)
			{
				::printf("******New threats found:\n");
				for (int j = 0; j != dwNewFixItemCount; ++j)
				{
					hResult = pGetFixItemsInfo->GetInfo(j, &pFixItem);
					PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
						"pGetFixItemsInfo->GetInfo",
						hResult
						);

					::wprintf(
						L"[%03u]. %ls\n",
						j+1,
						pFixItem->baseInfo.pszItemName
						);
				}
			}
		}

		// 检查是否扫描完毕
		if (em_KCLD_QUICK_ScanStatus_Complete ==
			pScanStatus->emSessionStatus)
		{
			::printf("\n\n********scanning completed********\n\n");
			::Sleep(500);
			break;
		}

		pGetScanStatusInfo->Release();
		pGetScanStatusInfo	= NULL;

		if (NULL != pGetFixItemsInfo)
		{
			pGetFixItemsInfo->Release();
			pGetFixItemsInfo	= NULL;
		}
	}

	/*
	* 处理
	*/
	if (dwTotalFixItemCount == 0)
	{
		::printf("No threat was found.\n");
	}
	else
	{
		::printf(
			"There are %u threat(s) been found.\n",
			dwTotalFixItemCount
			);

		S_KCLD_QUERY_SETTING	querySetting;
		IKCldGetFixItemsInfo*	pGetFixItemsInfoTmp = NULL;
		querySetting.dwStartIndex	= 0;
		querySetting.dwTotalCount	= dwTotalFixItemCount;
		hResult = pQuickScan->QuerySessionFixListEx(
			&querySetting,
			&pGetFixItemsInfoTmp
			);
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"pQuickScan->QuerySessionFixListEx",
			hResult
			);

		DWORD				dwFixCount	= 0;
		S_KCLD_FIX_ITEM_EX*	pFixItemEx	= NULL;
		hResult = pGetFixItemsInfoTmp->GetCount(&dwFixCount);
		PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
			"pGetFixItemsInfoTmp->GetCount",
			hResult
			);

		if (dwFixCount > 0)
		{
			::printf("Found threats:\n");

			S_KCLD_FIX_ITEM* pFixItems = new S_KCLD_FIX_ITEM[dwFixCount];
			if (NULL == pFixItems)
			{
				hResult = E_OUTOFMEMORY;
				PROCESS_COM_ERROR(hResult);
			}

			for (int i = 0; i != dwFixCount; ++i)
			{
				pFixItemEx = NULL;
				hResult = pGetFixItemsInfoTmp->GetInfo(i,  &pFixItemEx);
				PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
					"pGetFixItemsInfoTmp->GetInfo",
					hResult
					);

				pFixItems[i].dwID			= pFixItemEx->baseInfo.dwID;
				pFixItems[i].emActionType	= pFixItemEx->baseInfo.emActionType;
				pFixItems[i].emAdvice		= pFixItemEx->baseInfo.emAdvice;
				pFixItems[i].emLevel		= pFixItemEx->baseInfo.emLevel;
				pFixItems[i].emType			= pFixItemEx->baseInfo.emType;
				pFixItems[i].pszItemName	= pFixItemEx->baseInfo.pszItemName;

				::wprintf(L"[%03u]. %ls\n", i+1, pFixItems[i].pszItemName);
			}

			::printf("Start fix, please wait...\n");

			hResult = pQuickScan->StartFix(
				dwFixCount,
				pFixItems
				);
			PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
				"StartFix",
				hResult
				);

			::printf("Fix completed.\n");

			::printf("Query if need to reboot system...\n");

			BOOL bNeedToReboot = FALSE;
			hResult = pQuickScan->QueryNeedReboot(&bNeedToReboot);
			if (!bNeedToReboot)
			{
				::printf("Do not need.\n");
			}
			else
			{
				::printf("Need reboot, do you want to restart your computer[ Y/N ]?\n");
				if ('Y' == ::toupper(::_getch()))
				{
					::printf("Press any key to reboot.\n");
					::_getch();

					hResult = pQuickScan->Reboot();
					PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
						"pQuickScan->Reboot",
						hResult
						);
				}
			}
		}
	}

	::printf("Press any key to quit.\n");

Exit0:
	
	if (NULL != pGetScanStatusInfo)
	{
		pGetScanStatusInfo->Release();
		pGetScanStatusInfo = NULL;
	}

	if (NULL != pGetFixItemsInfo)
	{
		pGetFixItemsInfo->Release();
		pGetFixItemsInfo = NULL;
	}

	return hResult;
}

HRESULT TestQuickScan()
{
	HRESULT hResult = E_FAIL;
	IKCldQuickScanClient* pQuickScan = NULL;
	KSCOMDll scom_dll;

	hResult = scom_dll.Open(L"kcldscan.dll");
	PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
		"scom_dll.Open",
		hResult
		);

	scom_dll.GetClassObject(
		CLSID_KCldQuickScanClient,
		IID_IKCldQuickScanClient,
		(void**)&pQuickScan
		);
	if (NULL == pQuickScan)
		return 0;

	::printf("Connecting to kingsoft cloud security server end...\n");
	int nState = 0;
	hResult = pQuickScan->QueryConnectToCloudState(&nState);
	PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG("QueryConnectToCloudState", hResult);
	if (1 == nState)
	{
		::printf("It succeeded to connect to kingsoft cloud security server end.\n");
	}
	else if (0 == nState)
	{
		::printf("It failed to connect to kingsoft cloud security server end.\n");
	}
	else if (2 == nState)
	{
		::printf("Busy now.\n");
	}

	::printf("***************************NOTE***************************\n");
	::printf("**\tAfter scanning started,\n");
	::printf("**\tyou can press any key to stop scanning process.\n");
	::printf("**********************************************************\n");
	::printf("Now press any key to Start scan.\n");
	::_getch();

	::printf("\nStartScan...\n");
	hResult = pQuickScan->StartScan(enum_KCLD_All);
	PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(
		"StartScan",
		hResult
		);


	/*
	* 创建线程以查询扫描状态
	*/
	HANDLE hQueryThread = ::CreateThread(
		NULL,
		0,
		QuickScanThreadProc,
		pQuickScan,
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

	::_getch();

	g_bQuickScanNeedToStop = TRUE;

	::WaitForSingleObject(
		hQueryThread,
		INFINITE
		);

Exit0:

	if (NULL != pQuickScan)
	{
		pQuickScan->Release();
	}

	if (FAILED(hResult))
	{
		::printf("Press any key to quit.\n");
		::_getch();
	}

	return hResult;
}
