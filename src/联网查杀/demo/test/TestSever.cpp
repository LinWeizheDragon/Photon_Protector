//////////////////////////////////////////////////////////////////////
//
//  @ File		:	TestSever.cpp
//  @ Version	:	1.0
//  @ Author	:	EasyLogic <liangguangcai@kingsoft.com>
//  @ Datetime	:	[2010-6-5, 18:34:43]
//  @ Brief		:	²âÊÔ¼ÓÔØ¡¢Ð¶ÔØ SP
//
//////////////////////////////////////////////////////////////////////
#include "stdafx.h"

#include "TestSever.h"
#include "ikcldscanserver.h"
#include "public_def.h"
#include "kscomdll.h"

HRESULT TestServer(EM_TEST_TYPE testType)
{
	HRESULT hResult	= E_FAIL;
	IKCldScanServer* pscanServer = NULL;
	KSCOMDll scom_dll;

	hResult = scom_dll.Open(L"kcldscan.dll");
	PROCESS_COM_ERROR(hResult);

	hResult = scom_dll.GetClassObject(
		CLSID_KCldScanServer,
		IID_IKCldScanServer,
		(void**)&pscanServer
		);
	if (NULL == pscanServer)
	{
		return E_FAIL;
	}
	
	hResult = pscanServer->Init(NULL);
	PRINT_FUNCTION_CALL_ERR_MSG("Init", hResult);
	PROCESS_COM_ERROR(hResult);

	if ((enum_Test_Quick_Scan & testType) == enum_Test_Quick_Scan)
	{
		hResult = pscanServer->StartQuickScanSerice();
		PRINT_FUNCTION_CALL_ERR_MSG("StartQuickScanSerice", hResult);
		PROCESS_COM_ERROR(hResult);
	}
	
	if ((enum_Test_Custom_Scan & testType) == enum_Test_Custom_Scan)
	{
		hResult = pscanServer->StartCustomScanSerice();
		PRINT_FUNCTION_CALL_ERR_MSG("StartCustomScanSerice", hResult);
		PROCESS_COM_ERROR(hResult);
	}

	::printf("Press any key to Stop.\n");

	::_getch();

	if ((enum_Test_Quick_Scan & testType) == enum_Test_Quick_Scan)
	{
		hResult = pscanServer->StopQuickScanSerice();
		PRINT_FUNCTION_CALL_ERR_MSG("StopQuickScanSerice", hResult);
		PROCESS_COM_ERROR(hResult);
	}

	if ((enum_Test_Custom_Scan & testType) == enum_Test_Custom_Scan)
	{
		hResult = pscanServer->StopCustomScanSerice();
		PRINT_FUNCTION_CALL_ERR_MSG("StopCustomScanSerice", hResult);
		PROCESS_COM_ERROR(hResult);
	}

	hResult = pscanServer->UnInit();
	PRINT_FUNCTION_CALL_ERR_MSG("UnInit", hResult);
	PROCESS_COM_ERROR(hResult);

Exit0:

	if (NULL != pscanServer)
	{
		pscanServer->Release();
	}

	return hResult;
}