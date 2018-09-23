//////////////////////////////////////////////////////////////////////
//
//  @ File		:	testquickscan.h
//  @ Version	:	1.0
//  @ Author	:	EasyLogic <liangguangcai@kingsoft.com>
//  @ Datetime	:	[2010-6-2, 16:48:09]
//  @ Brief		:	测试快速扫描功能
//
//////////////////////////////////////////////////////////////////////

#pragma once
#include <windows.h>

//-------------------------------------------------------------------------

DWORD WINAPI QuickScanThreadProc(LPVOID lpParameter);

HRESULT TestQuickScan();
