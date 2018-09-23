//////////////////////////////////////////////////////////////////////
//
//  @ File		:	TestSever.h
//  @ Version	:	1.0
//  @ Author	:	EasyLogic <liangguangcai@kingsoft.com>
//  @ Datetime	:	[2010-6-5, 17:25:53]
//  @ Brief		:	测试加载、卸载 SP
//
//////////////////////////////////////////////////////////////////////
#pragma once
#include <windows.h>

typedef enum _EM_TEST_TYPE
{
	enum_Test_Quick_Scan	= 0x00000001,	///< 测试快速
	enum_Test_Custom_Scan	= 0x00000010,	///< 测试全盘及自定义
	enum_Test_All_Scan		= 0x00000011,	///< 测试全部（全盘&快速）
} EM_TEST_TYPE;

HRESULT TestServer(EM_TEST_TYPE testType);