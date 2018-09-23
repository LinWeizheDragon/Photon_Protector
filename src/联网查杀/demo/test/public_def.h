//////////////////////////////////////////////////////////////////////
//
//  @ File		:	public_def.h
//  @ Version	:	1.0
//  @ Author	:	EasyLogic <liangguangcai@kingsoft.com>
//  @ Datetime	:	[2010-6-2, 16:53:39]
//  @ Brief		:	公用函数及宏定义
//
//////////////////////////////////////////////////////////////////////
#pragma once
#include <windows.h>
#include <stdio.h>
#include <conio.h>

#include "ikcldcustomscanclient.h"

//-------------------------------------------------------------------------

#define PRINT_FUNCTION_CALL_ERR_MSG(func, hr)\
	do\
	{\
		if (FAILED(hr))\
		{\
			::printf(\
				"Call function [ %s ] failed with error code : %#010x\n",\
				func,\
				hr\
				);\
		}\
	} while (FALSE)

#define PROCESS_COM_ERROR(hr)\
	do\
	{\
		if (FAILED(hr))\
		{\
			goto Exit0;\
		}\
	} while (FALSE)

#define PROCESS_COM_ERROR_WITH_MSG(msg, hr)\
	do\
	{\
		if (FAILED(hr))\
		{\
			::printf("Error: %s\nError code:%#010x\n", msg, hr);\
		}\
		PROCESS_COM_ERROR(hr);\
	} while (FALSE)

#define PROCESS_COM_ERROR_WITH_FUNCTION_CALL_MSG(func, hr)\
	do\
	{\
		PRINT_FUNCTION_CALL_ERR_MSG(func, hr);\
		PROCESS_COM_ERROR(hr);\
	} while (FALSE)


