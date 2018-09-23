//=============================================================================
/**
* @file kcldcustomscandata_def.h
* @brief 
* @author qiuruifeng <qiuruifeng@kingsoft.com>
* @date 2010-5-28   14:48
*/
//=============================================================================
#pragma once 

/**
* @brief 威肋类型定义
*/
typedef enum _EM_KCLD_CUSTOM_SCAN_THREAT_TYPE
{
	em_KCLD_ThreatType_MalwareInSystem = 0,			///< 系统恶意软件
	em_KCLD_ThreatType_FileVirus,					///< 文件病毒
	em_KCLD_ThreatType_SystemRepairPoint,			///< 系统修复点
	em_KCLD_Max_Threat_Type = 16					///< 威胁类型的最大值
} EM_KCLD_CUSTOM_SCAN_THREAT_TYPE;


/**
* @brief ScanSession的状态
*/
typedef enum _EM_KCLD_CUSTOM_SCAN_STATUS
{
	em_KCLD_CUSTOM_ScanStatus_None          = 0,	///< 无效状态
	em_KCLD_CUSTOM_ScanStatus_Ready         = 1,	///< 任务就绪,准备运行(短暂的状态)
	em_KCLD_CUSTOM_ScanStatus_Running       = 2,	///< 任务正在运行
	em_KCLD_CUSTOM_ScanStatus_Paused        = 3,	///< 任务被暂停
	em_KCLD_CUSTOM_ScanStatus_Complete      = 4,	///< 任务完成(有可能是被中止导致的完成)
	em_KCLD_CUSTOM_ScanStatus_NetDetecting  = 5,	///< 等待网络检测
} EM_KCLD_CUSTOM_SCAN_STATUS;

/**
 * @brief 指定需要扫描的目标类型
 */
typedef enum _EM_KCLD_SCANTARGET_TYPE
{
    em_KCLD_Target_None              = 0x00000000,	///< 无效目标
    em_KCLD_Target_Win32_Directory   = 0x00010003,	///< Win32目录,TargetName为路径字符串，TargetID的取值为：0（包含子目录）1（不包含子目录）		 
	em_KCLD_Target_Win32_File        = 0x00010005,	///< Win32文件,需要指定TargetName为要扫描文件的全路径，TargetID为空。	
	em_KCLD_Target_All_Malware       = 0x00010200,	///< 所有恶意软件对象，TargetName为空，TargetID为空
	em_KCLD_Target_Malware           = 0x00010201,	///< 恶意软件对象,TargetName为空，需要指定TargetID为恶意软件ID
	em_KCLD_Target_All_SysRprPoints,				///< 所有系统修复点，TargetName为空，TargetID为空
	em_KCLD_Target_SysRprPoints,					///< 系统修复点,TargetName为空，需要指定TargetID为系统修复点ID
	em_KCLD_Target_Computer,						///< 我的电脑，TargetName为空，TargetID为空
	em_KCLD_Target_Autoruns,						///< 自启动对象，TargetName为空，TargetID为空
	em_KCLD_Target_Critical_Area,					///< 关键区域，TargetName为空，TargetID为空
    em_KCLD_Target_Removable,						///< 移动存储设备
	em_KCLD_Target_Memory							///< 内存扫描
} EM_KCLD_SCANTARGET_TYPE;

/**
 * @brief 需要扫描的目标内容
 */
typedef struct _S_KCLD_CUSTOM_SCAN_TARGET
{
	EM_KCLD_SCANTARGET_TYPE		emTargetType;		///< 需要扫描的类型
	const wchar_t*				pszTargetName;		///< 需要扫描的名字,与emTargetType相关
	DWORD						dwTargetID;			///< 需要扫描的ID,与emTargetType相关
} S_KCLD_CUSTOM_SCAN_TARGET;

/**
* @brief 获取进度信息
*/
typedef struct _S_KCLD_PROGRESS_INFO
{
	DWORD dwTimeProgress;							///< 剩余的时间(ms)
	DWORD dwStreamProgress;							///< 获取基于数据流的进度  (0 - 100)      	
	DWORD dwFilesProgress;							///< 获取基于文件数的进度  (0 - 100)
	DWORD dwFilesTotalCount;						///< 获取扫描文件的总数	
} S_KCLD_PROGRESS_INFO;

/**
 * @brief 当前正在扫描的状态
 */
typedef struct _S_KCLD_CUSTOM_SCAN_STATUS
{
	S_KCLD_CUSTOM_SCAN_TARGET	currentTarget;				///< 当前正在扫描关键区域，见发起扫描的目标定义
	EM_KCLD_CUSTOM_SCAN_STATUS  emSessionStatus;			///< 任务状态
	S_KCLD_PROGRESS_INFO		progress;					///< 进度信息
	DWORD                       dwTotalQuantity;			///< 总的任务量
	DWORD                       dwFinishedQuantity;			///< 完成的任务量
	DWORD                       dwFindThread;				///< 发现威胁的总数
	DWORD                       dwProcessSucceed;			///< 成功处理威胁的总数
	DWORD						FoundThreatsCountDetail[em_KCLD_Max_Threat_Type]; ///< 根据不同的威胁种类,告诉相关的数量 
	__time64_t		            tmSessionStartTime;             ///< 扫描启动的时间点
	__time64_t		            tmSessionCurrentTime;           ///< 当前时间点,用于计算已经扫描了多少时间
	__time64_t                  tmSessionEndTime;               ///< 扫描结束的时间点（可能是完成，也可能是被中断，原因见SessionSatus）
    HRESULT                     hErrCode;					///< 错误码
} S_KCLD_CUSTOM_SCAN_STATUS;

/**
* @brief 威助的处理结果
*/
typedef enum _EM_KCLD_THREAT_PROCESS_RESULT
{
	em_KCLD_Threat_Process_No_Op                            =  0x00000001,   ///< 未处理
	em_KCLD_Threat_Process_Unknown_Fail                     =  0x80000001,   ///< 未知失败

	em_KCLD_Threat_Process_Delay                            =  0x00000002,   ///< 延迟处理
	em_KCLD_Threat_Process_Skip                             =  0x00000003,   ///< 跳过

	// 文件相关的结果
	em_KCLD_Threat_Process_Clean_File_Success               =  0x00000004,   ///< 清除(修复)文件成功
	em_KCLD_Threat_Process_Clean_File_Fail                  =  0x80000004,   ///< 清除(修复)文件失败
	em_KCLD_Threat_Process_Delete_File_Success              =  0x00000005,   ///< 删除文件成功
	em_KCLD_Threat_Process_Delete_File_Fail                 =  0x80000005,   ///< 删除文件失败

	em_KCLD_Threat_Process_Reboot_Clean_File_Success        =  0x00000006,   ///< 重启后清除文件(调用成功)
	em_KCLD_Threat_Process_Reboot_Clean_File_Fail           =  0x80000006,   ///< 重启后清除文件(调用失败)
	em_KCLD_Threat_Process_Reboot_Delete_File_Success       =  0x00000007,   ///< 重启后删除文件(调用成功)
	em_KCLD_Threat_Process_Reboot_Delete_File_Fail          =  0x80000007,   ///< 重启后删除文件(调用失败)
	em_KCLD_Threat_Process_Rename_File_Success              =  0x00000008,   ///< 重命名文件成功
	em_KCLD_Threat_Process_Rename_File_Fail                 =  0x80000008,   ///< 重命名文件失败

	em_KCLD_Threat_Process_Quarantine_File_Success          =  0x00000009,   ///< 隔离文件成功
	em_KCLD_Threat_Process_Quarantine_File_Fail             =  0x80000009,   ///< 隔离文件失败
	em_KCLD_Threat_Process_Reboot_Quarantine_File_Success   =  0x0000000A,   ///< 重启后隔离文件(调用成功)
	em_KCLD_Threat_Process_Reboot_Quarantine_File_Fail      =  0x8000000A,   ///< 重启后隔离文件(调用失败)

	em_KCLD_Threat_Process_File_NoExist					   =  0x81000001,    ///< 文件不存在

	em_KCLD_Threat_Process_File_InWPL					   =  0x00000020,    ///< 文件在纯白名单中
	em_KCLD_Threat_Process_File_Restore_Success			   =  0x00000021,    ///< 该文件已经被恢复
} EM_KCLD_THREAT_PROCESS_RESULT;

/**
* @brief   病毒的查询类型
*/
typedef enum _EM_KCLD_THREAT_CHECKING_TYPE
{
	em_KCLD_ThreatCheckedByFileEngine,				///< 文件引擎报毒
	em_KCLD_ThreatCheckedByCloud,					///< 云查杀报毒
} EM_KCLD_THREAT_CHECKING_TYPE;

/**
* @brief   压缩,脱壳文件数据结构
*/
typedef struct _KCLD_ARCHIVE_THREAT
{
	const wchar_t*					pszFileName;			///< 压缩,文件名
	const wchar_t*					pszFullPath;			///< 虚拟全路径
	const wchar_t*					pszThreatName;			///< 威胁名称
	EM_KCLD_THREAT_PROCESS_RESULT	eResult;				///< 处理结果
	DWORD							dwThreatType;			///< 病毒类型
	EM_KCLD_THREAT_CHECKING_TYPE	eThreatCheckingType;	///< 报毒类型
} KCLD_ARCHIVE_THREAT;

typedef struct _KCLD_FILEVIRUS_THREAT
{
	DWORD							dwFoundThreatIndex;		///< 标示了机器上发现的一个威助
	DWORD							dwThreatID;				///< 此ID为威胁库内存在的ID
	bool							bFileInWrapper;			///< 是否是压缩包中的文件(不包括RTF)
	const wchar_t*					pszFileFullPath;		///< 病毒文件的全路径
	const wchar_t*					pszVirusDescription;	///< 病毒描述
	EM_KCLD_THREAT_PROCESS_RESULT	eResult;                ///< 处理结果
	HRESULT							hErrCode;               ///< 错误码
	DWORD							dwThreatType;			///< 病毒类型
	EM_KCLD_THREAT_CHECKING_TYPE	eThreatCheckingType;    ///< 报毒类型
	DWORD							dwVirtualFileCount;		///< 压缩,脱壳文件列表长度
	KCLD_ARCHIVE_THREAT**			ppVirtualFiles;			///< 压缩,脱壳文件列表
}KCLD_FILEVIRUS_THREAT;

/**
* @brief 查询数据的起始索引以及查询数量
*/
typedef struct _KCLD_QUERY_THREAT
{
	DWORD			dwStartIndex;					///<要查询威胁的起始索引
	DWORD			dwTotalCount;					///<本次查询最多返回的数量
}KCLD_QUERY_THREAT;

/**
* @brief 用于对指定的ScanHandle处理对应的威胁内容
*/
typedef struct _KCLD_PROCESS_SCAN_TARGET
{
	BOOL		bClearAllThread;					///< 是否清除所有威胁
	DWORD		dwThreatIndexCount;					///< 威胁列表大小
	DWORD*		pdwThreatIndexList;					///< 提交处理请求的威胁索引的列表（威胁索引在查询威胁是获得）
}KCLD_PROCESS_SCAN_TARGET;

//-------------------------------------------------------------------------

/**
 * @brief ScanSession的状态
 */
typedef enum _EM_KCLD_SCANSESSION_STATUS
{
    em_KCLD_ScanStatus_None          = 0,			///< 无效状态
	em_KCLD_ScanStatus_Ready         = 1,			///< 任务就绪,准备运行(短暂的状态)
	em_KCLD_ScanStatus_Running       = 2,			///< 任务正在运行
	em_KCLD_ScanStatus_Paused        = 3,			///< 任务被暂停
	em_KCLD_ScanStatus_Complete      = 4,			///< 任务完成(有可能是被中止导致的完成)
	em_KCLD_ScanStatus_NetDetecting  = 5,			///< 等待网络检测
} EM_KCLD_SCANSESSION_STATUS;

/**
* @brief 磁盘扫描项统计信息
*/
typedef struct _S_KCLD_DISKSCANITEM_INFO
{
	const wchar_t*	pszDriverName;					///< 驱动器名称
	DWORD			dwScanItems;					///< 已经扫描项的总数
	DWORD			dwFindThreats;					///< 发现的病毒数
    int				nStatus;						///< 表示当前的状态  参考EM_KXE_SCAN_DISK_STATUS
	
} S_KCLD_DISKSCANITEM_INFO;

/**
* @brief 全盘扫描的时候，状态信息的查询
*/
typedef struct _S_KCLD_FULL_SCAN_STATUS
{
	S_KCLD_CUSTOM_SCAN_STATUS	scanStatus;
	DWORD						dwScanItemInfoCount;///< 磁盘扫描项统计信息列表长度
	S_KCLD_DISKSCANITEM_INFO**	ppDiskInfo;			///< 磁盘扫描项统计信息列表
} S_KCLD_FULL_SCAN_STATUS;

/**
* @brief 威胁处理结果状态
*/
typedef struct _S_KCLD_THREAT_PROCESS_RESULT
{
    DWORD							dwThreadIndex;	///< 威胁索引ID
    EM_KCLD_THREAT_PROCESS_RESULT	eResult;		///< 处理结果
	DWORD							dwOtherCount;	///< 压缩包,脱壳文件处理结果数目
    EM_KCLD_THREAT_PROCESS_RESULT**	ppOtherResults;	///< 压缩包,脱壳文件处理结果
} S_KCLD_THREAT_PROCESS_RESULT;

/**
* @brief 进程扫描类型
**/
typedef enum _EM_KCLD_SCAN_PROCESS_TARGET_TYPE
{
	em_KCLD_ScanProcessInvaild	= 0,			///< 无效状态
	em_KCLD_ScanProcessAll		= 1,			///< 扫描所有进程
	em_KCLD_ScanProcessByPid	= 2,			///< 扫描指定pid进程
	em_KCLD_ScanProcessByName	= 3,			///< 扫描指定pid进程
	em_KCLD_ScanDrivers			= 4				///< 扫描drivers目录下的sys文件
}EM_KCLD_SCAN_PROCESS_TARGET_TYPE;

/**
* @brief 进程扫描项
**/
typedef struct _S_KCLD_SCAN_PROCESS_TARGET_ITEM
{
	DWORD								dwSize;		///< 结构体大小
	EM_KCLD_SCAN_PROCESS_TARGET_TYPE	emScanType; ///< 扫描类型
	BOOL								bScanModule;///< 是否扫描进程下的模块	
	DWORD								dwPid;		///< 进程id,仅扫描类型为em_KCLD_ScanProcessByPid时有效
	const wchar_t*						pszName;	///< 进程名,仅扫描类型为em_KCLD_ScanProcessByName时有效
}S_KCLD_SCAN_PROCESS_TARGET_ITEM;

/**
* @brief 进程扫描参数
**/
typedef struct _S_KCLD_SCAN_PROCESS_TARGET
{
	DWORD								dwSize;		///< 结构体大小
	S_KCLD_SCAN_PROCESS_TARGET_ITEM**	ppItems;	///< 扫描项指针数组
	DWORD								dwItemsCnt;	///< 扫描项数组大小	
}S_KCLD_SCAN_PROCESS_TARGET;