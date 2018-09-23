//=============================================================================
/**
* @file ikcldcustomscanclient.h
* @brief 
* @author qiuruifeng <qiuruifeng@kingsoft.com>
* @date 2010-5-28   11:32
*/
//=============================================================================
#pragma once 

#include <Unknwn.h>
#include "kcldcustomscandata_def.h"

//////////////////////////////////////////////////////////////////////////

// {A63AD1AB-1AB4-4f29-A7C6-69BD3916E563}
extern "C" const __declspec(selectany)GUID IID_IKCldGetFullScanStatusInfo =
{ 0xa63ad1ab, 0x1ab4, 0x4f29, { 0xa7, 0xc6, 0x69, 0xbd, 0x39, 0x16, 0xe5, 0x63 } };
MIDL_INTERFACE("{A63AD1AB-1AB4-4f29-A7C6-69BD3916E563}")
IKCldGetFullScanStatusInfo : public IUnknown
{
public:
	/**
	* @brief		获取全盘扫描状态信息
	* @remark		外部不需要释放内存，组件释放后会自动释放内存
	* @param[out]	ppInfo 保存扫描状态信息
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE GetInfo(
		/*[out]*/	S_KCLD_FULL_SCAN_STATUS** ppInfo
		) = 0;
};

// {98001E17-818E-4dd1-B79B-13EA88E03460}
extern "C" const __declspec(selectany)GUID IID_IKCldGetCustomScanStatusInfo =
{ 0x98001e17, 0x818e, 0x4dd1, { 0xb7, 0x9b, 0x13, 0xea, 0x88, 0xe0, 0x34, 0x60 } };
MIDL_INTERFACE("{98001E17-818E-4dd1-B79B-13EA88E03460}")
IKCldGetCustomScanStatusInfo : public IUnknown
{
public:
	/**
	* @brief		获取全盘扫描状态信息
	* @remark		外部不需要释放内存，组件释放后会自动释放内存
	* @param[out]	ppInfo 保存扫描状态信息
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE GetInfo(
		/*[out]*/	S_KCLD_CUSTOM_SCAN_STATUS** ppInfo
		) = 0;
};


// {E8FBEA13-5BEA-42e3-AB9A-FC96D4EE917B}
extern "C" const __declspec(selectany)GUID IID_IKCldFileVirusThreatsInfo =
{ 0xe8fbea13, 0x5bea, 0x42e3, { 0xab, 0x9a, 0xfc, 0x96, 0xd4, 0xee, 0x91, 0x7b } };
MIDL_INTERFACE("{E8FBEA13-5BEA-42e3-AB9A-FC96D4EE917B}")
IKCldGetFileVirusThreatsInfo : public IUnknown
{
public:
	/**
	* @brief		获取到文件病毒威胁项列表的长度
	* @remark		调用 GetCount 获得文件病毒威胁项列表的长度后，就可以调用
	*				GetInfo 并传入索引来取出威胁项列表中的数据
	* @param[out]	pdwCount 保存文件病毒威胁项列表的长度
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE GetCount(DWORD* pdwCount) = 0;

	/**
	* @brief		索引文件病毒威胁项列表中某个项目
	* @remark		直接取得组件中保存的可修复列表中的项目，
	*				外部不需要释放内存，组件释放后会自动释放内存
	* @param[in]	dwIndex 文件病毒列表索引，从 0 开始
	* @param[out]	ppInfo 保存该项目的数据信息
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE GetInfo(
		/*[in]*/	const DWORD dwIndex,
		/*[out]*/	KCLD_FILEVIRUS_THREAT** ppInfo
		) = 0;
};

// {635941CF-456F-454c-8C1D-18F877C6B046}
extern "C" const __declspec(selectany)GUID IID_IKCldGetProcessResultInfo =
{ 0x635941cf, 0x456f, 0x454c, { 0x8c, 0x1d, 0x18, 0xf8, 0x77, 0xc6, 0xb0, 0x46 } };
MIDL_INTERFACE("{635941CF-456F-454c-8C1D-18F877C6B046}")
IKCldGetProcessResultInfo : public IUnknown
{
public:
	/**
	* @brief		获取威胁的处理结果列表的长度
	* @remark		调用 GetCount 获得威胁的处理结果列表的长度后，就可以调用
	*				GetInfo 并传入索引来取出威胁的处理结果列表中的数据
	* @param[out]	pdwCount 保存威胁的处理结果列表的长度
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE GetCount(DWORD* pdwCount) = 0;

	/**
	* @brief		索引威胁的处理结果列表中某个项目
	* @remark		外部不需要释放内存，组件释放后会自动释放内存
	* @param[in]	dwIndex 威胁的处理结果索引，从 0 开始
	* @param[out]	ppInfo 保存该项目的数据信息
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE GetInfo(
		/*[in]*/	const DWORD dwIndex,
		/*[out]*/	S_KCLD_THREAT_PROCESS_RESULT** ppInfo
		) = 0;
};

// {DBD0490A-E102-4b84-BC0D-A725FE077E3F}
extern "C" const __declspec(selectany)GUID CLSID_KCldCustomScanClient =
{ 0xdbd0490a, 0xe102, 0x4b84, { 0xbc, 0xd, 0xa7, 0x25, 0xfe, 0x7, 0x7e, 0x3f } };

// {14EE3F04-F87E-4361-AA27-BB7E7DC225D5}
extern "C" const __declspec(selectany)GUID IID_IKCldCustomScanClient =
{ 0x14ee3f04, 0xf87e, 0x4361, { 0xaa, 0x27, 0xbb, 0x7e, 0x7d, 0xc2, 0x25, 0xd5 } };


MIDL_INTERFACE("{14EE3F04-F87E-4361-AA27-BB7E7DC225D5}")
IKCldCustomScanClient : public IUnknown
{
public:
	/**
	* @brief		保留
	* @remark		
	* @param[in]	pReserved 保留参数
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE Init(LPVOID ) = 0;

	/**
	* @brief		启动全盘扫描
	* @remark		
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE StartFullScan() = 0;

	/**
	* @brief		自定义扫描时添加扫描路径
	* @remark		
	* @param[in]	pszPath 待添加的扫描路径
	* @param[in]	bScanSubDir 为真则扫描路径 pszPath 下的子目录
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE AppendScanTargetPath(
		/*[in]*/	const wchar_t*					pszPath,
		/*[in]*/	const BOOL						bScanSubDir = TRUE
		) = 0;

	/**
	* @brief		启动自定义扫描
	* @remark		先调用 AppendScanTargetPath 添加扫描路径，再调用此
	*				接口开始自定义扫描
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE StartCustomScan() = 0;

	/**
	* @brief		停止扫描
	* @remark		
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE StopScan() = 0;

	/**
	* @brief		暂停扫描
	* @remark		
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE PauseScan() = 0;

	/**
	* @brief		恢复扫描
	* @remark		
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE ResumeScan() = 0;

	/**
	* @brief		查询全盘扫描状态信息
	* @remark		
	* @param[out]	ppScanStatus 保存查询到的状态信息的组件接口指针的地址
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE QueryFullScanStatus(
		/*[out]*/	IKCldGetFullScanStatusInfo**	ppStatusInfo
		) = 0;

	/**
	* @brief		查询自定义扫描状态信息
	* @remark		
	* @param[out]	ppScanStatus 保存查询到的状态信息的组件接口指针的地址
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE QueryCustomScanStatus(
		/*[out]*/	IKCldGetCustomScanStatusInfo**	ppStatusInfo
		) = 0;

	/**
	* @brief		查询扫描中发现的文件病毒威胁项列表
	* @remark		
	* @param[in]	pQuerySetting 需要查询的设置
	* @param[out]	ppThreatsInfo 返回威胁信息的组件接口指针的地址
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE QueryFileVirusThreats(
		/*[in]*/	const KCLD_QUERY_THREAT*		pQuerySetting,
		/*[out]*/	IKCldGetFileVirusThreatsInfo**	ppThreatsInfo
		) = 0;

	/**
	* @brief		对发现的威胁进行处理
	* @remark		
	* @param[in]	pProcessScanTarget 需要处理的威胁
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE ProcessScanResult(
		/*[in]*/	const KCLD_PROCESS_SCAN_TARGET*	pProcessScanTarget
		) = 0;

	/**
	* @brief		查询扫描到的威胁的处理结果
	* @remark		
	* @param[in]	uThreatCount 要查询的威胁 ID 列表长度
	* @param[in]	pThreatIDs 要查询的威胁 ID 列表
	* @param[out]	ppResultInfo 接收威胁的处理结果
	* @return		0 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE QueryScanThreatProcessResult(
		/*[in]*/	unsigned int					uThreatCount,
		/*[in]*/	const DWORD*					pThreatIDs,
		/*[out]*/	IKCldGetProcessResultInfo**		ppResultInfo
		);

	/**
	* @brief		检查是否需要重启
	* @remark		
	* @param[out]	pbNeedReBoot 标示是否需要重启
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE QueryNeedReboot(
		/*[out]*/	BOOL* pbNeedReBoot
		) = 0;

	/**
	* @brief		保留
	* @remark		
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE UnInit() = 0;

	/**
	* @brief		增加进程扫描的目标
	* @remark		
	* @param[in]	pTarget 扫描目标
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE AppendScanProcessTarget(
		/*[in]*/	const S_KCLD_SCAN_PROCESS_TARGET* pTarget
		) = 0;

	/**
	* @brief		控制是否上报样本
	* @remark		
	* @param[in]	bAutoUploadFile true 为允许上报, false 为禁止自动上报
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE SetAutoUploadFile(
		/*[in]*/	BOOL bAutoUploadFile
		) = 0;

	/**
	* @brief		获取是否上报样本
	* @remark		
	* @param[in]	bAutoUploadFile true 为允许上报, false 为禁止自动上报
	* @return		S_OK 成功，其他为错误码
	**/
	virtual HRESULT STDMETHODCALLTYPE GetAutoUploadFile(
		/*[out]*/	BOOL& bAutoUploadFile
		) = 0;
};



