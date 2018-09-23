// IniControl.cpp: implementation of the IniControl class.
//
//////////////////////////////////////////////////////////////////////

//#include "stdafx.h"
//#include "prmdlg.h"
#include "IniControl.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#define MAX_ALLSECTIONS 2048  //全部的段名
#define MAX_SECTION 260  //一个段名长度
#define MAX_ALLKEYS 6000  //全部的键名
#define MAX_KEY 260  //一个键名长度
 
//----------------------------------------------------------------------------------
/*
 类名：CIni
 版本：v2.0
 最后更新：
 v2.0
 梦小孩于2004年2月14日情人节
 加入高级操作的功能
 v1.0
 梦小孩于2003年某日
 一般操作完成
 
 类描述：
 本类可以于.ini文件进行操作
 */
 
//文件 1:

#pragma once
 
//#include "afxTempl.h"
 
class CIni
{
private:
 CString m_strFileName;
public:
 CIni(CString strFileName):m_strFileName(strFileName)
 {
 }
public:
 //一般性操作：
 BOOL SetFileName(LPCTSTR lpFileName);  //设置文件名
 CString GetFileName(void); //获得文件名
 BOOL SetValue(LPCTSTR lpSection, LPCTSTR lpKey, LPCTSTR lpValue,bool bCreate=true); //设置键值，bCreate是指段名及键名未存在时，是否创建。
 CString GetValue(LPCTSTR lpSection, LPCTSTR lpKey); //得到键值.
 BOOL DelSection(LPCTSTR strSection);  //删除段名
 BOOL DelKey(LPCTSTR lpSection, LPCTSTR lpKey);  //删除键名
 

public:
 //高级操作：
 int GetSections(CStringArray& arrSection);  //枚举出全部的段名
 int GetKeyValues(CStringArray& arrKey,CStringArray& arrValue,LPCTSTR lpSection);  //枚举出一段内的全部键名及值
 
 BOOL DelAllSections();
 
};



BOOL CIni::SetFileName(LPCTSTR lpFileName)
{
 CFile file;
 CFileStatus status;
 
 if(!file.GetStatus(lpFileName,status))
  return TRUE;
 
 m_strFileName=lpFileName;
 return FALSE;
}
 
CString CIni::GetFileName(void)
{
 return m_strFileName;
}
 
BOOL CIni::SetValue(LPCTSTR lpSection, LPCTSTR lpKey, LPCTSTR lpValue,bool bCreate)
{
 TCHAR lpTemp[MAX_PATH] ={0};
 
 //以下if语句表示如果设置bCreate为false时，当没有这个键名时则返回TRUE（表示出错）
 //!*&*none-value*&!* 这是个垃圾字符没有特别意义，这样乱写是防止凑巧相同。
 if (!bCreate)
 {
  GetPrivateProfileString(lpSection,lpKey,"!*&*none-value*&!*",lpTemp,MAX_PATH,m_strFileName);
  if(strcmp(lpTemp,"!*&*none-value*&!*")==0)
   return TRUE;
 }
 
 if(WritePrivateProfileString(lpSection,lpKey,lpValue,m_strFileName))
  return FALSE;
 else
  return GetLastError();
}
 
CString CIni::GetValue(LPCTSTR lpSection, LPCTSTR lpKey)
{
 DWORD dValue;
 TCHAR lpValue[MAX_PATH] ={0};
 
 dValue=GetPrivateProfileString(lpSection,lpKey,"",lpValue,MAX_PATH,m_strFileName);
 return lpValue;
}
 
BOOL CIni::DelSection(LPCTSTR lpSection)
{
 if(WritePrivateProfileString(lpSection,NULL,NULL,m_strFileName))
  return FALSE;
 else
  return GetLastError();
}
 
BOOL CIni::DelKey(LPCTSTR lpSection, LPCTSTR lpKey)
{
 if(WritePrivateProfileString(lpSection,lpKey,NULL,m_strFileName))
  return FALSE;
 else
  return GetLastError();
}
 

int CIni::GetSections(CStringArray& arrSection)
{
 /*
 本函数基础：
 GetPrivateProfileSectionNames - 从 ini 文件中获得 Section 的名称
 如果 ini 中有两个 Section: [sec1] 和 [sec2]，则返回的是 'sec1',0,'sec2',0,0 ，当你不知道  
 ini 中有哪些 section 的时候可以用这个 api 来获取名称 
 */
 int i;  
 int iPos=0;  
 int iMaxCount;
 TCHAR chSectionNames[MAX_ALLSECTIONS]={0}; //总的提出来的字符串
 TCHAR chSection[MAX_SECTION]={0}; //存放一个段名。
 GetPrivateProfileSectionNames(chSectionNames,MAX_ALLSECTIONS,m_strFileName);
 
 //以下循环，截断到两个连续的0
 for(i=0;i<MAX_ALLSECTIONS;i++)
 {
  if (chSectionNames[i]==0)
   if (chSectionNames[i]==chSectionNames[i+1])
    break;
 }
 
 iMaxCount=i+1; //要多一个0号元素。即找出全部字符串的结束部分。
 arrSection.RemoveAll();//清空原数组
 
 for(i=0;i<iMaxCount;i++)
 {
  chSection[iPos++]=chSectionNames[i];
  if(chSectionNames[i]==0)
  {   
   arrSection.Add(chSection);
   memset(chSection,0,MAX_SECTION);
   iPos=0;
  }
 
 }
 
 return (int)arrSection.GetSize();
}
 
int CIni::GetKeyValues(CStringArray& arrKey,CStringArray& arrValue, LPCTSTR lpSection)
{
 /*
 本函数基础：
 GetPrivateProfileSection- 从 ini 文件中获得一个Section的全部键名及值名
 如果ini中有一个段，其下有 "段1=值1" "段2=值2"，则返回的是 '段1=值1',0,'段2=值2',0,0 ，当你不知道  
 获得一个段中的所有键及值可以用这个。 
 */
 int i;  
 int iPos=0;
 CString strKeyValue;
 int iMaxCount;
 TCHAR chKeyNames[MAX_ALLKEYS]={0}; //总的提出来的字符串
 TCHAR chKey[MAX_KEY]={0}; //提出来的一个键名
 
 GetPrivateProfileSection(lpSection,chKeyNames,MAX_ALLKEYS,m_strFileName);
 
 for(i=0;i<MAX_ALLKEYS;i++)
 {
  if (chKeyNames[i]==0)
   if (chKeyNames[i]==chKeyNames[i+1])
    break;
 }
 
 iMaxCount=i+1; //要多一个0号元素。即找出全部字符串的结束部分。
 arrKey.RemoveAll();//清空原数组
 arrValue.RemoveAll();
 
 for(i=0;i<iMaxCount;i++)
 {
  chKey[iPos++]=chKeyNames[i];
  if(chKeyNames[i]==0)
  {
   strKeyValue=chKey;
   arrKey.Add(strKeyValue.Left(strKeyValue.Find("=")));
   arrValue.Add(strKeyValue.Mid(strKeyValue.Find("=")+1));
   memset(chKey,0,MAX_KEY);
   iPos=0;
  }
 
 }
 
 return (int)arrKey.GetSize();
}
 
BOOL CIni::DelAllSections()
{
 int nSection;
 CStringArray arrSection;
 nSection=GetSections(arrSection);
 for(int i=0;i<nSection;i++)
 {
  if(DelSection(arrSection[i]))
   return GetLastError();
 }
 return FALSE;
}

/*
使用方法：
CIni ini("c:\\a.ini");
int n;

/*获得值
TRACE("%s",ini.GetValue("段1","键1"));
 */

/*添加值
ini.SetValue("自定义段","键1","值");
ini.SetValue("自定义段2","键1","值",false);
 */

/*枚举全部段名
CStringArray arrSection;
n=ini.GetSections(arrSection);
for(int i=0;i<n;i++)
TRACE("%s\n",arrSection[i]);
 */

/*枚举全部键名及值
CStringArray arrKey,arrValue;
n=ini.GetKeyValues(arrKey,arrValue,"段1");
for(int i=0;i<n;i++)
TRACE("键：%s\n值：%s\n",arrKey[i],arrValue[i]);
*/

/*删除键值
ini.DelKey("段1","键1");
*/

/*删除段
ini.DelSection("段1");
*/

/*删除全部
ini.DelAllSections();
*/
*/
