/* -+-+-+-+-+-+-+-+-+-+-+-+-+-+ 汎用関数郡定義ファイル -+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

#pragma once

#ifndef GFUNC_INCLUDE
#define GFUNC_INCLUDE

#include <string>
#include <tchar.h>
#include <vector>
#include <shlwapi.h>
#pragma comment(lib, "shlwapi.lib")

// 改行代替リテラル
#define CRLF_AGENCY_LITERAL	L'$'

#define NOT_FOUND _T("default")
// wcscpy warning disable
#pragma warning (disable: 4996)

// Typedef
#ifndef TSTRDEFINE
	#define TSTRDEFINE
	typedef std::basic_string<TCHAR> TSTR;
#endif

namespace chl
{
	// Function Prototype
	std::vector<TSTR> Split(const TCHAR* const str, const TCHAR* const delim);	
	TSTR GetFileName(const TCHAR* const fullPath);	
	TSTR GetDirName(const TCHAR* const fullPath);
	TSTR GetAppMeDir(const int& el = 0);
	TSTR LngToStr(const long& val);
	TSTR DblToStr(const double& val);
	double StrToDbl(const TSTR& str);
	long StrToLng(const TSTR& str);
	TSTR Dec2HexStr(const ULONG& val);
	TSTR Dec2HexStr(const TSTR& val);
	TSTR GetNumInStr(const TSTR& src);
	BOOL WriteInitFile(const TCHAR* const IniName, const TCHAR* const section, const TCHAR* const param, const TCHAR* const val, const TCHAR* const dir = NULL);
	TSTR GetInitParam(const TCHAR* const IniName, const TCHAR* const section, const TCHAR* const param, const TCHAR* const dir = NULL);
	template<typename TOBJ> std::vector<TOBJ> GetArryEqualVal(const std::vector<TOBJ>& src);
	TSTR CrLfConvGUIEdit(const TSTR& src);
	TSTR CrLfReturnGUIEdit(const TSTR& src);
	TSTR GetPCAndUserName(void);
	BOOL PutLogFile(const TCHAR* const path, const TCHAR* const header = NULL);	
	BOOL PathIsDirectoryEX(const TCHAR* const path);
	BOOL PathFileExistsEX(const TCHAR* const path);	
	TSTR GetMakeFilePath(const TCHAR* const path, const TCHAR* const opfname = _T("log"));	
	TSTR GetStrToUnicode(const TCHAR* const src, const ULONG& size);
	BOOL PutStrToCode(const TCHAR* const src, const ULONG& size);	
	TSTR PathEscSeqConv(const TCHAR* const src, const TCHAR* const val = _T("/"));
	TSTR GetOSVersion(void);



	/*=========================================================================
	Structure
	=========================================================================*/

	/**************************************************************************
		[Structure]
			FileInfo
		[Details]
			開いたファイルのパスを管理する構造体
	**************************************************************************/
	typedef struct FILE_PATH_INFO{
		// ディレクトリ
		TSTR dir;
		// ファイル名
		TSTR file;
		// フルパス
		TSTR full;
		
		/* Constructor */
		FILE_PATH_INFO()
		{
			dir.empty();
			file.empty();
			full.empty();
		}
	} FilePInfo;

}; // End namespace


#endif /* GFUNC_INCLUDE */




