//#pragma once

#ifndef CCEXCEL_INCLUDE
#define CCEXCEL_INCLUDE


/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
 Excelタイプライブラリ読み込み
   ビルドするPC環境に応じてインポート場所を変更
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*

//#include "stdafx.h"

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\mso.dll" no_namespace \
rename("DocumentProperties","XLDocumentProperties") 

#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB" no_namespace

#import "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" \
rename("DialogBox","XLDialogBox"),rename("CopyFile","XLCopyFile"), \
rename("RGB","XLRGB"),rename("ReplaceText","XLReplaceText") \
exclude("IFont", "IPicture") ,no_dual_interfaces
*/


#if 1
/**************************************************************************
Define
**************************************************************************/
// OS Bit
#define WIN64 0
// Excel Version
#define _EXCEL2000 0
#define _EXCEL2002 0
#define _EXCEL2003 1
#define _EXCEL2007 0
#define _EXCEL2010 0


#if WIN64
	#define _MSOVBE6EXT_PATH "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB"
#else /* WIN64 */
	#define _MSOVBE6EXT_PATH "C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
#endif /* WIN64 */

#if _EXCEL2000
	#if WIN64
		#define _MSODLL_PATH  "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files (x86)\Microsoft Office\Office\EXCEL.EXE"
	#else /* WIN64 */
		#define _MSODLL_PATH  "C:\Program Files\Common Files\Microsoft Shared\Office\MSO.DLL"	
		#define _MSEXCEL_PATH "C:\Program Files\Microsoft Office\Office\EXCEL.EXE"
	#endif /* WIN64 */
#endif
#if _EXCEL2002
	#if WIN64
		#define _MSODLL_PATH  "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE10\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files (x86)\Microsoft Office\Office10\EXCEL.EXE"
	#else /* WIN64 */
		#define _MSODLL_PATH  "C:\Program Files\Common Files\Microsoft Shared\Office10\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files\Microsoft Office\Office10\EXCEL.EXE"
	#endif /* WIN64 */
#endif
#if _EXCEL2003
	#if WIN64
		#define _MSODLL_PATH  "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE11\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files (x86)\Microsoft Office\Office11\EXCEL.EXE"
	#else /* WIN64 */
		#define _MSODLL_PATH  "C:\Program Files\Common Files\Microsoft Shared\Office11\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files\Microsoft Office\Office11\EXCEL.EXE"
	#endif /* WIN64 */	
#endif
#if _EXCEL2007
	#if WIN64
		#define _MSODLL_PATH  "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE12\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE"
	#else /* WIN64 */
		#define _MSODLL_PATH  "C:\Program Files\Common Files\Microsoft Shared\Office12\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files\Microsoft Office\Office12\EXCEL.EXE"
	#endif /* WIN64 */
#endif
#if _EXCEL2010
	#if WIN64
		#define _MSODLL_PATH  "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE"
	#else /* WIN64 */
		#define _MSODLL_PATH  "C:\Program Files\Common Files\Microsoft Shared\Office14\MSO.DLL"
		#define _MSEXCEL_PATH "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE"
	#endif /* WIN64 */
#endif

/**************************************************************************
Import
**************************************************************************/
// GetFirstChildの実引数不一致警告を無効にする(importの前でやる)
#pragma warning(disable:4003)
#pragma warning(disable:4278)
#pragma warning(disable:4192)
//#pragma warning( disable : 4786 )
#import _MSODLL_PATH no_namespace rename("DocumentProperties", "DocumentPropertiesXL")   
#import _MSOVBE6EXT_PATH no_namespace
//#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7\VBE7.DLL" no_namespace
#import _MSEXCEL_PATH rename("DialogBox", "DialogBoxXL") rename("RGB", "RBGXL") rename("DocumentProperties", "DocumentPropertiesXL") no_dual_interfaces

#endif


// CCExcel : 定義
// 

#include <vector>
#include "GFunc.h"

using namespace Excel;

// Typedef
#ifndef TSTRDEFINE
	#define TSTRDEFINE
	typedef std::basic_string<TCHAR> TSTR;
#endif


/**************************************************************************
	[Class]
		CCExcel
	[Details]
		Excel操作サポートクラス
	[Remarks]
		Win32SDK,MFC両用
**************************************************************************/
class CCExcel
{
public:
	CCExcel(void);
	~CCExcel(void);


private:
	/**************************************************************************
		[Structure]
			eSheetInfo
		[Details]
			ブック1つあたりのシート情報を管理する構造体
	**************************************************************************/
	typedef struct EXCEL_SHEET_INFO {
		// WorkSheetポインタ
		std::vector<_WorksheetPtr> pSheet;
		// シート名
		std::vector<TSTR> name;
		// シート数
		long cnt;
	} eSheetInfo;
	
	/**************************************************************************
		[Structure]
			eFileInfo
		[Details]
			開いたExcelファイルを管理する構造体
	**************************************************************************/
	typedef struct EXCEL_FILE_INFO {
		// ファイルパス情報
		chl::FilePInfo fpinfo;
		// WorkBookポインタ
		_WorkbookPtr pBook;
		// Sheet情報
		eSheetInfo sheetinfo;
	} eFileInfo;
	
	// ExcelControlPtr
	_ApplicationPtr pXL;
	WorkbooksPtr pBooks;
	
	// 操作ファイルのポインタ
	_WorkbookPtr pBook;
	_WorksheetPtr pSheet;
	
	// 開いたファイル情報管理
	std::vector<eFileInfo> efinfo;
	// 現在開いているもの
	eFileInfo nowfinfo;


public:
	/* Excelファイルオープン */
	BOOL FileOpen(const TCHAR* const, const int = 0);
	/* Excelファイルオープン & 情報セット */
	BOOL eFileOpen(const TCHAR* const, const int = 0);
	/* Excelファイルを上書き保存 */
	BOOL FileSave(void);
	/* Excelファイルを名前をつけて保存 */
	BOOL FileSaveAs(const TCHAR* const);
	/* Cellの値を取得 */
	TSTR GetCellsValue(const ULONG&, const ULONG&);
	/* Cellに値を書き込む */
	BOOL SetCellsValue(const ULONG&, const ULONG&, const TCHAR* const = _T(""));
	/* Cell値のフォント色変更 */
	BOOL CellsFtColorChange(const ULONG& line, const ULONG& length, const int& color = 0);
	/* Cell値のフォントサイズ変更 */
	BOOL CellsFtSizeChange(const ULONG& line, const ULONG& length, const int& size = 10);
	/* Cellの背景色変更 */
	BOOL CellsBkColorChange(const ULONG& line, const ULONG& length, const int& color = Excel::xlNone);
	/* Cellのパターン変更 */
	BOOL CellsPatternChange(const ULONG& line, const ULONG& length, const long& pattern = 777);
	/* Cellの背景色取得 */
	long GetCellsColor(const ULONG& line, const ULONG& length);
	/* Cellのフォント色取得 */
	long GetCellsFontColor(const ULONG& line, const ULONG& length);
	/* 図形を回転させる */
	void ShapesRotation(const float& rad);
	/* 四角形シェイプのテキスト編集 */
	void EditRectanglesCaption(const TCHAR* const str);
	/* シート内のシェイプの座標を求める */
	void CalcShapePos(float& x, float& y);
	
	/* アクティブシートを削除する */
	BOOL ActvSheetClear(void);
	/* アクティブシート名の変更 */
	void ActvSheetNameChange(const TCHAR* const NewName);
	/* Excelファイルのシート数を取得 */
	long GetSheetCount(void);
	/* アクティブシートのシート番号を取得 */
	long GetActvSheetIndex(void);
	/* 操作シートの変更 */
	BOOL ActvSheetChange(const long& StNo = 1);
	/* 現在操作しているファイルのパスを取得 */
	TSTR GetActvBookPath(void);
	/* 操作ファイルを変更 */
	BOOL SetActvBook(const TCHAR* const);
	
	/* マクロ実行 */
	BOOL MacroExecution(const TCHAR* const MacroName, const TCHAR* const Argument = NULL);
	
	/* 指定場所のセルをクリア */
	void CellsClear(const ULONG&, const ULONG&);
	/* 現在操作しているファイルを閉じる */
	void FileClose(void);

	/* OSの名前とバージョンを取得する */
	TSTR GetOSNameAndVersion(void);


private:
	// typedef
	typedef std::vector<eFileInfo>::iterator EINFOITR;
	
	/* 現在操作しているファイルを閉じる */
	void SetActvSheet(ULONG);
	/* 開いたファイルからシート情報をセットする */
	void SetSheetName(eSheetInfo&);
	/* 閉じたファイルの情報を削除する */
	BOOL eInfoDelete(void);
	/* ファイルを閉じる(単独) */
	void eFileClose(void);
	/* Excel操作ポインタ全てを破棄する */
	void DestroyEXLPtr(void);

}; // End Class

#endif /* CCEXCEL_INCLUDE */
