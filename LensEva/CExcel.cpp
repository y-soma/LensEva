#include "CExcel.h"


/*=========================================================================
Constructor / Destructor
=========================================================================*/
CCExcel::CCExcel(void)
{
	// COM初期化
	CoInitialize(NULL);
	// Excel起動
	pXL.CreateInstance(L"Excel.Application");
	pBooks = pXL->Workbooks;

	//メンバ初期化 
	efinfo.clear();
}

CCExcel::~CCExcel(void)
{
	efinfo.clear();
	DestroyEXLPtr();

	// Free
	CoUninitialize();
}


/*=========================================================================
Public Member Function
=========================================================================*/


/**************************************************************************
	[Function]
		BOOL FileOpen()
	[Details]
		Excel形式のファイルを開く
	[Argument1] path
		ファイルパスへのポインタ
	[Argument2] disp
		表示の有効:1 / 無効:0
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::FileOpen(const TCHAR* const path, const int disp)
{
	if(!pXL){
		DestroyEXLPtr();
		pXL.CreateInstance(L"Excel.Application");
		pBooks = pXL->Workbooks;
	}
	
	if(!pBooks)
		pBooks = pXL->Workbooks;

	if(!disp)
		pXL->Visible = FALSE;
	else	
		pXL->Visible = TRUE;

    try
    {
		if( !(pBook = pBooks->Open(path)) )
			return FALSE;
		pSheet = pXL->ActiveSheet;
    }
    catch(...)
	{
		return FALSE;
	}

	return TRUE;
}

/**************************************************************************
	[Function]
		BOOL eFileOpen()
	[Details]
		Excel形式のファイルを開き、更にファイル情報をセットする
	[Argument1] path
		ファイルパスへのポインタ
	[Argument2] disp
		表示の有効:1 / 無効:0
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::eFileOpen(const TCHAR* const path, const int disp)
{
	if(!FileOpen(path,disp))
		return FALSE;

	{ // ファイル情報セット
		eFileInfo opfile;
		opfile.fpinfo.full = path;
		opfile.fpinfo.file = chl::GetFileName(path);
		opfile.fpinfo.dir = chl::GetDirName(path);
		opfile.pBook = pBook;
		SetSheetName(opfile.sheetinfo);
		efinfo.push_back(opfile);
		nowfinfo = opfile;
	}

	return TRUE;
}


/**************************************************************************
	[Function]
		BOOL eFileOpen()
	[Details]
		Excel形式のファイルを上書き保存する
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::FileSave(void)
{
	pBook->Save();
	return TRUE;
}

/**************************************************************************
	[Function]
		BOOL eFileOpen()
	[Details]
		Excel形式のファイルを名前をつけて保存する
	[Argument1] path
		WorkBookパス
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::FileSaveAs(const TCHAR* const path)
{
	if(!path || !pBook)
		return FALSE;
	//ATLASSERT(pBook != NULL);
	/* SaveAs Argument Details */
	// _WorkBook::SaveAs(Filename,FileFormat,Password,WriteResPassword,ReadOnlyRecommended,CreateBackup,AccessMode,ConflictResolution,AddToMru,TextCodepage,TextVisualLayout)
	pBook->SaveAs(
		//_variant_t(L"C:\\Users\\Yoshinori-Soma\\Desktop\\test00\\test01\\Test01.xls"),
		_variant_t(path),
		//_variant_t(path),
		static_cast<long>(Excel::xlWorkbookNormal),
		vtMissing,
		vtMissing,
		vtMissing,
		vtMissing,
		Excel::xlExclusive,
		static_cast<long>(Excel::xlLocalSessionChanges),
		false,
		vtMissing,
		vtMissing
	);

	return TRUE;
}


/**************************************************************************
	[Function]
		const TSTR GetCellsValue()
	[Details]
		指定座標のセル値を文字列で取得する
	[Argument1] line
		行位置への参照
	[Argument1] length
		列位置への参照
	[Return]
		Complete : セルの値		Error : NULL
**************************************************************************/
TSTR CCExcel::GetCellsValue(const ULONG& line, const ULONG& length)
{
	TSTR dst = _T("");
	
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return dst;

	// 値そのものを取得
	const TSTR str = (const _bstr_t)pRange->Value;

	{// セルが時間形式だった場合の調整
		const VARIANT format = pRange->GetNumberFormatLocal();
		const TSTR formatstr = format.bstrVal;
		if(formatstr == _T("h:mm") || formatstr == _T("h:mm:ss")){
			TCHAR szBuf[MAX_PATH];
			SYSTEMTIME systemTime;
			const VARIANT var = pRange->Value;
			VariantTimeToSystemTime(var.date, &systemTime);
			wsprintf(szBuf, TEXT("%d:%d"), systemTime.wHour, systemTime.wMinute);
			//dst.Empty();
			dst = szBuf;
			
			// 時間処理を入れること
			//dst = CTimeProc::HourMinCnvStr(CTimeProc::HourMinCnvLng(dst));
		}
		else{
			dst = str;
		}
	}
	
	return dst;
}


/**************************************************************************
	[Function]
		GetCellsColor()
	[Details]
		指定座標のセルの背景色をIDで取得する
	[Argument1] line
		行位置への参照
	[Argument1] length
		列位置への参照
	[Return]
		Complete : 色のID		Error : 
**************************************************************************/
long CCExcel::GetCellsColor(const ULONG& line, const ULONG& length)
{
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return NULL;

	const _variant_t vColorIdx = pRange->Interior->ColorIndex;
	const long ret = vColorIdx.lVal;
	
	return ret;
}

/**************************************************************************
	[Function]
		GetCellsColor()
	[Details]
		指定座標のセルのフォント色をIDで取得する
	[Argument1] line
		行位置への参照
	[Argument1] length
		列位置への参照
	[Return]
		Complete : 色のID		Error : 
**************************************************************************/
long CCExcel::GetCellsFontColor(const ULONG& line, const ULONG& length)
{
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return NULL;
	
	const _variant_t vColorIdx = pRange->Font->ColorIndex;
	const long ret = vColorIdx.lVal;
	//const _variant_t vColorIdx = (short)color;
	//if(pRange->Font->ColorIndex == (_variant_t)vColorIdx)
	
	return ret;
}


/**************************************************************************
	[Function]
		ShapesRotation()
	[Details]
		図形を回転させる
	[Argument1] rad
		回転角度
**************************************************************************/
void CCExcel::ShapesRotation(const float& rad)
{
	if(!pXL || rad < 0)
		return;
	Excel::ShapePtr pShape = NULL;
	{// ptr get
		const long scnt = pSheet->Shapes->GetCount();
		if(scnt < 1)
			return;
		pShape = pSheet->Shapes->Item(scnt);
	}
	if(!pShape)
		return;
	pShape->PutRotation(rad);
}


/**************************************************************************
	[Function]
		EditRectanglesCaption()
	[Details]
		四角形シェイプのテキスト編集
	[Argument1] str
		文字列
	[Remarks]
		注) シート上全ての四角形に対して適用する可能性があるので注意
**************************************************************************/
void CCExcel::EditRectanglesCaption(const TCHAR* const str)
{
	if(!pXL || !str)
		return;
	// ptr
	Excel::RectanglesPtr pRectangles = pSheet->Rectangles();
	if(!pRectangles)
		return;
	pRectangles->PutCaption(str);
}


/**************************************************************************
	[Function]
		CalcShapePos()
	[Details]
		シート内のシェイプの座標を求める
	[Argument1] x
		x座標算出結果
	[Argument2] y
		y座標算出結果
**************************************************************************/
void CCExcel::CalcShapePos(float& x, float& y)
{
	if(!pXL)
		return;
	Excel::ShapePtr pShape = NULL;
	{// ptr get
		const long scnt = pSheet->Shapes->GetCount();
		if(scnt < 1)
			return;
		pShape = pSheet->Shapes->Item(scnt);
	}
	if(!pShape)
		return;
	
	x = pShape->GetTop();
	y = pShape->GetLeft();
}


/**************************************************************************
	[Function]
		BOOL SetCellsValue()
	[Details]
		指定座標のセルの値を変更する
	[Argument1] line
		行位置への参照
	[Argument2] length
		列位置への参照
	[Argument3] setstr
		変更値
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::SetCellsValue(const ULONG &line, const ULONG& length, const TCHAR* const setstr)
{
	if(!pXL)
		return FALSE;
	RangePtr pCells = pSheet->Cells->Item[line][length];
	//Excel2002以上の場合はValue2
	pCells->Value2 = setstr;
	
	return TRUE;
}


/**************************************************************************
	[Function]
		MacroExecution()
	[Details]
		マクロを実行する
	[Argument1] MacroName
		マクロ名
	[Argument2] Argument=NULL
		マクロへ渡す引数
	[Return]
		Complete : TRUE		Error : FALSE
	[Remarks]
		現在開いているBookに対して有効です
		Excelマクロ側へ渡せる引数は1個まで対応  指定しない場合は引数無しと判断する
**************************************************************************/
BOOL CCExcel::MacroExecution(const TCHAR* const MacroName, const TCHAR* const Argument)
{
	BOOL ret = FALSE;
	
	TSTR RunName = _T("");
	{// マクロ関数名取得
		if(nowfinfo.fpinfo.file != _T(""))
			RunName = nowfinfo.fpinfo.file + _T('!') + MacroName;
		else
			RunName = MacroName;
	}
	
	HRESULT m_hr = E_FAIL;
	{// マクロ実行
		if(!Argument)
			m_hr = pXL->Run(RunName.c_str());
		else
			m_hr = pXL->Run(RunName.c_str(),Argument);
	}

	switch(m_hr)
	{// 処理結果
		case S_OK:
		case S_FALSE:
			ret = TRUE;
			break;
		case E_FAIL:
		default:
			ret = FALSE;
	}

	return ret;
}


/**************************************************************************
	[Function]
		BOOL CellsFtColorChange()
	[Details]
		指定座標のセル値のフォント色を変更
	[Argument1] line
		行位置への参照
	[Argument2] length
		列位置への参照
	[Argument3] color
		変更色
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::CellsFtColorChange(const ULONG& line, const ULONG& length, const int& color)
{
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return FALSE;
	
	//Cell値のフォント色変更
	const _variant_t vColorIdx = (short)color;
	pRange->Font->ColorIndex = vColorIdx;

	//FontPtr pFont = pRange->Font;
	//pFont->ColorIndex = color;
	
	return TRUE;
}


/**************************************************************************
	[Function]
		BOOL CellsFtColorChange()
	[Details]
		指定座標のセルのフォントサイズを変更
	[Argument1] line
		行位置への参照
	[Argument2] length
		列位置への参照
	[Argument3] size
		変更サイズ
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::CellsFtSizeChange(const ULONG& line, const ULONG& length, const int& size)
{
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return FALSE;

	const _variant_t vSize = (short)size;
	//pRange->Font->Size = vSize;
	pRange->Font->PutSize(vSize);

	return TRUE;
}


/**************************************************************************
	[Function]
		BOOL CellsFtColorChange()
	[Details]
		指定座標のセルのパターンを変更
	[Argument1] line
		行位置への参照
	[Argument2] length
		列位置への参照
	[Argument3] pattern
		指定パターン
		( 例> xlSolid: 白  xlGray16: 12.5%網掛け  xlGray25: 25%網掛け )
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::CellsPatternChange(const ULONG& line, const ULONG& length, const long& pattern)
{
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return FALSE;
	
	try
    {   
		if(pattern == 777){
			CellsBkColorChange(line,length);
			return TRUE;
		}
		//Cellのパターンを指定(Ver2007以外)
		const _variant_t vPattern = pattern;
		pRange->Interior->PutPattern(vPattern);
		
    }
    catch(...)
	{
		return FALSE;
	}

	return TRUE;
}


/**************************************************************************
	[Function]
		BOOL CellsBkColorChange()
	[Details]
		指定座標のセルを指定色で塗りつぶす
	[Argument1] line
		行位置への参照
	[Argument1] length
		列位置への参照
	[Argument3] color
		変更色
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::CellsBkColorChange(const ULONG& line, const ULONG& length, const int& color)
{
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return FALSE;
	const _variant_t vColorIdx = (short)color;
	pRange->Interior->ColorIndex = vColorIdx;

	return TRUE;
}


/**************************************************************************
	[Function]
		ActvSheetClear()
	[Details]
		アクティブシートを削除する
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::ActvSheetClear(void)
{
	if(!pBook || !pSheet)
		return FALSE;

	pBook->Application->PutDisplayAlerts(FALSE);
	pSheet->Delete();

	return TRUE;
}


/**************************************************************************
	[Function]
		ActvSheetNameChange()
	[Details]
		アクティブシート名の変更
	[Argument1] NewName
		変更名
**************************************************************************/
void CCExcel::ActvSheetNameChange(const TCHAR* const NewName)
{
	if(!pSheet)
		return;
	pSheet->PutName((_bstr_t)NewName);
}


/**************************************************************************
	[Function]
		BOOL ActvSheetChange()
	[Details]
		アクティブシートの変更
	[Argument1] StNo
		アクティブにしたいシート番号
		(開いた時点で先頭になっているシートがSheet1(1)になります)
	[Return]
		Complete : TRUE		Error : FALSE
	[Remarks]
		0を渡すとアクセスエラーになります
**************************************************************************/
BOOL CCExcel::ActvSheetChange(const long& StNo/*=1*/)
{
	const _WorksheetPtr CmpPtr = pSheet;
	if(!StNo)
		return FALSE;
	pSheet = pBook->Worksheets->Item[_variant_t((short)StNo)];

	if(!pSheet || pSheet==CmpPtr)
		return FALSE;
	return TRUE;
}

/**************************************************************************
	[Function]
		GetSheetCount()
	[Details]
		Excelファイルのシート数を取得
	[Return]
		Complete : 1以上のシート数	Error : 0以下の整数	
**************************************************************************/
long CCExcel::GetSheetCount(void)
{
	long ret = 0;
	
	if(!pBook || !pSheet)
		return ret;
	ret = pBook->Worksheets->GetCount();
	//ret = pSheet->GetIndex();

	return ret;
}


/**************************************************************************
	[Function]
		GetActvSheetIndex()
	[Details]
		アクティブシートのシート番号を取得
	[Return]
		Complete : 1以上のシート番号	Error : 0以下の整数	
**************************************************************************/
long CCExcel::GetActvSheetIndex(void)
{
	long ret = 0;
	
	if(!pBook || !pSheet)
		return ret;
	//ret = pBook->Worksheets->GetCount();
	ret = pSheet->GetIndex();

	return ret;
}



/**************************************************************************
	[Function]
		TSTR GetActvBookPath()
	[Details]
		現在操作しているブックのパスを取得する
	[Return]
		Complete : ブックのパス名		Error : Empty
**************************************************************************/
TSTR CCExcel::GetActvBookPath(void)
{
	TSTR ret = _T("");
	if(!pBook)
		return ret;
	for(UINT i=0; i < efinfo.size(); i++) {
		if(pBook == efinfo[i].pBook) {
			ret = efinfo[i].fpinfo.full;
			break;
		}
	}
	return ret;
}


/**************************************************************************
	[Function]
		BOOL SetActvBook()
	[Details]
		操作ブックを切り替える
	[Remarks]
		一度読んだ(閉じていない)ブックを指定しないと失敗する
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::SetActvBook(const TCHAR* const path)
{
	const TSTR pathstr = path;
	BOOL ret = FALSE;
	if(!pBook)
		return ret;
	for(UINT i=0; i < efinfo.size(); i++) {
		if(pathstr == efinfo[i].fpinfo.full) {
			pBook = efinfo[i].pBook;
			ret = TRUE;
			break;
		}
	}
	return ret;
}


/**************************************************************************
	[Function]
		BOOL SetCellsValue()
	[Details]
		指定座標のセル値,パターン等の全てをクリア
	[Argument1] line
		行位置への参照
	[Argument1] length
		列位置への参照
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
void CCExcel::CellsClear(const ULONG& line, const ULONG& length)
{
	if(!pXL)
		return;
	RangePtr pCells = pSheet->Cells->Item[line][length];
	pCells->Value2 = _T("");

	// フォントデフォルト
	CellsFtColorChange(line, length);
	// 背景色デフォルト
	CellsBkColorChange(line, length);
	// パターンデフォルト
	CellsPatternChange(line, length);
}


/**************************************************************************
	[Function]
		void FileClose()
	[Details]
		現在開いているファイルを閉じる
	[Remarks]
		全てのExcel操作が完了したときのみ使用すること
**************************************************************************/
void CCExcel::FileClose(void)
{
	eInfoDelete();
	eFileClose();
}


/**************************************************************************
	[Function]
		GetOSNameAndVersion
	[Details]
		OSの名前とバージョンを取得する(Excel仕様)
	[Return]
		取得したバージョン情報
**************************************************************************/

TSTR CCExcel::GetOSNameAndVersion(void) 
{
	const _bstr_t strOS = pXL->OperatingSystem;
	return static_cast<TSTR>(strOS);
}




/*=========================================================================
Private Member Function
=========================================================================*/

/*-------------------------------------------------------------------------
	[Function]
		void SetActvSheet()
	[Details]
		アクティブシートを変更する
	[Argument1] i
		アクティブにしたいシート№
	[Return]
		このクラス以外のアクセス禁止
-------------------------------------------------------------------------*/
void CCExcel::SetActvSheet(ULONG ShtNo)
{		
	if(!pXL || !pBook || !ShtNo)
		return;
	
	//アクティブ・シートの変更
	const _variant_t data = (short)ShtNo;
	static const SheetsPtr pSheets = pBook->Worksheets->Item[data];
	pSheets->Select(data);
	
}


/*-------------------------------------------------------------------------
	[Function]
		void SetSheetName()
	[Details]
		現在操作しているブックのシート情報をセットする
	[Argument1] lifo
		シート情報構造体の参照
	[Return]
		このクラス以外のアクセス禁止
-------------------------------------------------------------------------*/
void CCExcel::SetSheetName(eSheetInfo& info)
{		
	if(!pXL || !pBook)
		return;
	
	// 現在開いているシートを記憶する
	const _WorksheetPtr defptr = pSheet;
	
	// シート数の取得
	info.cnt = pXL->Worksheets->Count;
	// 存在するシート名を全てセット
	for(long i=1; i<info.cnt+1; i++) {
		_variant_t data = (short)i;
		//アクティブ・シートポインタの取得
		ActvSheetChange(i);
		info.pSheet.push_back(pSheet);
		//シート名を取得
		info.name.push_back(static_cast<TCHAR*>(pSheet->Name));
	}

	pSheet = defptr;
}


/*-------------------------------------------------------------------------
	[Function]
		BOOL eInfoDelete()
	[Details]
		現在開いているファイルの情報を削除する
	[Remarks]
		他の場所から単独での使用を禁止するためプライベート
	[Return]
		Complete  : TRUE   (削除成功)
		Error     : FALSE  (削除失敗もしくは対象が見つからない)
-------------------------------------------------------------------------*/
BOOL CCExcel::eInfoDelete(void)
{
	BOOL ret = FALSE;
	// 開いているファイルの情報を削除する
	const EINFOITR strtptr = efinfo.begin();
	const EINFOITR endptr = efinfo.end();
	for(EINFOITR itr=strtptr; itr!=endptr; itr++) {
		const eFileInfo& einfo = *itr;
		if(pBook == einfo.pBook) {
			efinfo.erase(itr);
			ret = TRUE;
			break;
		}
	}

	return ret;
}


/*-------------------------------------------------------------------------
	[Function]
		void eFileClose()
	[Details]
		現在開いているファイルを閉じる
	[Remarks]
		現在操作中のBookとSheetに対して有効
-------------------------------------------------------------------------*/
void CCExcel::eFileClose(void)
{
	// Excel操作主幹メンバの解放
	if(pBooks)
		pBooks.Release();
	
	// 操作Book, Sheet を一時解放
	if(pSheet)
		pSheet.Release();
	if(pBook)
		pBook.Release();
}


/*-------------------------------------------------------------------------
	[Function]
		void DestroyEXLPtr()
	[Details]
		Excel操作ポインタ全てを破棄
	[Remarks]
		デストラクタ専用
-------------------------------------------------------------------------*/
void CCExcel::DestroyEXLPtr(void)
{
	eFileClose();
	// pXL Free
	if(pXL){
		pXL->DisplayAlerts = FALSE;
		pXL->Quit();
		pXL.Release();
	}
}




/*=========================================================================
Static Member Function
=========================================================================*/



