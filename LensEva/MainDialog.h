#pragma once

// CMainDialog : 定義

#include <windows.h>
#include <windowsx.h>
#include <stdio.h>
#include <string.h>
#include <tchar.h>
#include <time.h>
#include <algorithm>
#include <shlobj.h>
#include "resource.h"
#include "CExcel.h"
#include "LProc.h"

#include <commctrl.h>
#pragma comment(lib, "comctl32.lib")

// Define
#define THIS_VERSION	_T("1.0.0.2")
#define COPYRIGHT		_T("Copyright (C) 2011 OLYMPUS IMAGING CORP. All Rights Reserved.")
#define EVABODY_INI		_T("body.ini")
#define LISTBOTTOM	65


/* Gloval */

// 読み込んだファイル情報管理
std::vector<chl::FilePInfo> Fileinfo;
// 評価ボディ
std::vector<TSTR> Evabody;
// 出力先
TSTR Outdir;

// ファイルリストサブクラスプロパティ設定
const TCHAR Filelist_Propname[] = _T("FilelistProp");
// 出力先エディットプロパティ名
const TCHAR Outdir_Propname[] = _T("OutdirProp");



// disable
#pragma warning(disable:4311)
#pragma warning(disable:4312)
#pragma warning(disable:4244)


/**************************************************************************
	[Class]
		CMainDialog
	[Details]
		メインダイアログ
**************************************************************************/
class CMainDialog
{

/* Constructor / Destructor */
public:
	CMainDialog(void)
	{
		FieldVarInit();
	};
	~CMainDialog(void)
	{
		FieldVarInit();
	};

protected:
	HWND ghDlg;
	WNDPROC DefStaticProc;

private:
	HWND hList, hCombo, hOutdir;


/* Public Member Function */
public:
	/**************************************************************************
		[Function]
			DlgCreate
		[Details]
			ダイアログ作成
		[Argument1] hInstance
			インスタンスハンドル
		[Return]
			Falseを返します
	**************************************************************************/
	int DlgCreate(const HINSTANCE hInstance)
	{
		return (int)DialogBox(hInstance, MAKEINTRESOURCE(IDD_MAIN_DIALOG),(HWND)NULL,(DLGPROC)DlgProc);
	}


/* CALLBACK Function */
private:
	/*------------------------------------------------------------------------
		[Function]
			DialogProc
		[Details]
			ダイアログボックスのコールバック関数
		[Argument1] hDlg
			ダイアログのウインドウハンドル
		[Argument2] uMsg
			取得メッセージ
		[Argument3] wParam
			メッセージの最初のパラメータ
		[Argument4] lParam
			メッセージの2番目のパラメータ
		[Return]
			メッセージ処理の結果が返ります。
			戻り値の意味は、送信されたメッセージによって異なります。
			初期化時を除き、Falseが返ります
	------------------------------------------------------------------------*/
	static LRESULT CALLBACK DlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		static const TCHAR s_prop_name[] = _T("MainDlgProp");
		
		switch(uMsg)
		{
			case WM_INITDIALOG:
			{// 初期化時
				// 既存のコントロールからthisポインタ取得
				CMainDialog* md = reinterpret_cast<CMainDialog*>(GetWindowLong(hDlg,GWL_EXSTYLE));
				
				/* ※作成したウインドウから取得する場合はこちら
				::CREATESTRUCT* cs = reinterpret_cast<::CREATESTRUCT*>(lParam);
				if(!cs)
					return -1;
				CMainDialog* md = reinterpret_cast<CMainDialog*>(cs->lpCreateParams);
				*/
				
				if(!md)
					return -1;
				// プロパティにWindowクラスのポインタを設定する
				if( !::SetProp(hDlg, s_prop_name, reinterpret_cast<HANDLE>(md)) )
					return -1;

				// ウィンドウハンドルセット
				md->ghDlg = hDlg;
				// メッセージ処理関数コール
				return md->MsgProc(hDlg, uMsg, wParam, lParam);
			}
			case WM_CLOSE:
			{// 終了時
				CMainDialog* md = reinterpret_cast<CMainDialog*>(::GetProp(hDlg, s_prop_name));

				LRESULT ret = 0;
				// メッセージ処理関数コール
				if(md != NULL)
				{
					ret = md->MsgProc(hDlg, uMsg, wParam, lParam);
					md->ghDlg = NULL;
				}
				else
				{
					ret = ::DefWindowProc(hDlg, uMsg, wParam, lParam);
				}

				// 設定したプロパティのデータを削除する
				::RemoveProp(hDlg, s_prop_name);
				return ret;
			}
			default:
			{// その他
				CMainDialog* md = reinterpret_cast<CMainDialog*>(::GetProp(hDlg, s_prop_name));
				if(md != NULL)
					return md->MsgProc(hDlg, uMsg, wParam, lParam);
				else
					return ::DefWindowProc(hDlg, uMsg, wParam, lParam);
			}
		}
		return FALSE;
	}


/* Protected Member Function */
protected:
	/*------------------------------------------------------------------------
		[Function]
			DialogProc
		[Details]
			ダイアログボックスのコールバック関数
		[Argument1] hDlg
			ダイアログのウインドウハンドル
		[Argument2] uMsg
			取得メッセージ
		[Argument3] wParam
			メッセージの最初のパラメータ
		[Argument4] lParam
			メッセージの2番目のパラメータ
		[Return]
			メッセージ処理の結果が返ります。
			戻り値の意味は、送信されたメッセージによって異なります。
			初期化時を除き、Falseが返ります
	------------------------------------------------------------------------*/
	LRESULT CALLBACK MsgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		switch(uMsg)
		{
			case WM_INITDIALOG:
			{
				// リストコントロールをサブクラス化
				static CFilelistSubcls FileList(hDlg,uMsg,wParam,lParam,IDC_FILELIST);
				// 出力先エディットをサブクラス化
				static COutdirSubcls Outdir(hDlg,uMsg,wParam,lParam,IDC_EDIT_OUTDIR);
				InitCommonControls();
				return HANDLE_WM_INITDIALOG(hDlg, wParam, lParam, OnInitDialog);
			}
			case WM_COMMAND:
				return GetWMCOMMAND(wParam);
			//case WM_DROPFILES:	
				//メッセージをファイルリストサブクラス側へ譲渡
			case WM_TIMER:
				//return HANDLE_WM_TIMER(hDlg, wParam, lParam, OnTimer);
				break;
			case WM_CLOSE:
				EndDialog(hDlg, FALSE);
				return FALSE;
			default:
				break;
		}

		return FALSE;
	}

	
	/*------------------------------------------------------------------------
		[Function]
			FieldVarInit
		[Details]
			この領域内のグローバル変数を初期化する
	------------------------------------------------------------------------*/
	void FieldVarInit(void)
	{
		Fileinfo.clear();
		Evabody.clear();
		Outdir.empty();
	}
	
	
	/*------------------------------------------------------------------------
		[Function]
			MoveCenter
		[Details]
			ダイアログを画面の中心に移動させる : 親がないので何もしないと左上に配置される
		[Argument1] hwnd
			ウインドウハンドル
	------------------------------------------------------------------------*/
	void MoveCenter(HWND hDlg)
	{
		int w = 0, h = 0, wpos = 0, hpos = 0;
		{// ダイアログのサイズと移動位置を取得する
			{// ダイアログの縦横サイズ
				RECT rc;
				GetWindowRect(hDlg, &rc);
				w = rc.right - rc.left;
				h = rc.bottom - rc.top;
			}
			{// スクリーンサイズから計算
				const int wFull = GetSystemMetrics(SM_CXSCREEN);
				const int hFull = GetSystemMetrics(SM_CYSCREEN);
				wpos = static_cast<int>((wFull-w)*0.5);
				hpos = static_cast<int>((hFull-h)*0.5);
			}
		}
		MoveWindow(hDlg, wpos, hpos, w, h, FALSE);
	}

	
	/*------------------------------------------------------------------------
		[Function]
			center_window
		[Details]
			ダイアログを画面の中央に配置する(方法その2)
		[Argument1] hwnd
			ウインドウハンドル
	------------------------------------------------------------------------*/
	void center_window(const HWND hwnd)
	{
		RECT desktop;
		RECT rect;
		int width, height;
		int x, y;

		GetWindowRect(GetDesktopWindow(), &desktop);
		GetWindowRect(hwnd, &rect);

		width = rect.right - rect.left;
		height = rect.bottom - rect.top;

		x = (desktop.left + desktop.right) / 2 - width / 2;
		y = (desktop.top + desktop.bottom) / 2 - height / 2;

		if (x < desktop.left) {
			x = desktop.left;
		} else if (x + width > desktop.right) {
			x = desktop.right - width;
		}

		if (y < desktop.top) {
			y = desktop.top;
		} else if (y + height > desktop.bottom) {
			y = desktop.bottom - height;
		}

		SetWindowPos(hwnd, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
	}


	/*------------------------------------------------------------------------
		[Function]
			InitListView
		[Details]
			リストビュー初期化
		[Argument1] hDlg
			ウインドウハンドル
	------------------------------------------------------------------------*/
	void InitListView(const HWND hDlg)
	{
		// 既存のリストビューのハンドルを取得
		hList = GetDlgItem(hDlg, IDC_FILELIST);
		
		{// リストビューのスタイル指定
			DWORD dwStyle;
			dwStyle = ListView_GetExtendedListViewStyle(hList);
			dwStyle |= LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES;   // LVS_EX_HEADERDRAGDROP; ドラッグ＆ドロップ可能  LVS_EX_CHECKBOXES | チェックボックスをつける
			ListView_SetExtendedListViewStyle(hList, dwStyle);
		}

		LV_COLUMN lvcol;
		{
			lvcol.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
			lvcol.fmt = LVCFMT_LEFT;
		}

		TCHAR caption[][32] = {_T("FILE"), _T("DIR")};
		const UINT clmNum = sizeof caption /sizeof caption[0];
	    
		// 倍率指定
		const double mag[] = { 0.295,0.695 };
		
		if(sizeof(caption)/sizeof(caption[0]) != sizeof(mag)/sizeof(mag[0]))
			return;

		long width = 0;
		{// リストコントロールの幅サイズ取得
			RECT rect;
			::GetWindowRect(hList, &rect);
			width = rect.right - rect.left;
			//height = rect.bottom - rect.top;
		}
		for(int i = 0; i < sizeof(caption)/sizeof(caption[0]); i++)
		{
			// サブアイテム番号
			lvcol.iSubItem = i;
			// 見出しテキスト
			lvcol.pszText = caption[i];
			
			{// 横幅設定
				const double setsize = width * mag[i];
				lvcol.cx = static_cast<int>(setsize);
			}
			if(ListView_InsertColumn(hList,i,&lvcol) == -1)
				return;
		}
	}


	/*------------------------------------------------------------------------
		[Function]
			InitEvabodyComboBox
		[Details]
			コンボボックス初期化
		[Argument1] hDlg
			ウインドウハンドル
	------------------------------------------------------------------------*/
	void InitEvabodyComboBox(const HWND hDlg)
	{
		// 既存のコンボボックスのハンドルを取得
		hCombo = GetDlgItem(hDlg, IDC_COMBO_EVABODY);
		{// iniファイルから取得したボディ情報をコンボボックスへ追加
			const TSTR evabody = chl::GetInitParam(EVABODY_INI,_T("BODY"),_T("value"));
			if(evabody != NOT_FOUND)
			{
				const std::vector<TSTR> bodyname = chl::Split(evabody.data(),_T(","));
				for(UINT i=0; i<bodyname.size(); i++)
				{
					SendMessage(hCombo,CB_INSERTSTRING,i,(LPARAM)bodyname[i].data());
				}
				// 一番目のアイテムを選択
				SendMessage(hCombo, CB_SETCURSEL, 0, 0);
				Evabody = bodyname;
			}
		}
	}


	/*------------------------------------------------------------------------
		[Function]
			UpdateEvabodyComboBox
		[Details]
			コンボボックスを更新する
		[Argument1] hDlg
			ウインドウハンドル
	------------------------------------------------------------------------*/
	void UpdateEvabodyComboBox(void)
	{
		while(SendMessage(hCombo, CB_GETCOUNT, 0, 0) != 0)
		{// コンボボックスの中身を全て削除
			SendMessage(hCombo, CB_DELETESTRING, 0, 0);
		}

		for(UINT i=0; i<Evabody.size(); i++)
		{
			SendMessage(hCombo,CB_INSERTSTRING,i,(LPARAM)Evabody[i].data());
		}
		// 一番目のアイテムを選択
		SendMessage(hCombo, CB_SETCURSEL, 0, 0);
	}


	/*------------------------------------------------------------------------
		[Function]
			SelectedListDelete
		[Details]
			リストから選択しているファイルを1つ削除する
	------------------------------------------------------------------------*/
	void SelectedListDelete(const HWND hDlg)
	{
		std::vector<chl::FilePInfo>::iterator it = Fileinfo.begin();
		for(UINT i=0; i<Fileinfo.size(); i++)
		{
			const UINT uiState = ListView_GetItemState(hDlg, i, LVIS_SELECTED);
			if(uiState & LVIS_SELECTED){
				Fileinfo.erase(it);
				break;
			}
			++it;
		}

		// 更新
		UpdateFileList(hDlg);
	}


	/*------------------------------------------------------------------------
		[Function]
			OnInitDialog
		[Details]
			ダイアログ初期化
		[Return]
			FALSEを返す
	------------------------------------------------------------------------*/
	BOOL OnInitDialog(const HWND hDlg, const HWND hwndFocus, const LPARAM lParam)
	{
		// リストビュー初期化
		InitListView(hDlg);
		// コンボボックス初期化
		InitEvabodyComboBox(hDlg);
		// 出力先ディレクトリエディットのハンドル取得
		hOutdir = GetDlgItem(hDlg, IDC_EDIT_OUTDIR);
		// ウインドウを中央へ配置
		//center_window(hDlg);
		MoveCenter(hDlg);
		return FALSE;
	}


	/*------------------------------------------------------------------------
		[Function]
			ReadFileDlg
		[Details]
			ダイアログからファイル読み込み
		[Return]
			読み込みファイル情報
	------------------------------------------------------------------------*/
	std::vector<chl::FilePInfo> ReadFileDlg(void)
	{
		std::vector<chl::FilePInfo> dst;
		{
			static OPENFILENAME ofn;
			static TCHAR filename_full[MAX_PATH*0xFF];
			static TCHAR filename_n[MAX_PATH*0xFF];
			{// ダイアログ初期設定
				ZeroMemory(&ofn, sizeof(ofn));
				ZeroMemory(&filename_full, sizeof(filename_full));
				ZeroMemory(&filename_n, sizeof(filename_n));
				ofn.lStructSize = sizeof(ofn);
				ofn.hwndOwner = ghDlg;
				ofn.lpstrFile = filename_full;
				ofn.nMaxFile = sizeof(filename_full);
				ofn.lpstrFilter = _T("All files(*.*)\0*.*\0\0");
				ofn.lpstrTitle = _T("ファイルを開く");
				ofn.lpstrDefExt = _T("csv");
				ofn.lpstrFileTitle = filename_n;
				ofn.nMaxFileTitle = sizeof(filename_n);
				ofn.Flags = OFN_EXPLORER | OFN_ALLOWMULTISELECT;
			}

			if(!GetOpenFileName(&ofn))
				return dst;

			/* こうすると保存ダイアログになる
			ofn->Flags = OFN_OVERWRITEPROMPT;
			ofn->lpstrTitle = _T("名前を付けて保存");
			if(!GetSaveFileName(&ofn))
				return dst;
			*/

			// パスのみセット
			const TSTR path = filename_full;

			for(ULONG i=ofn.nFileOffset-1; i<ofn.nMaxFile; i++)
			{// 全てのファイル情報取得
				if(filename_full[i])
					continue;
				
				if(!filename_full[i+1])
				{// 1つ先の場所がNULLの場合、終了
					break;
				}
				else
				{// NULLじゃない場合、ファイル情報が続いている
					ULONG cp = i+1;
					TSTR fNameTemp = _T("");
					while(filename_full[cp] != NULL)
					{// 次のNULLポイントまでを1文字単位で結合させる
						fNameTemp += filename_full[cp];
						cp++;
					}
					{// データ整理して取得
						chl::FilePInfo buf;
						{
							buf.dir = path;
							buf.file = fNameTemp;
							buf.full = path + _T('\\') + fNameTemp;
						}
						dst.push_back(buf);
					}
				}
			}

			if(!dst.size())
			{// ファイルが1個だけだった場合
				chl::FilePInfo buf;
				{
					buf.dir = chl::GetDirName(path.data());
					buf.file = chl::GetFileName(path.data());
					buf.full = path;
				}
				dst.push_back(buf);
			}
		}
		SortFileData(dst);
		return dst;
	}


	/*------------------------------------------------------------------------
		[Function]
			ReadDirDlg
		[Details]
			ダイアログからフォルダ選択
		[Return]
			成功：選択したディレクトリ　　失敗：empty
	------------------------------------------------------------------------*/
	TSTR ReadDirDlg(void)
	{
		TSTR dst = _T("");
		{
			BROWSEINFO bInfo;
			TCHAR szDisplayName[MAX_PATH*0xFF];
			{// ダイアログ構造体の準備
				bInfo.hwndOwner             = ghDlg;          // ダイアログの親ウインドウのハンドル
				bInfo.pidlRoot              = NULL;                             // ルートフォルダを示すITEMIDLISTのポインタ (NULLの場合デスクトップフォルダが使われます）
				bInfo.pszDisplayName        = szDisplayName;                    // 選択されたフォルダ名を受け取るバッファのポインタ
				bInfo.lpszTitle             = _T("フォルダの選択");             // ツリービューの上部に表示される文字列 
				bInfo.ulFlags               = BIF_RETURNONLYFSDIRS;             // 表示されるフォルダの種類を示すフラグ
				bInfo.lpfn                  = NULL;                             // BrowseCallbackProc関数のポインタ
				bInfo.lParam                = (LPARAM)0;                        // コールバック関数に渡す値
			}

			// フォルダ選択ダイアログを表示
			LPITEMIDLIST pIDList = ::SHBrowseForFolder(&bInfo);
			if(pIDList)
			{
				if(::SHGetPathFromIDList(pIDList, szDisplayName))
					dst = szDisplayName;
				else
					MessageBox(ghDlg, _T("フォルダ選択に失敗しました"), _T("Error"), MB_OKCANCEL | MB_ICONERROR);
				// ポインタの解放
				::CoTaskMemFree(pIDList);
			}
		}
		return dst;
	}


	/*------------------------------------------------------------------------
		[Function]
			GetWindowTxtEX
		[Details]
			コントロールに表示されているテキストを取得する
		[Argument1] gWnd
			指定ウインドウハンドル
		[Return]
			表示されているテキストが文字列で返る
	------------------------------------------------------------------------*/
	TSTR GetWindowTxtEX(const HWND& gWnd)
	{
		TSTR dst = _T("");
		{// 取得
			TCHAR str[MAX_PATH*0xFF];
			GetWindowText(gWnd,str,MAX_PATH*0xFF);
			dst = str;
		}
		return dst;
	}


	/*------------------------------------------------------------------------
		[Function]
			Execution
		[Details]
			処理を実行する
	------------------------------------------------------------------------*/
	void Execution(void)
	{
		try
		{
			if(Fileinfo.size() < 1){
				MessageBox(ghDlg, _T("ファイルが選択されていません"), _T("NO FILES"), MB_OKCANCEL | MB_ICONERROR);
				return;
			}
			{// 実行
				CLProc Proc;
				Proc.Execution(Fileinfo,GetWindowTxtEX(hCombo),GetWindowTxtEX(hOutdir));
			}
		}
		catch(...)
		{
			MessageBox(ghDlg, _T("処理中に致命的なエラーが発生しました"), _T("Fatal Error"), MB_OKCANCEL | MB_ICONERROR);
		}
	}


	/*------------------------------------------------------------------------
		[Function]
			GetWMCOMMAND
		[Details]
			コマンド系メッセージ処理取得
		[Argument1] wParam
			コマンド系メッセージ
		[Return]
			メッセージ処理の結果が返ります。
			戻り値の意味は、送信されたメッセージによって異なります。
	------------------------------------------------------------------------*/
	LRESULT GetWMCOMMAND(const WPARAM wParam)
	{
		switch(LOWORD(wParam))
		{
			// キャンセルボタン
			case IDCANCEL:
				EndDialog(ghDlg, FALSE);
				return FALSE;
			// 実行ボタン
			case IDC_BUTTON_EXECUTION:
			{
				Execution();
				return FALSE;
			}
			// メニューからファイルを開く
			case ID_FILEOPEN:
				AddFileList(hList,ReadFileDlg());
				return FALSE;
			// 選択されているファイルを1つ削除
			case ID_FILEDELETE:
				SelectedListDelete(hList);
				return FALSE;
			// ファイルを全て削除
			case ID_ALLCLEAR:
				Fileinfo.clear();
				ListView_DeleteAllItems(hList);
				return FALSE;
			// 閉じる
			case ID_FILECLOSE:
				EndDialog(ghDlg, FALSE);
				return FALSE;
			// 出力先をダイアログで選択
			case IDC_OUTDIR_SELECT:
			{
				const TSTR outdir = ReadDirDlg();
				if(outdir != _T(""))
					SetWindowText(hOutdir,outdir.data());
				return FALSE;
			}
			// 出力先を削除
			case IDC_OUTDIR_CLEAR:
			{
				SetWindowText(hOutdir,_T(""));
				return FALSE;
			}
			// バージョン情報
			case ID_VERSION:
			{
				TSTR mes = _T("LensEva Version ");
				mes += THIS_VERSION;
				mes += _T("\n");
				mes += COPYRIGHT;
				MessageBox(ghDlg, mes.data(), _T("バージョン情報"), MB_OK | MB_ICONINFORMATION);
				return FALSE;
			}
			// コンボボックス関連
			case IDC_COMBO_EVABODY:
				// 選択などの特殊なイベントハンドラを得たい場合は HIWORD(wParam) を評価する
				if(HIWORD(wParam) == CBN_EDITCHANGE)
				{// 直接入力後のイベントハンドラ
					//hCombo = GetDlgItem(hDlg, IDC_COMBO_EVABODY);
					//hCombo = GetDlgItemText(hDlg, IDC_COMBO_EVABODY);
					
					//if(0){
					//	UpdateEvabodyComboBox();
					//}
				}
				return FALSE;
		}
		return FALSE;
	}


	
/* Static Member Function */
	/*------------------------------------------------------------------------
		[Function]
			ReadDropfile
		[Details]
			ドラッグ＆ドロップファイル読み込み
		[Argument1] hWnd
			ウインドウハンドル
		[Argument2] wParam
			パラメータ
		[Return]
			読み込みファイル情報
	------------------------------------------------------------------------*/
	static std::vector<chl::FilePInfo> ReadDropfile(const WPARAM wParam)
	{
		std::vector<chl::FilePInfo> dst;
		{
			TCHAR FileName[MAX_PATH * 0xFF] = _T("");
			HDROP hDrop = (HDROP)wParam;

			int ic = 0;
			{// ドロップファイル数を取得
				ic = DragQueryFile(hDrop,0xFFFFFFFF,FileName,256);
				if(ic < 1){
					MessageBox(NULL,_T("ファイルをドロップできませんでした"),_T("Error"), MB_OK | MB_ICONERROR);
					return dst;
				}
			}
			
			{// 取得
				POINT pDrop;
				DragQueryPoint(hDrop,&pDrop);
				for(int i=0; i < ic; i++)
				{// ファイル数分
					const int StrLength = DragQueryFile(hDrop,i,FileName,MAX_PATH);
					FileName[StrLength] = '\0';
					chl::FilePInfo buf;
					{
						const TCHAR* dir = FileName;
						buf.full = dir;
						buf.dir = chl::GetDirName(dir);
						buf.file = chl::GetFileName(dir);
					}
					dst.push_back(buf);
				}
				DragFinish(hDrop);
			}
		}
		SortFileData(dst);
		return dst;
	}

	
	/*------------------------------------------------------------------------
		[Function]
			SortFileData
		[Details]
			ファイル情報を昇順にソートする
		[Argument1] src
			ソートするデータへの参照
		[Remarks]
			構造体のvector配列はstd::sortの対象外だったのでメンバに分解してソートする
	------------------------------------------------------------------------*/
	static void SortFileData(std::vector<chl::FilePInfo>& src)
	{
		std::vector<chl::FilePInfo> Buf;
		{// ソートデータを取得
			std::vector<TSTR> sorttmp(0,_T(""));
			{// 比較用の構造体のメンバだけ取り出してソート
				for(UINT i=0; i<src.size(); i++){
					sorttmp.push_back(src[i].full);
				}
				std::sort(sorttmp.begin(), sorttmp.end());
			}
			for(UINT i=0; i<sorttmp.size(); i++)
			{// 上記でソートされたデータを利用して本ソートする
				for(UINT j=0; j<src.size(); j++)
				{
					if(sorttmp[i] == src[j].full)
					{
						Buf.push_back(src[j]);
						break;
					}
				}
			}
		}
		// コピー
		src = Buf;
	}


	/*------------------------------------------------------------------------
		[Function]
			ListViewAddItem
		[Details]
			リストビューへアイテムを1つ追加する
		[Argument1] hCtrl
			コントロールへのウインドウハンドル
		[Argument2] iItem
			アイテム番号
		[Argument3] SubItem
			項目番号
		[Argument4] Text
			挿入文字列
		[Return]
			FALSEを返します
			※この操作がイベントの最後にくるときのみ戻り値を評価します
	------------------------------------------------------------------------*/
	static void ListViewAddItem(const HWND hCtrl, const int& iItem, const int& SubItem, const TCHAR* const Text)
	{
		LVITEM item;
		{
			//item.mask = LVIF_TEXT | LVIF_PARAM;
			item.mask = LVIF_TEXT;
			item.iItem = iItem;
			item.iSubItem = SubItem;
			item.pszText = (LPWSTR)Text;
			//item.lParam = 0;
		}
		if(!SubItem)
			ListView_InsertItem(hCtrl, &item);
		else
			ListView_SetItem(hCtrl , &item);
	}


	/*------------------------------------------------------------------------
		[Function]
			UpdateFileList
		[Details]
			ファイルリストの状態を更新する
		[Argument1] hCtrl
			コントロールへのウインドウハンドル
	------------------------------------------------------------------------*/
	static void UpdateFileList(const HWND hCtrl)
	{
		// 一旦リストの中身をクリア
		ListView_DeleteAllItems(hCtrl);
		
		for(ULONG i=0; i<Fileinfo.size(); i++)
		{// 現在の状況で更新
			ListViewAddItem(hCtrl,i,0,Fileinfo[i].file.data());
			ListViewAddItem(hCtrl,i,1,Fileinfo[i].dir.data());
		}
	}


	/*------------------------------------------------------------------------
		[Function]
			AddFileList
		[Details]
			読み込んだファイルをリストへ追加する
		[Argument1] hCtrl
			コントロールへのウインドウハンドル
		[Argument2] Addsrc
			追加したいデータ
		[Return]
			FALSEを返します
			※この操作がイベントの最後にくるときのみ戻り値を評価します
	------------------------------------------------------------------------*/
	static void AddFileList(const HWND hCtrl, const std::vector<chl::FilePInfo>& Addsrc)
	{
		std::vector<chl::FilePInfo> Addbuf;
		{// 重複していないファイルだけ取り出す
			for(ULONG i=0; i<Addsrc.size(); i++)
			{// チェックloop
				BOOL find = FALSE;
				for(ULONG j=0; j<Fileinfo.size(); j++)
				{
					if(Addsrc[i].full == Fileinfo[j].full){
						find = TRUE;
						break;
					}
				}
				if(!find)
					Addbuf.push_back(Addsrc[i]);
			}
		}

		for(ULONG i=0; i<Addbuf.size(); i++)
		{// 重複しなかったものを追加
			Fileinfo.push_back(Addbuf[i]);
		}

		// 更新
		UpdateFileList(hCtrl);
	}




/*-+-+-+-+-+ サブクラスサポート用クラス -+-+-+-+-+*/
// ※ 2011/1/28 それぞれ共通の処理が多いので、出来れば抽象クラスから継承して作れるようにすること


	/**************************************************************************
		[Class]
			COutdirSubcls
		[Details]
			出力先入力エディットのサブクラス化をサポートするクラス
	**************************************************************************/
	class COutdirSubcls
	{

	/* Constructor / Destructor */
	public:
		COutdirSubcls(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, const int& hID)
		{
			hCtrl = GetDlgItem(hWnd, hID);
			BOOL result = SetProp(hCtrl, Outdir_Propname, this) != 0;
			if(result)
				DefStaticProc = SubclassWindow(hCtrl, CtrlWndProc);
		};
		~COutdirSubcls(void)
		{
		};


	/* コピー系を無効にする */
	private:
		COutdirSubcls(COutdirSubcls& obj){};
		COutdirSubcls& operator = (const COutdirSubcls& obj);


	/* Member */
	protected:
		HWND hCtrl;
		WNDPROC DefStaticProc;

	
	protected:
		/*------------------------------------------------------------------------
			[Function]
				MsgProc
			[Details]
				メッセージ処理専用関数
			[Argument1] hDlg
				ダイアログのウインドウハンドル
			[Argument2] uMsg
				取得メッセージ
			[Argument3] wParam
				メッセージの最初のパラメータ
			[Argument4] lParam
				メッセージの2番目のパラメータ
			[Return]
				メッセージ処理の結果が返ります。
				戻り値の意味は、送信されたメッセージによって異なります。
				初期化時を除き、Falseが返ります
		------------------------------------------------------------------------*/
		LRESULT CALLBACK MsgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
		{
			if(uMsg == WM_DROPFILES)
				WriteOutdir(hDlg,ReadDropfile(wParam));
			return CallWindowProc(DefStaticProc,hDlg,uMsg,wParam,lParam);
		}


	private:
		/*------------------------------------------------------------------------
			[Function]
				WriteOutdir
			[Details]
				読み込んだファイルを出力先エディットへ書き込む
			[Argument1] hCtrl
				コントロールへのウインドウハンドル
			[Argument2] Addsrc
				追加したいデータ
		------------------------------------------------------------------------*/
		void WriteOutdir(const HWND hCtrl, const std::vector<chl::FilePInfo>& Addsrc)
		{
			if(Addsrc.size()<1)
				return;	
			SetWindowText(hCtrl,Addsrc[0].full.data());
			//GetWindowText(ghDlg,buf,1000);
		}


	/* CALLBACK Function */
	private:
		/*------------------------------------------------------------------------
			[Function]
				CtrlWndProc
			[Details]
				ウインドへ組み込むコールバック関数
			[Argument1] hDlg
				ダイアログのウインドウハンドル
			[Argument2] uMsg
				取得メッセージ
			[Argument3] wParam
				メッセージの最初のパラメータ
			[Argument4] lParam
				メッセージの2番目のパラメータ
			[Return]
				メッセージ処理の結果が返ります。
				戻り値の意味は、送信されたメッセージによって異なります。
				初期化時を除き、Falseが返ります
		------------------------------------------------------------------------*/
		static LRESULT CALLBACK CtrlWndProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
		{
			// 既存のコントロールからthisポインタ取得
			COutdirSubcls* outdircls = reinterpret_cast<COutdirSubcls*>(::GetProp(hDlg, Outdir_Propname));
			if(!outdircls)
				return -1;
			// ウィンドウハンドルセット
			outdircls->hCtrl = hDlg;
			// メッセージ処理関数コール
			return outdircls->MsgProc(hDlg, uMsg, wParam, lParam);
		}

	};// COutdirSubcls
	
	
	
	/**************************************************************************
		[Class]
			CFilelistSubcls
		[Details]
			ファイルリストのサブクラス化をサポートするクラス
	**************************************************************************/
	class CFilelistSubcls
	{
	/* Constructor / Destructor */
	public:
		CFilelistSubcls(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, const int& hID)
		{
			hCtrl = GetDlgItem(hWnd, hID);
			BOOL result = SetProp(hCtrl, Filelist_Propname, this) != 0;
			if(result)
				DefStaticProc = SubclassWindow(hCtrl, CtrlWndProc);
		};
		~CFilelistSubcls(void)
		{
		};


	/* コピー系を無効にする */
	private:
		CFilelistSubcls(CFilelistSubcls& obj){};
		CFilelistSubcls& operator = (const CFilelistSubcls& obj);


	/* Member */
	protected:
		HWND hCtrl;
		WNDPROC DefStaticProc;


	protected:
		/*------------------------------------------------------------------------
			[Function]
				MsgProc
			[Details]
				メッセージ処理専用関数
			[Argument1] hDlg
				ダイアログのウインドウハンドル
			[Argument2] uMsg
				取得メッセージ
			[Argument3] wParam
				メッセージの最初のパラメータ
			[Argument4] lParam
				メッセージの2番目のパラメータ
			[Return]
				メッセージ処理の結果が返ります。
				戻り値の意味は、送信されたメッセージによって異なります。
				初期化時を除き、Falseが返ります
		------------------------------------------------------------------------*/
		LRESULT CALLBACK MsgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
		{
			if(uMsg == WM_DROPFILES)
				AddFileList(hDlg,ReadDropfile(wParam));
			return CallWindowProc(DefStaticProc,hDlg,uMsg,wParam,lParam);
		}


	/* CALLBACK Function */
	private:
		/*------------------------------------------------------------------------
			[Function]
				CtrlWndProc
			[Details]
				ウインドへ組み込むコールバック関数
			[Argument1] hDlg
				ダイアログのウインドウハンドル
			[Argument2] uMsg
				取得メッセージ
			[Argument3] wParam
				メッセージの最初のパラメータ
			[Argument4] lParam
				メッセージの2番目のパラメータ
			[Return]
				メッセージ処理の結果が返ります。
				戻り値の意味は、送信されたメッセージによって異なります。
				初期化時を除き、Falseが返ります
		------------------------------------------------------------------------*/
		static LRESULT CALLBACK CtrlWndProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
		{
			// 既存のコントロールからthisポインタ取得
			CFilelistSubcls* FilelistSub = reinterpret_cast<CFilelistSubcls*>(::GetProp(hDlg, Filelist_Propname));
			if(!FilelistSub)
				return -1;
			// ウィンドウハンドルセット
			FilelistSub->hCtrl = hDlg;
			// メッセージ処理関数コール
			return FilelistSub->MsgProc(hDlg, uMsg, wParam, lParam);
		}

	};// CFilelistSubcls

}; // CMainDialog



// MEMO
	
// 元からGUIにコントロールが存在する場合、以下の要領でウインドウハンドルを割り当てる
// hList = GetDlgItem(hWnd, リストビューID);

// リストビューのフォーカス状態を取得する場合
//const UINT uiState = ListView_GetItemState(hList, item, LVIS_SELECTED | LVIS_FOCUSED);
//if(uiState & LVIS_FOCUSED){ フォーカスされている;}

/* ウインドウに直接リストコントロールを作成する方法
long width = 0, height = 0;
{// メインダイアログサイズ取得
	RECT rect;
	::GetWindowRect(hDlg, &rect);
	width = rect.right - rect.left;
	height = rect.bottom - rect.top;
}

{// 作成したリストビューのハンドル取得
	hList = CreateWindowEx(
		NULL,
		WC_LISTVIEW, _T(""),
		WS_CHILD | WS_VISIBLE | LVS_REPORT,
		0, LISTBOTTOM, width, static_cast<int>(height*0.5),
		hDlg,
		(HMENU)0,
		ghInstance,
		NULL
	);
}
*/
