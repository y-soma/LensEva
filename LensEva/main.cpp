#pragma once

#include "MainDialog.h"

/**************************************************************************
	[Function]
		WinMain
	[Details]
		Application Entry Point
	[Argument1] hInstance
		Win32インスタンスハンドル
	[Argument2] hPrevInstance
		Win16インスタンスハンドル(NULL)
	[Argument3] lpCmdLine
		コマンドライン引数
	[Argument4] nCmdShow
		アプリケーションの初期表示方法(表示状態)
	[Return]
		関数が WM_QUIT メッセージを受け取って正常に終了する場合は、
		メッセージの wParam パラメータに格納されている終了コードを返してください。
		関数がメッセージループに入る前に終了する場合は、0 を返してください。
**************************************************************************/
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR pCmdLine, int iCmdShow)
{
	CMainDialog dlg;
	return dlg.DlgCreate(hInstance);
}

