#pragma once

#include "MainDialog.h"

/**************************************************************************
	[Function]
		WinMain
	[Details]
		Application Entry Point
	[Argument1] hInstance
		Win32�C���X�^���X�n���h��
	[Argument2] hPrevInstance
		Win16�C���X�^���X�n���h��(NULL)
	[Argument3] lpCmdLine
		�R�}���h���C������
	[Argument4] nCmdShow
		�A�v���P�[�V�����̏����\�����@(�\�����)
	[Return]
		�֐��� WM_QUIT ���b�Z�[�W���󂯎���Đ���ɏI������ꍇ�́A
		���b�Z�[�W�� wParam �p�����[�^�Ɋi�[����Ă���I���R�[�h��Ԃ��Ă��������B
		�֐������b�Z�[�W���[�v�ɓ���O�ɏI������ꍇ�́A0 ��Ԃ��Ă��������B
**************************************************************************/
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR pCmdLine, int iCmdShow)
{
	CMainDialog dlg;
	return dlg.DlgCreate(hInstance);
}

