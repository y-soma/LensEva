#pragma once

// CMainDialog : ��`

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

// �ǂݍ��񂾃t�@�C�����Ǘ�
std::vector<chl::FilePInfo> Fileinfo;
// �]���{�f�B
std::vector<TSTR> Evabody;
// �o�͐�
TSTR Outdir;

// �t�@�C�����X�g�T�u�N���X�v���p�e�B�ݒ�
const TCHAR Filelist_Propname[] = _T("FilelistProp");
// �o�͐�G�f�B�b�g�v���p�e�B��
const TCHAR Outdir_Propname[] = _T("OutdirProp");



// disable
#pragma warning(disable:4311)
#pragma warning(disable:4312)
#pragma warning(disable:4244)


/**************************************************************************
	[Class]
		CMainDialog
	[Details]
		���C���_�C�A���O
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
			�_�C�A���O�쐬
		[Argument1] hInstance
			�C���X�^���X�n���h��
		[Return]
			False��Ԃ��܂�
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
			�_�C�A���O�{�b�N�X�̃R�[���o�b�N�֐�
		[Argument1] hDlg
			�_�C�A���O�̃E�C���h�E�n���h��
		[Argument2] uMsg
			�擾���b�Z�[�W
		[Argument3] wParam
			���b�Z�[�W�̍ŏ��̃p�����[�^
		[Argument4] lParam
			���b�Z�[�W��2�Ԗڂ̃p�����[�^
		[Return]
			���b�Z�[�W�����̌��ʂ��Ԃ�܂��B
			�߂�l�̈Ӗ��́A���M���ꂽ���b�Z�[�W�ɂ���ĈقȂ�܂��B
			���������������AFalse���Ԃ�܂�
	------------------------------------------------------------------------*/
	static LRESULT CALLBACK DlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		static const TCHAR s_prop_name[] = _T("MainDlgProp");
		
		switch(uMsg)
		{
			case WM_INITDIALOG:
			{// ��������
				// �����̃R���g���[������this�|�C���^�擾
				CMainDialog* md = reinterpret_cast<CMainDialog*>(GetWindowLong(hDlg,GWL_EXSTYLE));
				
				/* ���쐬�����E�C���h�E����擾����ꍇ�͂�����
				::CREATESTRUCT* cs = reinterpret_cast<::CREATESTRUCT*>(lParam);
				if(!cs)
					return -1;
				CMainDialog* md = reinterpret_cast<CMainDialog*>(cs->lpCreateParams);
				*/
				
				if(!md)
					return -1;
				// �v���p�e�B��Window�N���X�̃|�C���^��ݒ肷��
				if( !::SetProp(hDlg, s_prop_name, reinterpret_cast<HANDLE>(md)) )
					return -1;

				// �E�B���h�E�n���h���Z�b�g
				md->ghDlg = hDlg;
				// ���b�Z�[�W�����֐��R�[��
				return md->MsgProc(hDlg, uMsg, wParam, lParam);
			}
			case WM_CLOSE:
			{// �I����
				CMainDialog* md = reinterpret_cast<CMainDialog*>(::GetProp(hDlg, s_prop_name));

				LRESULT ret = 0;
				// ���b�Z�[�W�����֐��R�[��
				if(md != NULL)
				{
					ret = md->MsgProc(hDlg, uMsg, wParam, lParam);
					md->ghDlg = NULL;
				}
				else
				{
					ret = ::DefWindowProc(hDlg, uMsg, wParam, lParam);
				}

				// �ݒ肵���v���p�e�B�̃f�[�^���폜����
				::RemoveProp(hDlg, s_prop_name);
				return ret;
			}
			default:
			{// ���̑�
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
			�_�C�A���O�{�b�N�X�̃R�[���o�b�N�֐�
		[Argument1] hDlg
			�_�C�A���O�̃E�C���h�E�n���h��
		[Argument2] uMsg
			�擾���b�Z�[�W
		[Argument3] wParam
			���b�Z�[�W�̍ŏ��̃p�����[�^
		[Argument4] lParam
			���b�Z�[�W��2�Ԗڂ̃p�����[�^
		[Return]
			���b�Z�[�W�����̌��ʂ��Ԃ�܂��B
			�߂�l�̈Ӗ��́A���M���ꂽ���b�Z�[�W�ɂ���ĈقȂ�܂��B
			���������������AFalse���Ԃ�܂�
	------------------------------------------------------------------------*/
	LRESULT CALLBACK MsgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		switch(uMsg)
		{
			case WM_INITDIALOG:
			{
				// ���X�g�R���g���[�����T�u�N���X��
				static CFilelistSubcls FileList(hDlg,uMsg,wParam,lParam,IDC_FILELIST);
				// �o�͐�G�f�B�b�g���T�u�N���X��
				static COutdirSubcls Outdir(hDlg,uMsg,wParam,lParam,IDC_EDIT_OUTDIR);
				InitCommonControls();
				return HANDLE_WM_INITDIALOG(hDlg, wParam, lParam, OnInitDialog);
			}
			case WM_COMMAND:
				return GetWMCOMMAND(wParam);
			//case WM_DROPFILES:	
				//���b�Z�[�W���t�@�C�����X�g�T�u�N���X���֏��n
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
			���̗̈���̃O���[�o���ϐ�������������
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
			�_�C�A���O����ʂ̒��S�Ɉړ������� : �e���Ȃ��̂ŉ������Ȃ��ƍ���ɔz�u�����
		[Argument1] hwnd
			�E�C���h�E�n���h��
	------------------------------------------------------------------------*/
	void MoveCenter(HWND hDlg)
	{
		int w = 0, h = 0, wpos = 0, hpos = 0;
		{// �_�C�A���O�̃T�C�Y�ƈړ��ʒu���擾����
			{// �_�C�A���O�̏c���T�C�Y
				RECT rc;
				GetWindowRect(hDlg, &rc);
				w = rc.right - rc.left;
				h = rc.bottom - rc.top;
			}
			{// �X�N���[���T�C�Y����v�Z
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
			�_�C�A���O����ʂ̒����ɔz�u����(���@����2)
		[Argument1] hwnd
			�E�C���h�E�n���h��
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
			���X�g�r���[������
		[Argument1] hDlg
			�E�C���h�E�n���h��
	------------------------------------------------------------------------*/
	void InitListView(const HWND hDlg)
	{
		// �����̃��X�g�r���[�̃n���h�����擾
		hList = GetDlgItem(hDlg, IDC_FILELIST);
		
		{// ���X�g�r���[�̃X�^�C���w��
			DWORD dwStyle;
			dwStyle = ListView_GetExtendedListViewStyle(hList);
			dwStyle |= LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES;   // LVS_EX_HEADERDRAGDROP; �h���b�O���h���b�v�\  LVS_EX_CHECKBOXES | �`�F�b�N�{�b�N�X������
			ListView_SetExtendedListViewStyle(hList, dwStyle);
		}

		LV_COLUMN lvcol;
		{
			lvcol.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
			lvcol.fmt = LVCFMT_LEFT;
		}

		TCHAR caption[][32] = {_T("FILE"), _T("DIR")};
		const UINT clmNum = sizeof caption /sizeof caption[0];
	    
		// �{���w��
		const double mag[] = { 0.295,0.695 };
		
		if(sizeof(caption)/sizeof(caption[0]) != sizeof(mag)/sizeof(mag[0]))
			return;

		long width = 0;
		{// ���X�g�R���g���[���̕��T�C�Y�擾
			RECT rect;
			::GetWindowRect(hList, &rect);
			width = rect.right - rect.left;
			//height = rect.bottom - rect.top;
		}
		for(int i = 0; i < sizeof(caption)/sizeof(caption[0]); i++)
		{
			// �T�u�A�C�e���ԍ�
			lvcol.iSubItem = i;
			// ���o���e�L�X�g
			lvcol.pszText = caption[i];
			
			{// �����ݒ�
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
			�R���{�{�b�N�X������
		[Argument1] hDlg
			�E�C���h�E�n���h��
	------------------------------------------------------------------------*/
	void InitEvabodyComboBox(const HWND hDlg)
	{
		// �����̃R���{�{�b�N�X�̃n���h�����擾
		hCombo = GetDlgItem(hDlg, IDC_COMBO_EVABODY);
		{// ini�t�@�C������擾�����{�f�B�����R���{�{�b�N�X�֒ǉ�
			const TSTR evabody = chl::GetInitParam(EVABODY_INI,_T("BODY"),_T("value"));
			if(evabody != NOT_FOUND)
			{
				const std::vector<TSTR> bodyname = chl::Split(evabody.data(),_T(","));
				for(UINT i=0; i<bodyname.size(); i++)
				{
					SendMessage(hCombo,CB_INSERTSTRING,i,(LPARAM)bodyname[i].data());
				}
				// ��Ԗڂ̃A�C�e����I��
				SendMessage(hCombo, CB_SETCURSEL, 0, 0);
				Evabody = bodyname;
			}
		}
	}


	/*------------------------------------------------------------------------
		[Function]
			UpdateEvabodyComboBox
		[Details]
			�R���{�{�b�N�X���X�V����
		[Argument1] hDlg
			�E�C���h�E�n���h��
	------------------------------------------------------------------------*/
	void UpdateEvabodyComboBox(void)
	{
		while(SendMessage(hCombo, CB_GETCOUNT, 0, 0) != 0)
		{// �R���{�{�b�N�X�̒��g��S�č폜
			SendMessage(hCombo, CB_DELETESTRING, 0, 0);
		}

		for(UINT i=0; i<Evabody.size(); i++)
		{
			SendMessage(hCombo,CB_INSERTSTRING,i,(LPARAM)Evabody[i].data());
		}
		// ��Ԗڂ̃A�C�e����I��
		SendMessage(hCombo, CB_SETCURSEL, 0, 0);
	}


	/*------------------------------------------------------------------------
		[Function]
			SelectedListDelete
		[Details]
			���X�g����I�����Ă���t�@�C����1�폜����
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

		// �X�V
		UpdateFileList(hDlg);
	}


	/*------------------------------------------------------------------------
		[Function]
			OnInitDialog
		[Details]
			�_�C�A���O������
		[Return]
			FALSE��Ԃ�
	------------------------------------------------------------------------*/
	BOOL OnInitDialog(const HWND hDlg, const HWND hwndFocus, const LPARAM lParam)
	{
		// ���X�g�r���[������
		InitListView(hDlg);
		// �R���{�{�b�N�X������
		InitEvabodyComboBox(hDlg);
		// �o�͐�f�B���N�g���G�f�B�b�g�̃n���h���擾
		hOutdir = GetDlgItem(hDlg, IDC_EDIT_OUTDIR);
		// �E�C���h�E�𒆉��֔z�u
		//center_window(hDlg);
		MoveCenter(hDlg);
		return FALSE;
	}


	/*------------------------------------------------------------------------
		[Function]
			ReadFileDlg
		[Details]
			�_�C�A���O����t�@�C���ǂݍ���
		[Return]
			�ǂݍ��݃t�@�C�����
	------------------------------------------------------------------------*/
	std::vector<chl::FilePInfo> ReadFileDlg(void)
	{
		std::vector<chl::FilePInfo> dst;
		{
			static OPENFILENAME ofn;
			static TCHAR filename_full[MAX_PATH*0xFF];
			static TCHAR filename_n[MAX_PATH*0xFF];
			{// �_�C�A���O�����ݒ�
				ZeroMemory(&ofn, sizeof(ofn));
				ZeroMemory(&filename_full, sizeof(filename_full));
				ZeroMemory(&filename_n, sizeof(filename_n));
				ofn.lStructSize = sizeof(ofn);
				ofn.hwndOwner = ghDlg;
				ofn.lpstrFile = filename_full;
				ofn.nMaxFile = sizeof(filename_full);
				ofn.lpstrFilter = _T("All files(*.*)\0*.*\0\0");
				ofn.lpstrTitle = _T("�t�@�C�����J��");
				ofn.lpstrDefExt = _T("csv");
				ofn.lpstrFileTitle = filename_n;
				ofn.nMaxFileTitle = sizeof(filename_n);
				ofn.Flags = OFN_EXPLORER | OFN_ALLOWMULTISELECT;
			}

			if(!GetOpenFileName(&ofn))
				return dst;

			/* ��������ƕۑ��_�C�A���O�ɂȂ�
			ofn->Flags = OFN_OVERWRITEPROMPT;
			ofn->lpstrTitle = _T("���O��t���ĕۑ�");
			if(!GetSaveFileName(&ofn))
				return dst;
			*/

			// �p�X�̂݃Z�b�g
			const TSTR path = filename_full;

			for(ULONG i=ofn.nFileOffset-1; i<ofn.nMaxFile; i++)
			{// �S�Ẵt�@�C�����擾
				if(filename_full[i])
					continue;
				
				if(!filename_full[i+1])
				{// 1��̏ꏊ��NULL�̏ꍇ�A�I��
					break;
				}
				else
				{// NULL����Ȃ��ꍇ�A�t�@�C����񂪑����Ă���
					ULONG cp = i+1;
					TSTR fNameTemp = _T("");
					while(filename_full[cp] != NULL)
					{// ����NULL�|�C���g�܂ł�1�����P�ʂŌ���������
						fNameTemp += filename_full[cp];
						cp++;
					}
					{// �f�[�^�������Ď擾
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
			{// �t�@�C����1�����������ꍇ
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
			�_�C�A���O����t�H���_�I��
		[Return]
			�����F�I�������f�B���N�g���@�@���s�Fempty
	------------------------------------------------------------------------*/
	TSTR ReadDirDlg(void)
	{
		TSTR dst = _T("");
		{
			BROWSEINFO bInfo;
			TCHAR szDisplayName[MAX_PATH*0xFF];
			{// �_�C�A���O�\���̂̏���
				bInfo.hwndOwner             = ghDlg;          // �_�C�A���O�̐e�E�C���h�E�̃n���h��
				bInfo.pidlRoot              = NULL;                             // ���[�g�t�H���_������ITEMIDLIST�̃|�C���^ (NULL�̏ꍇ�f�X�N�g�b�v�t�H���_���g���܂��j
				bInfo.pszDisplayName        = szDisplayName;                    // �I�����ꂽ�t�H���_�����󂯎��o�b�t�@�̃|�C���^
				bInfo.lpszTitle             = _T("�t�H���_�̑I��");             // �c���[�r���[�̏㕔�ɕ\������镶���� 
				bInfo.ulFlags               = BIF_RETURNONLYFSDIRS;             // �\�������t�H���_�̎�ނ������t���O
				bInfo.lpfn                  = NULL;                             // BrowseCallbackProc�֐��̃|�C���^
				bInfo.lParam                = (LPARAM)0;                        // �R�[���o�b�N�֐��ɓn���l
			}

			// �t�H���_�I���_�C�A���O��\��
			LPITEMIDLIST pIDList = ::SHBrowseForFolder(&bInfo);
			if(pIDList)
			{
				if(::SHGetPathFromIDList(pIDList, szDisplayName))
					dst = szDisplayName;
				else
					MessageBox(ghDlg, _T("�t�H���_�I���Ɏ��s���܂���"), _T("Error"), MB_OKCANCEL | MB_ICONERROR);
				// �|�C���^�̉��
				::CoTaskMemFree(pIDList);
			}
		}
		return dst;
	}


	/*------------------------------------------------------------------------
		[Function]
			GetWindowTxtEX
		[Details]
			�R���g���[���ɕ\������Ă���e�L�X�g���擾����
		[Argument1] gWnd
			�w��E�C���h�E�n���h��
		[Return]
			�\������Ă���e�L�X�g��������ŕԂ�
	------------------------------------------------------------------------*/
	TSTR GetWindowTxtEX(const HWND& gWnd)
	{
		TSTR dst = _T("");
		{// �擾
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
			���������s����
	------------------------------------------------------------------------*/
	void Execution(void)
	{
		try
		{
			if(Fileinfo.size() < 1){
				MessageBox(ghDlg, _T("�t�@�C�����I������Ă��܂���"), _T("NO FILES"), MB_OKCANCEL | MB_ICONERROR);
				return;
			}
			{// ���s
				CLProc Proc;
				Proc.Execution(Fileinfo,GetWindowTxtEX(hCombo),GetWindowTxtEX(hOutdir));
			}
		}
		catch(...)
		{
			MessageBox(ghDlg, _T("�������ɒv���I�ȃG���[���������܂���"), _T("Fatal Error"), MB_OKCANCEL | MB_ICONERROR);
		}
	}


	/*------------------------------------------------------------------------
		[Function]
			GetWMCOMMAND
		[Details]
			�R�}���h�n���b�Z�[�W�����擾
		[Argument1] wParam
			�R�}���h�n���b�Z�[�W
		[Return]
			���b�Z�[�W�����̌��ʂ��Ԃ�܂��B
			�߂�l�̈Ӗ��́A���M���ꂽ���b�Z�[�W�ɂ���ĈقȂ�܂��B
	------------------------------------------------------------------------*/
	LRESULT GetWMCOMMAND(const WPARAM wParam)
	{
		switch(LOWORD(wParam))
		{
			// �L�����Z���{�^��
			case IDCANCEL:
				EndDialog(ghDlg, FALSE);
				return FALSE;
			// ���s�{�^��
			case IDC_BUTTON_EXECUTION:
			{
				Execution();
				return FALSE;
			}
			// ���j���[����t�@�C�����J��
			case ID_FILEOPEN:
				AddFileList(hList,ReadFileDlg());
				return FALSE;
			// �I������Ă���t�@�C����1�폜
			case ID_FILEDELETE:
				SelectedListDelete(hList);
				return FALSE;
			// �t�@�C����S�č폜
			case ID_ALLCLEAR:
				Fileinfo.clear();
				ListView_DeleteAllItems(hList);
				return FALSE;
			// ����
			case ID_FILECLOSE:
				EndDialog(ghDlg, FALSE);
				return FALSE;
			// �o�͐���_�C�A���O�őI��
			case IDC_OUTDIR_SELECT:
			{
				const TSTR outdir = ReadDirDlg();
				if(outdir != _T(""))
					SetWindowText(hOutdir,outdir.data());
				return FALSE;
			}
			// �o�͐���폜
			case IDC_OUTDIR_CLEAR:
			{
				SetWindowText(hOutdir,_T(""));
				return FALSE;
			}
			// �o�[�W�������
			case ID_VERSION:
			{
				TSTR mes = _T("LensEva Version ");
				mes += THIS_VERSION;
				mes += _T("\n");
				mes += COPYRIGHT;
				MessageBox(ghDlg, mes.data(), _T("�o�[�W�������"), MB_OK | MB_ICONINFORMATION);
				return FALSE;
			}
			// �R���{�{�b�N�X�֘A
			case IDC_COMBO_EVABODY:
				// �I���Ȃǂ̓���ȃC�x���g�n���h���𓾂����ꍇ�� HIWORD(wParam) ��]������
				if(HIWORD(wParam) == CBN_EDITCHANGE)
				{// ���ړ��͌�̃C�x���g�n���h��
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
			�h���b�O���h���b�v�t�@�C���ǂݍ���
		[Argument1] hWnd
			�E�C���h�E�n���h��
		[Argument2] wParam
			�p�����[�^
		[Return]
			�ǂݍ��݃t�@�C�����
	------------------------------------------------------------------------*/
	static std::vector<chl::FilePInfo> ReadDropfile(const WPARAM wParam)
	{
		std::vector<chl::FilePInfo> dst;
		{
			TCHAR FileName[MAX_PATH * 0xFF] = _T("");
			HDROP hDrop = (HDROP)wParam;

			int ic = 0;
			{// �h���b�v�t�@�C�������擾
				ic = DragQueryFile(hDrop,0xFFFFFFFF,FileName,256);
				if(ic < 1){
					MessageBox(NULL,_T("�t�@�C�����h���b�v�ł��܂���ł���"),_T("Error"), MB_OK | MB_ICONERROR);
					return dst;
				}
			}
			
			{// �擾
				POINT pDrop;
				DragQueryPoint(hDrop,&pDrop);
				for(int i=0; i < ic; i++)
				{// �t�@�C������
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
			�t�@�C�����������Ƀ\�[�g����
		[Argument1] src
			�\�[�g����f�[�^�ւ̎Q��
		[Remarks]
			�\���̂�vector�z���std::sort�̑ΏۊO�������̂Ń����o�ɕ������ă\�[�g����
	------------------------------------------------------------------------*/
	static void SortFileData(std::vector<chl::FilePInfo>& src)
	{
		std::vector<chl::FilePInfo> Buf;
		{// �\�[�g�f�[�^���擾
			std::vector<TSTR> sorttmp(0,_T(""));
			{// ��r�p�̍\���̂̃����o�������o���ă\�[�g
				for(UINT i=0; i<src.size(); i++){
					sorttmp.push_back(src[i].full);
				}
				std::sort(sorttmp.begin(), sorttmp.end());
			}
			for(UINT i=0; i<sorttmp.size(); i++)
			{// ��L�Ń\�[�g���ꂽ�f�[�^�𗘗p���Ė{�\�[�g����
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
		// �R�s�[
		src = Buf;
	}


	/*------------------------------------------------------------------------
		[Function]
			ListViewAddItem
		[Details]
			���X�g�r���[�փA�C�e����1�ǉ�����
		[Argument1] hCtrl
			�R���g���[���ւ̃E�C���h�E�n���h��
		[Argument2] iItem
			�A�C�e���ԍ�
		[Argument3] SubItem
			���ڔԍ�
		[Argument4] Text
			�}��������
		[Return]
			FALSE��Ԃ��܂�
			�����̑��삪�C�x���g�̍Ō�ɂ���Ƃ��̂ݖ߂�l��]�����܂�
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
			�t�@�C�����X�g�̏�Ԃ��X�V����
		[Argument1] hCtrl
			�R���g���[���ւ̃E�C���h�E�n���h��
	------------------------------------------------------------------------*/
	static void UpdateFileList(const HWND hCtrl)
	{
		// ��U���X�g�̒��g���N���A
		ListView_DeleteAllItems(hCtrl);
		
		for(ULONG i=0; i<Fileinfo.size(); i++)
		{// ���݂̏󋵂ōX�V
			ListViewAddItem(hCtrl,i,0,Fileinfo[i].file.data());
			ListViewAddItem(hCtrl,i,1,Fileinfo[i].dir.data());
		}
	}


	/*------------------------------------------------------------------------
		[Function]
			AddFileList
		[Details]
			�ǂݍ��񂾃t�@�C�������X�g�֒ǉ�����
		[Argument1] hCtrl
			�R���g���[���ւ̃E�C���h�E�n���h��
		[Argument2] Addsrc
			�ǉ��������f�[�^
		[Return]
			FALSE��Ԃ��܂�
			�����̑��삪�C�x���g�̍Ō�ɂ���Ƃ��̂ݖ߂�l��]�����܂�
	------------------------------------------------------------------------*/
	static void AddFileList(const HWND hCtrl, const std::vector<chl::FilePInfo>& Addsrc)
	{
		std::vector<chl::FilePInfo> Addbuf;
		{// �d�����Ă��Ȃ��t�@�C���������o��
			for(ULONG i=0; i<Addsrc.size(); i++)
			{// �`�F�b�Nloop
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
		{// �d�����Ȃ��������̂�ǉ�
			Fileinfo.push_back(Addbuf[i]);
		}

		// �X�V
		UpdateFileList(hCtrl);
	}




/*-+-+-+-+-+ �T�u�N���X�T�|�[�g�p�N���X -+-+-+-+-+*/
// �� 2011/1/28 ���ꂼ�ꋤ�ʂ̏����������̂ŁA�o����Β��ۃN���X����p�����č���悤�ɂ��邱��


	/**************************************************************************
		[Class]
			COutdirSubcls
		[Details]
			�o�͐���̓G�f�B�b�g�̃T�u�N���X�����T�|�[�g����N���X
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


	/* �R�s�[�n�𖳌��ɂ��� */
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
				���b�Z�[�W������p�֐�
			[Argument1] hDlg
				�_�C�A���O�̃E�C���h�E�n���h��
			[Argument2] uMsg
				�擾���b�Z�[�W
			[Argument3] wParam
				���b�Z�[�W�̍ŏ��̃p�����[�^
			[Argument4] lParam
				���b�Z�[�W��2�Ԗڂ̃p�����[�^
			[Return]
				���b�Z�[�W�����̌��ʂ��Ԃ�܂��B
				�߂�l�̈Ӗ��́A���M���ꂽ���b�Z�[�W�ɂ���ĈقȂ�܂��B
				���������������AFalse���Ԃ�܂�
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
				�ǂݍ��񂾃t�@�C�����o�͐�G�f�B�b�g�֏�������
			[Argument1] hCtrl
				�R���g���[���ւ̃E�C���h�E�n���h��
			[Argument2] Addsrc
				�ǉ��������f�[�^
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
				�E�C���h�֑g�ݍ��ރR�[���o�b�N�֐�
			[Argument1] hDlg
				�_�C�A���O�̃E�C���h�E�n���h��
			[Argument2] uMsg
				�擾���b�Z�[�W
			[Argument3] wParam
				���b�Z�[�W�̍ŏ��̃p�����[�^
			[Argument4] lParam
				���b�Z�[�W��2�Ԗڂ̃p�����[�^
			[Return]
				���b�Z�[�W�����̌��ʂ��Ԃ�܂��B
				�߂�l�̈Ӗ��́A���M���ꂽ���b�Z�[�W�ɂ���ĈقȂ�܂��B
				���������������AFalse���Ԃ�܂�
		------------------------------------------------------------------------*/
		static LRESULT CALLBACK CtrlWndProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
		{
			// �����̃R���g���[������this�|�C���^�擾
			COutdirSubcls* outdircls = reinterpret_cast<COutdirSubcls*>(::GetProp(hDlg, Outdir_Propname));
			if(!outdircls)
				return -1;
			// �E�B���h�E�n���h���Z�b�g
			outdircls->hCtrl = hDlg;
			// ���b�Z�[�W�����֐��R�[��
			return outdircls->MsgProc(hDlg, uMsg, wParam, lParam);
		}

	};// COutdirSubcls
	
	
	
	/**************************************************************************
		[Class]
			CFilelistSubcls
		[Details]
			�t�@�C�����X�g�̃T�u�N���X�����T�|�[�g����N���X
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


	/* �R�s�[�n�𖳌��ɂ��� */
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
				���b�Z�[�W������p�֐�
			[Argument1] hDlg
				�_�C�A���O�̃E�C���h�E�n���h��
			[Argument2] uMsg
				�擾���b�Z�[�W
			[Argument3] wParam
				���b�Z�[�W�̍ŏ��̃p�����[�^
			[Argument4] lParam
				���b�Z�[�W��2�Ԗڂ̃p�����[�^
			[Return]
				���b�Z�[�W�����̌��ʂ��Ԃ�܂��B
				�߂�l�̈Ӗ��́A���M���ꂽ���b�Z�[�W�ɂ���ĈقȂ�܂��B
				���������������AFalse���Ԃ�܂�
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
				�E�C���h�֑g�ݍ��ރR�[���o�b�N�֐�
			[Argument1] hDlg
				�_�C�A���O�̃E�C���h�E�n���h��
			[Argument2] uMsg
				�擾���b�Z�[�W
			[Argument3] wParam
				���b�Z�[�W�̍ŏ��̃p�����[�^
			[Argument4] lParam
				���b�Z�[�W��2�Ԗڂ̃p�����[�^
			[Return]
				���b�Z�[�W�����̌��ʂ��Ԃ�܂��B
				�߂�l�̈Ӗ��́A���M���ꂽ���b�Z�[�W�ɂ���ĈقȂ�܂��B
				���������������AFalse���Ԃ�܂�
		------------------------------------------------------------------------*/
		static LRESULT CALLBACK CtrlWndProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
		{
			// �����̃R���g���[������this�|�C���^�擾
			CFilelistSubcls* FilelistSub = reinterpret_cast<CFilelistSubcls*>(::GetProp(hDlg, Filelist_Propname));
			if(!FilelistSub)
				return -1;
			// �E�B���h�E�n���h���Z�b�g
			FilelistSub->hCtrl = hDlg;
			// ���b�Z�[�W�����֐��R�[��
			return FilelistSub->MsgProc(hDlg, uMsg, wParam, lParam);
		}

	};// CFilelistSubcls

}; // CMainDialog



// MEMO
	
// ������GUI�ɃR���g���[�������݂���ꍇ�A�ȉ��̗v�̂ŃE�C���h�E�n���h�������蓖�Ă�
// hList = GetDlgItem(hWnd, ���X�g�r���[ID);

// ���X�g�r���[�̃t�H�[�J�X��Ԃ��擾����ꍇ
//const UINT uiState = ListView_GetItemState(hList, item, LVIS_SELECTED | LVIS_FOCUSED);
//if(uiState & LVIS_FOCUSED){ �t�H�[�J�X����Ă���;}

/* �E�C���h�E�ɒ��ڃ��X�g�R���g���[�����쐬������@
long width = 0, height = 0;
{// ���C���_�C�A���O�T�C�Y�擾
	RECT rect;
	::GetWindowRect(hDlg, &rect);
	width = rect.right - rect.left;
	height = rect.bottom - rect.top;
}

{// �쐬�������X�g�r���[�̃n���h���擾
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
