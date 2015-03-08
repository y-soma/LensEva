//#pragma once

#ifndef CCEXCEL_INCLUDE
#define CCEXCEL_INCLUDE


/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
 Excel�^�C�v���C�u�����ǂݍ���
   �r���h����PC���ɉ����ăC���|�[�g�ꏊ��ύX
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
// GetFirstChild�̎������s��v�x���𖳌��ɂ���(import�̑O�ł��)
#pragma warning(disable:4003)
#pragma warning(disable:4278)
#pragma warning(disable:4192)
//#pragma warning( disable : 4786 )
#import _MSODLL_PATH no_namespace rename("DocumentProperties", "DocumentPropertiesXL")   
#import _MSOVBE6EXT_PATH no_namespace
//#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7\VBE7.DLL" no_namespace
#import _MSEXCEL_PATH rename("DialogBox", "DialogBoxXL") rename("RGB", "RBGXL") rename("DocumentProperties", "DocumentPropertiesXL") no_dual_interfaces

#endif


// CCExcel : ��`
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
		Excel����T�|�[�g�N���X
	[Remarks]
		Win32SDK,MFC���p
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
			�u�b�N1������̃V�[�g�����Ǘ�����\����
	**************************************************************************/
	typedef struct EXCEL_SHEET_INFO {
		// WorkSheet�|�C���^
		std::vector<_WorksheetPtr> pSheet;
		// �V�[�g��
		std::vector<TSTR> name;
		// �V�[�g��
		long cnt;
	} eSheetInfo;
	
	/**************************************************************************
		[Structure]
			eFileInfo
		[Details]
			�J����Excel�t�@�C�����Ǘ�����\����
	**************************************************************************/
	typedef struct EXCEL_FILE_INFO {
		// �t�@�C���p�X���
		chl::FilePInfo fpinfo;
		// WorkBook�|�C���^
		_WorkbookPtr pBook;
		// Sheet���
		eSheetInfo sheetinfo;
	} eFileInfo;
	
	// ExcelControlPtr
	_ApplicationPtr pXL;
	WorkbooksPtr pBooks;
	
	// ����t�@�C���̃|�C���^
	_WorkbookPtr pBook;
	_WorksheetPtr pSheet;
	
	// �J�����t�@�C�����Ǘ�
	std::vector<eFileInfo> efinfo;
	// ���݊J���Ă������
	eFileInfo nowfinfo;


public:
	/* Excel�t�@�C���I�[�v�� */
	BOOL FileOpen(const TCHAR* const, const int = 0);
	/* Excel�t�@�C���I�[�v�� & ���Z�b�g */
	BOOL eFileOpen(const TCHAR* const, const int = 0);
	/* Excel�t�@�C�����㏑���ۑ� */
	BOOL FileSave(void);
	/* Excel�t�@�C���𖼑O�����ĕۑ� */
	BOOL FileSaveAs(const TCHAR* const);
	/* Cell�̒l���擾 */
	TSTR GetCellsValue(const ULONG&, const ULONG&);
	/* Cell�ɒl���������� */
	BOOL SetCellsValue(const ULONG&, const ULONG&, const TCHAR* const = _T(""));
	/* Cell�l�̃t�H���g�F�ύX */
	BOOL CellsFtColorChange(const ULONG& line, const ULONG& length, const int& color = 0);
	/* Cell�l�̃t�H���g�T�C�Y�ύX */
	BOOL CellsFtSizeChange(const ULONG& line, const ULONG& length, const int& size = 10);
	/* Cell�̔w�i�F�ύX */
	BOOL CellsBkColorChange(const ULONG& line, const ULONG& length, const int& color = Excel::xlNone);
	/* Cell�̃p�^�[���ύX */
	BOOL CellsPatternChange(const ULONG& line, const ULONG& length, const long& pattern = 777);
	/* Cell�̔w�i�F�擾 */
	long GetCellsColor(const ULONG& line, const ULONG& length);
	/* Cell�̃t�H���g�F�擾 */
	long GetCellsFontColor(const ULONG& line, const ULONG& length);
	/* �}�`����]������ */
	void ShapesRotation(const float& rad);
	/* �l�p�`�V�F�C�v�̃e�L�X�g�ҏW */
	void EditRectanglesCaption(const TCHAR* const str);
	/* �V�[�g���̃V�F�C�v�̍��W�����߂� */
	void CalcShapePos(float& x, float& y);
	
	/* �A�N�e�B�u�V�[�g���폜���� */
	BOOL ActvSheetClear(void);
	/* �A�N�e�B�u�V�[�g���̕ύX */
	void ActvSheetNameChange(const TCHAR* const NewName);
	/* Excel�t�@�C���̃V�[�g�����擾 */
	long GetSheetCount(void);
	/* �A�N�e�B�u�V�[�g�̃V�[�g�ԍ����擾 */
	long GetActvSheetIndex(void);
	/* ����V�[�g�̕ύX */
	BOOL ActvSheetChange(const long& StNo = 1);
	/* ���ݑ��삵�Ă���t�@�C���̃p�X���擾 */
	TSTR GetActvBookPath(void);
	/* ����t�@�C����ύX */
	BOOL SetActvBook(const TCHAR* const);
	
	/* �}�N�����s */
	BOOL MacroExecution(const TCHAR* const MacroName, const TCHAR* const Argument = NULL);
	
	/* �w��ꏊ�̃Z�����N���A */
	void CellsClear(const ULONG&, const ULONG&);
	/* ���ݑ��삵�Ă���t�@�C������� */
	void FileClose(void);

	/* OS�̖��O�ƃo�[�W�������擾���� */
	TSTR GetOSNameAndVersion(void);


private:
	// typedef
	typedef std::vector<eFileInfo>::iterator EINFOITR;
	
	/* ���ݑ��삵�Ă���t�@�C������� */
	void SetActvSheet(ULONG);
	/* �J�����t�@�C������V�[�g�����Z�b�g���� */
	void SetSheetName(eSheetInfo&);
	/* �����t�@�C���̏����폜���� */
	BOOL eInfoDelete(void);
	/* �t�@�C�������(�P��) */
	void eFileClose(void);
	/* Excel����|�C���^�S�Ă�j������ */
	void DestroyEXLPtr(void);

}; // End Class

#endif /* CCEXCEL_INCLUDE */
