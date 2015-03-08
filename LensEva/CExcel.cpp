#include "CExcel.h"


/*=========================================================================
Constructor / Destructor
=========================================================================*/
CCExcel::CCExcel(void)
{
	// COM������
	CoInitialize(NULL);
	// Excel�N��
	pXL.CreateInstance(L"Excel.Application");
	pBooks = pXL->Workbooks;

	//�����o������ 
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
		Excel�`���̃t�@�C�����J��
	[Argument1] path
		�t�@�C���p�X�ւ̃|�C���^
	[Argument2] disp
		�\���̗L��:1 / ����:0
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
		Excel�`���̃t�@�C�����J���A�X�Ƀt�@�C�������Z�b�g����
	[Argument1] path
		�t�@�C���p�X�ւ̃|�C���^
	[Argument2] disp
		�\���̗L��:1 / ����:0
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::eFileOpen(const TCHAR* const path, const int disp)
{
	if(!FileOpen(path,disp))
		return FALSE;

	{ // �t�@�C�����Z�b�g
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
		Excel�`���̃t�@�C�����㏑���ۑ�����
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
		Excel�`���̃t�@�C���𖼑O�����ĕۑ�����
	[Argument1] path
		WorkBook�p�X
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
		�w����W�̃Z���l�𕶎���Ŏ擾����
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument1] length
		��ʒu�ւ̎Q��
	[Return]
		Complete : �Z���̒l		Error : NULL
**************************************************************************/
TSTR CCExcel::GetCellsValue(const ULONG& line, const ULONG& length)
{
	TSTR dst = _T("");
	
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return dst;

	// �l���̂��̂��擾
	const TSTR str = (const _bstr_t)pRange->Value;

	{// �Z�������Ԍ`���������ꍇ�̒���
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
			
			// ���ԏ��������邱��
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
		�w����W�̃Z���̔w�i�F��ID�Ŏ擾����
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument1] length
		��ʒu�ւ̎Q��
	[Return]
		Complete : �F��ID		Error : 
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
		�w����W�̃Z���̃t�H���g�F��ID�Ŏ擾����
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument1] length
		��ʒu�ւ̎Q��
	[Return]
		Complete : �F��ID		Error : 
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
		�}�`����]������
	[Argument1] rad
		��]�p�x
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
		�l�p�`�V�F�C�v�̃e�L�X�g�ҏW
	[Argument1] str
		������
	[Remarks]
		��) �V�[�g��S�Ă̎l�p�`�ɑ΂��ēK�p����\��������̂Œ���
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
		�V�[�g���̃V�F�C�v�̍��W�����߂�
	[Argument1] x
		x���W�Z�o����
	[Argument2] y
		y���W�Z�o����
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
		�w����W�̃Z���̒l��ύX����
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument2] length
		��ʒu�ւ̎Q��
	[Argument3] setstr
		�ύX�l
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::SetCellsValue(const ULONG &line, const ULONG& length, const TCHAR* const setstr)
{
	if(!pXL)
		return FALSE;
	RangePtr pCells = pSheet->Cells->Item[line][length];
	//Excel2002�ȏ�̏ꍇ��Value2
	pCells->Value2 = setstr;
	
	return TRUE;
}


/**************************************************************************
	[Function]
		MacroExecution()
	[Details]
		�}�N�������s����
	[Argument1] MacroName
		�}�N����
	[Argument2] Argument=NULL
		�}�N���֓n������
	[Return]
		Complete : TRUE		Error : FALSE
	[Remarks]
		���݊J���Ă���Book�ɑ΂��ėL���ł�
		Excel�}�N�����֓n���������1�܂őΉ�  �w�肵�Ȃ��ꍇ�͈��������Ɣ��f����
**************************************************************************/
BOOL CCExcel::MacroExecution(const TCHAR* const MacroName, const TCHAR* const Argument)
{
	BOOL ret = FALSE;
	
	TSTR RunName = _T("");
	{// �}�N���֐����擾
		if(nowfinfo.fpinfo.file != _T(""))
			RunName = nowfinfo.fpinfo.file + _T('!') + MacroName;
		else
			RunName = MacroName;
	}
	
	HRESULT m_hr = E_FAIL;
	{// �}�N�����s
		if(!Argument)
			m_hr = pXL->Run(RunName.c_str());
		else
			m_hr = pXL->Run(RunName.c_str(),Argument);
	}

	switch(m_hr)
	{// ��������
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
		�w����W�̃Z���l�̃t�H���g�F��ύX
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument2] length
		��ʒu�ւ̎Q��
	[Argument3] color
		�ύX�F
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
BOOL CCExcel::CellsFtColorChange(const ULONG& line, const ULONG& length, const int& color)
{
	RangePtr pRange = NULL;
	if(!pXL || !(pRange = pSheet->GetCells()->GetItem(line,length)))
		return FALSE;
	
	//Cell�l�̃t�H���g�F�ύX
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
		�w����W�̃Z���̃t�H���g�T�C�Y��ύX
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument2] length
		��ʒu�ւ̎Q��
	[Argument3] size
		�ύX�T�C�Y
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
		�w����W�̃Z���̃p�^�[����ύX
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument2] length
		��ʒu�ւ̎Q��
	[Argument3] pattern
		�w��p�^�[��
		( ��> xlSolid: ��  xlGray16: 12.5%�Ԋ|��  xlGray25: 25%�Ԋ|�� )
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
		//Cell�̃p�^�[�����w��(Ver2007�ȊO)
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
		�w����W�̃Z�����w��F�œh��Ԃ�
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument1] length
		��ʒu�ւ̎Q��
	[Argument3] color
		�ύX�F
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
		�A�N�e�B�u�V�[�g���폜����
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
		�A�N�e�B�u�V�[�g���̕ύX
	[Argument1] NewName
		�ύX��
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
		�A�N�e�B�u�V�[�g�̕ύX
	[Argument1] StNo
		�A�N�e�B�u�ɂ������V�[�g�ԍ�
		(�J�������_�Ő擪�ɂȂ��Ă���V�[�g��Sheet1(1)�ɂȂ�܂�)
	[Return]
		Complete : TRUE		Error : FALSE
	[Remarks]
		0��n���ƃA�N�Z�X�G���[�ɂȂ�܂�
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
		Excel�t�@�C���̃V�[�g�����擾
	[Return]
		Complete : 1�ȏ�̃V�[�g��	Error : 0�ȉ��̐���	
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
		�A�N�e�B�u�V�[�g�̃V�[�g�ԍ����擾
	[Return]
		Complete : 1�ȏ�̃V�[�g�ԍ�	Error : 0�ȉ��̐���	
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
		���ݑ��삵�Ă���u�b�N�̃p�X���擾����
	[Return]
		Complete : �u�b�N�̃p�X��		Error : Empty
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
		����u�b�N��؂�ւ���
	[Remarks]
		��x�ǂ�(���Ă��Ȃ�)�u�b�N���w�肵�Ȃ��Ǝ��s����
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
		�w����W�̃Z���l,�p�^�[�����̑S�Ă��N���A
	[Argument1] line
		�s�ʒu�ւ̎Q��
	[Argument1] length
		��ʒu�ւ̎Q��
	[Return]
		Complete : TRUE		Error : FALSE
**************************************************************************/
void CCExcel::CellsClear(const ULONG& line, const ULONG& length)
{
	if(!pXL)
		return;
	RangePtr pCells = pSheet->Cells->Item[line][length];
	pCells->Value2 = _T("");

	// �t�H���g�f�t�H���g
	CellsFtColorChange(line, length);
	// �w�i�F�f�t�H���g
	CellsBkColorChange(line, length);
	// �p�^�[���f�t�H���g
	CellsPatternChange(line, length);
}


/**************************************************************************
	[Function]
		void FileClose()
	[Details]
		���݊J���Ă���t�@�C�������
	[Remarks]
		�S�Ă�Excel���삪���������Ƃ��̂ݎg�p���邱��
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
		OS�̖��O�ƃo�[�W�������擾����(Excel�d�l)
	[Return]
		�擾�����o�[�W�������
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
		�A�N�e�B�u�V�[�g��ύX����
	[Argument1] i
		�A�N�e�B�u�ɂ������V�[�g��
	[Return]
		���̃N���X�ȊO�̃A�N�Z�X�֎~
-------------------------------------------------------------------------*/
void CCExcel::SetActvSheet(ULONG ShtNo)
{		
	if(!pXL || !pBook || !ShtNo)
		return;
	
	//�A�N�e�B�u�E�V�[�g�̕ύX
	const _variant_t data = (short)ShtNo;
	static const SheetsPtr pSheets = pBook->Worksheets->Item[data];
	pSheets->Select(data);
	
}


/*-------------------------------------------------------------------------
	[Function]
		void SetSheetName()
	[Details]
		���ݑ��삵�Ă���u�b�N�̃V�[�g�����Z�b�g����
	[Argument1] lifo
		�V�[�g���\���̂̎Q��
	[Return]
		���̃N���X�ȊO�̃A�N�Z�X�֎~
-------------------------------------------------------------------------*/
void CCExcel::SetSheetName(eSheetInfo& info)
{		
	if(!pXL || !pBook)
		return;
	
	// ���݊J���Ă���V�[�g���L������
	const _WorksheetPtr defptr = pSheet;
	
	// �V�[�g���̎擾
	info.cnt = pXL->Worksheets->Count;
	// ���݂���V�[�g����S�ăZ�b�g
	for(long i=1; i<info.cnt+1; i++) {
		_variant_t data = (short)i;
		//�A�N�e�B�u�E�V�[�g�|�C���^�̎擾
		ActvSheetChange(i);
		info.pSheet.push_back(pSheet);
		//�V�[�g�����擾
		info.name.push_back(static_cast<TCHAR*>(pSheet->Name));
	}

	pSheet = defptr;
}


/*-------------------------------------------------------------------------
	[Function]
		BOOL eInfoDelete()
	[Details]
		���݊J���Ă���t�@�C���̏����폜����
	[Remarks]
		���̏ꏊ����P�Ƃł̎g�p���֎~���邽�߃v���C�x�[�g
	[Return]
		Complete  : TRUE   (�폜����)
		Error     : FALSE  (�폜���s�������͑Ώۂ�������Ȃ�)
-------------------------------------------------------------------------*/
BOOL CCExcel::eInfoDelete(void)
{
	BOOL ret = FALSE;
	// �J���Ă���t�@�C���̏����폜����
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
		���݊J���Ă���t�@�C�������
	[Remarks]
		���ݑ��쒆��Book��Sheet�ɑ΂��ėL��
-------------------------------------------------------------------------*/
void CCExcel::eFileClose(void)
{
	// Excel����劲�����o�̉��
	if(pBooks)
		pBooks.Release();
	
	// ����Book, Sheet ���ꎞ���
	if(pSheet)
		pSheet.Release();
	if(pBook)
		pBook.Release();
}


/*-------------------------------------------------------------------------
	[Function]
		void DestroyEXLPtr()
	[Details]
		Excel����|�C���^�S�Ă�j��
	[Remarks]
		�f�X�g���N�^��p
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



