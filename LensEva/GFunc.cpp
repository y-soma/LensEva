#include "GFunc.h"


/**************************************************************************
	[Function]
		Split()
	[Details]
		�w�肵���������Split(����)����
	[Argument1] str
		�Ԃ����؂�Ώە�����
	[Argument2] delim
		�f���~�^
	[Return]
		���� : ���������z��   �o���Ȃ� : �󕶎�
**************************************************************************/
std::vector<TSTR> chl::Split(const TCHAR* const str, const TCHAR* const delim)
{	
	std::vector<TSTR> items;
	TSTR src = str;
	std::size_t dlm_idx;
	if(src.npos == (dlm_idx = src.find_first_of(delim))) {
		items.push_back(src.substr(0, dlm_idx));
	}
	while(src.npos != (dlm_idx = src.find_first_of(delim))) {
		if(src.npos == src.find_first_not_of(delim)) {
			break;
		}
		items.push_back(src.substr(0, dlm_idx));
		dlm_idx++;
		src = src.erase(0, dlm_idx);
		if(src.npos == src.find_first_of(delim) && _T("") != src.data()) {
			items.push_back(src);
			break;
		}
	}
	return items;
}


/**************************************************************************
	[Function]
		GetFileName()
	[Details]
		�t���p�X����t�@�C���������o��
	[Argument1] fullPath
		�t���p�X�ւ̃|�C���^
	[Return]
		���� : �t�@�C����	���s : �󕶎�
**************************************************************************/
TSTR chl::GetFileName(const TCHAR* const fullPath)
{ 
	const TSTR fpath = fullPath;
	TSTR FileName = _T("");
	int find = 0;	
	if ((find = (int)fpath.rfind('\\')) > 0 ){ 
		for (UINT i = find+1; i < fpath.length() ; i++) {
			FileName += fpath.substr(i,1);
		}
	}
	else {
		return NULL;
	}

	return FileName;
}


/**************************************************************************
	[Function]
		GetDirName()
	[Details]
		�t���p�X����f�B���N�g���������o��
	[Argument1] fullPath
		�t���p�X�ւ̃|�C���^
	[Return]
		���� : �f�B���N�g����	���s : �󕶎�
**************************************************************************/
TSTR chl::GetDirName(const TCHAR* const fullPath)
{ 
	const TSTR fpath = fullPath;
	TSTR DirName = _T("");
	int find = 0;	
	if ((find = (int)fpath.rfind('\\')) > 0 ){ 
		for (int i = 0; i < find ; i++) {
			DirName += fpath.substr(i,1);
		}
	}
	else {
		return NULL;
	}

	return DirName;
}


/**************************************************************************
	[Function]
		GetAppMeDir()
	[Details]
		(�����̂���)���[�g�f�B���N�g�������擾����
	[Argument1] el = 0
		���o���v�f
		( 0:�f�B���N�g�� 1:�t���p�X 2:�h���C�u 3:���s�t�@�C���� 4:�g���q )
	[Return]
		���� : ����(.exe)�̂���f�B���N�g��	���s : Error
	[Remarks]
		�����w�肵�Ȃ��ꍇ�̓f�B���N�g�����Ԃ�
		�f�t�H���g�̏ꍇ�A�Ō��'\\'�͎�菜���Ă���̂ŁA�K�v�ɉ����ĕt��������
**************************************************************************/
TSTR chl::GetAppMeDir(const int& el)
{
	TSTR ret = _T("Error");
	{ // ����(���s�t�@�C��)�̂���f�B���N�g�����擾
		TCHAR path[MAX_PATH], drive[MAX_PATH], dir[MAX_PATH], fname[MAX_PATH], ext[MAX_PATH];
		const DWORD gfret = GetModuleFileName(NULL, path, MAX_PATH);
		const int spret= _wsplitpath_s(path,drive,dir,fname,ext); // Return(Comp:0  Error:EINVAL)
		if(!gfret || spret)
			return ret;

		switch(el)
		{
			case 0:
				ret = GetDirName(path);
				//ret = dir;
				break;
			case 1:
				ret = path;
				break;
			case 2:
				ret = drive;
				break;
			case 3:
				ret = fname;
				break;
			case 4:
				ret = ext;
				break;
			//default:
		}
	}

	return ret;
}


/**************************************************************************
	[Function]
		LngToStr()
	[Details]
		���l�𕶎���ɕϊ�����
	[Argument1] val
		�ϊ��Ώ�
	[Return]
		���� : ������̕ϊ��Ώ�	   ���s : �󕶎�
**************************************************************************/
TSTR chl::LngToStr(const long& val)
{ 
	TSTR dst = _T("");
	TCHAR buf[MAX_PATH] = {0};
	wsprintf(buf, _T("%d"), val);
	dst = buf;

	return dst;
}


/**************************************************************************
	[Function]
		DblToStr()
	[Details]
		�����𕶎���ɕϊ�����
	[Argument1] val
		�ϊ��Ώ�
	[Return]
		���� : ������̕ϊ��Ώ�	   ���s : �󕶎�
**************************************************************************/
TSTR chl::DblToStr(const double& val)
{ 
	TSTR dst = _T("");
	TCHAR buf[MAX_PATH] = {0};
	_stprintf_s(buf, _T("%f"), val);
	//wsprintf(buf, _T("%f"), val);
	dst = buf;
	return dst;
}


/**************************************************************************
	[Function]
		StrToLng()
	[Details]
		������̐��l�������^�ɕϊ�����
	[Argument1] str
		�ϊ��Ώ�
	[Return]
		�����^�̕ϊ��Ώ�
**************************************************************************/
double chl::StrToDbl(const TSTR& str)
{ 
	const double dst = _wtof(str.data());
	return dst;
}


/**************************************************************************
	[Function]
		StrToLng()
	[Details]
		������̐��l�𐮐��^�ɕϊ�����
	[Argument1] str
		�ϊ��Ώ�
	[Return]
		�����^�̕ϊ��Ώ�
**************************************************************************/
long chl::StrToLng(const TSTR& str)
{ 
	const long dst = _wtol(str.data());
	return dst;
}


/**************************************************************************
	[Function]
		Dec2HexStr()
	[Details]
		�����𕶎����16�i���ɕϊ�����
	[Argument1] val
		�ϊ��Ώ�
	[Return]
		16�i���֕ϊ�����������
**************************************************************************/
TSTR chl::Dec2HexStr(const ULONG& val)
{
	wchar_t buf[MAX_PATH*0xFF];
	{// �Z�o
		wchar_t c = 0;
		ULONG k = 1, n = 0;
		while((k<<=4) <=val)
			n++;
		n++;
		for(ULONG i=0;i<n;i++){
			c = (wchar_t)(0xf & (val>>(i*4)));
			buf[n-i-1] = (wchar_t)((c>9)?(c+L'A'-10):(c+L'0'));
		}
		buf[n] = L'\0';
	}

	const TSTR dst = buf;
	return dst;
}


/**************************************************************************
	[Function]
		Dec2HexStr()
	[Details]
		������̐��l�𕶎����16�i���ɕϊ�����
	[Argument1] val
		�ϊ��Ώ�
	[Return]
		16�i���֕ϊ�����������
**************************************************************************/
TSTR chl::Dec2HexStr(const TSTR& val)
{
	return Dec2HexStr(StrToLng(val));
}


/**************************************************************************
	[Function]
		GetNumInStr()
	[Details]
		�����񒆂̐��l�����𒊏o����
	[Argument1] src
		���o���Ώ�
	[Return]
		���� : ���o�A���������l	  ������Ȃ� : �󕶎�
**************************************************************************/
TSTR chl::GetNumInStr(const TSTR& src)
{
	TSTR dest = _T("");
	for(UINT i=0; i<src.length(); i++){
		for(long j=0; j<10; j++){
			const TSTR bufstr = LngToStr(j);
			if(bufstr == src.substr(i,1))
				dest += bufstr;
		}
	}

	return dest;
}


/**************************************************************************
	[Function]
		BOOL WriteInitFile()
	[Details]
		Init�t�@�C���������݂��s��
	[Argument1] IniName
		��������ini�t�@�C����
	[Argument2] section
		�w��Z�N�V����
	[Argument3] param
		�w��p�����[�^
	[Argument4] val
		�������ݒl
	[Argument5] dir
		�w��f�B���N�g��
	[Return]
		���� : TRUE		���s : FALSE
**************************************************************************/
BOOL chl::WriteInitFile(const TCHAR* const IniName, const TCHAR* const section, const TCHAR* const param, const TCHAR* const val, const TCHAR* const dir)
{
	BOOL ret = FALSE;
	TSTR rdir = _T("");
	if(!dir){
		if( (rdir = chl::GetAppMeDir()) == _T("Error") ){
			return ret;
		}
	}
	else
		rdir = dir;

	const TSTR setdir = rdir + _T('\\') + IniName;
	ret = WritePrivateProfileString(section, param, val, setdir.data());
	
	return ret;
}


/**************************************************************************
	[Function]
		const CString GetInitParam()
	[Details]
		Init�t�@�C���̃p�����[�^���擾����
	[Argument1] IniName
		ini�t�@�C����
	[Argument2] section
		�w��Z�N�V����
	[Argument3] param
		�w��p�����[�^
	[Argument4] dir
		�w��f�B���N�g��
	[Return]
		���� : �擾�����p�����[�^	���s : Error  �擾�ł��� : default
**************************************************************************/
TSTR chl::GetInitParam(const TCHAR* const IniName, const TCHAR* const section, const TCHAR* const param, const TCHAR* const dir)
{
	TSTR ret = _T("Error");

	TSTR rdir = _T("");
	if(!dir){
		if( (rdir = chl::GetAppMeDir()) == _T("Error") ){
			return ret;
		}
	}
	else
		rdir = dir;

	const TSTR setdir = rdir + _T('\\') + IniName;

	TCHAR getval[_MAX_PATH] = {0};
	const DWORD getcnt = GetPrivateProfileString(section, param, NOT_FOUND, getval, _MAX_PATH, setdir.data());
	ret = getval;

	return ret;
}



/**************************************************************************
	[Function]
		GetArryEqualVal()
	[Details]
		vector�z����̏d������菜��(�ėp�^)
	[Argument1] src
		�Ώ۔z��
	[Return]
		�d�����O��̔z��
**************************************************************************/
template<typename TOBJ> std::vector<TOBJ> chl::GetArryEqualVal(const std::vector<TOBJ>& src)
{
	TOBJ* p = new(std::nothrow) TOBJ;
	std::vector<TOBJ> dst(0,*p);
	
	TOBJ val = *p;
	ULONG found = 0;
	for(ULONG i=0; i<src.size(); i++){
		if(val != src[i]){
			ULONG jfound = 0;
			for(ULONG j=i+1; j<src.size(); j++){
				if(src[i] == src[j])
					jfound++;
			}
			if(!jfound){
				val = src[i];
				dst.push_back(src[i]);
			}
		}
		else{
			found++;
		}
	}
	delete p;
	
	if(found)
		GetArryEqualVal(dst);

	return dst;
}


/**************************************************************************
	[Function]
		CrLfConvGUIEdit()
	[Details]
		GUI(�G�f�B�b�g�{�b�N�X�Ȃ�)�ɉ��s���܂߂��������ǉ��ł���悤�ɓƎ��̌`���ɕϊ�����
	[Argument1] src
		�ϊ��Ώ�
	[Return]
		���� : �ϊ�����	  ���s��������Ȃ� : ���̂܂�
**************************************************************************/
TSTR chl::CrLfConvGUIEdit(const TSTR& src)
{
	TSTR dst = _T("");

	wchar_t* buf;
	const int bufsize = (int)src.length()+1;
	buf = new wchar_t[bufsize];
	wcscpy_s(buf, bufsize, src.data());
	int find = 0;
	for(int i=0; i<bufsize; i++){
		if(buf[i] == 10 || buf[i] == 13 || buf[i] == 218){
			find++;
			buf[i] = CRLF_AGENCY_LITERAL;
		}
	}
	if(!find)
		dst = src;
	else
		dst = (TCHAR*)buf;
	
	delete []buf;
	
	return dst;
}


/**************************************************************************
	[Function]
		CrLfReturnGUIEdit()
	[Details]
		GUI(�G�f�B�b�g�{�b�N�X�Ȃ�)�p�ɕϊ��������s���e�������܂܂ꂽ����������̌`���֖߂�
	[Argument1] src
		�ϊ��Ώ�
	[Return]
		���� : �ϊ�����	  ���s��������Ȃ� : ���̂܂�
**************************************************************************/
TSTR chl::CrLfReturnGUIEdit(const TSTR& src)
{
	TSTR dst = _T("");

	wchar_t* buf;
	const int bufsize = (int)src.length()+1;
	buf = new wchar_t[bufsize];
	wcscpy_s(buf, bufsize, src.data());
	int find = 0;
	for(int i=0; i<bufsize; i++){
		if(buf[i] == CRLF_AGENCY_LITERAL){
			find++;
			buf[i] = 10;
		}
	}
	if(!find)
		dst = src;
	else
		dst = (TCHAR*)buf;
	
	delete []buf;
	
	return dst;
}


/**************************************************************************
	[Function]
		GetPCAndUserName()
	[Details]
		User�APC�̎擾
	[Return]
		���� : User,PC    ���s : Empty
**************************************************************************/
TSTR chl::GetPCAndUserName(void)
{
	TSTR dst = _T("");
	
	TSTR username = _T("");
	{//Get User
		wchar_t usertemp[MAX_PATH];
		DWORD dwLeng = MAX_PATH;
		if(!GetUserName(usertemp,&dwLeng))
			return dst;
		username = usertemp;
		
		{// if number user 
			const long numu = StrToLng(username);
			if(numu != 0L && numu > 5500000)
				username = _T("0x") + Dec2HexStr(numu-5500000);
		}
	}
	TSTR pcname = _T("");
	{//Get PC
		wchar_t pctemp[MAX_PATH];
		DWORD dwLeng = MAX_PATH;
		if(!GetComputerName(pctemp,&dwLeng))
			return dst;
		pcname = pctemp;
	}
	dst = username + _T(',') + pcname;

	return dst;
}


/**************************************************************************
	[Function]
		PutLogFile()
	[Details]
		���O�t�@�C���o��
	[Argument1] path
		�f�B���N�g���̂�
	[Argument2] header
		���O�ɏ������ޓ��e�̐擪�ɕt�������镶����
	[Return]
		���� : TRUE    ���s : FALSE
**************************************************************************/
BOOL chl::PutLogFile(const TCHAR* const path, const TCHAR* const header)
{
	BOOL ret =FALSE;
	if(!path)
		return ret;
	
	try
	{
		const TSTR Path = path;
		// if not found
		if(!PathIsDirectoryEX(Path.data()))
			return ret;
		
		// get make file path
		const TSTR filePath = GetMakeFilePath(Path.data());
	
		{// file put
			FILE* fp;
			// �� L"w,ccs=UNICODE" ���w�肵�Ȃ��ƃ��C�h������(�����Ȃ�)���t�@�C���ɐ������o�͂���Ȃ�
			if ((fp = _wfopen(filePath.data(), L"w,ccs=UNICODE")) == NULL) {
				return ret;
			}

			TSTR Log = _T("");
			/*
			while(1)
			{
				TSTR rstr = _T("");
				if(!stdFile.ReadString(rstr))
					break;
				//stdFile.SeekToBegin();
				Log += (rstr+'\n');
			}
			*/
			if(header){
				const TSTR headtmp = header;
				Log += (_T("[INFO]\n")+GetPCAndUserName()+L'\n'+headtmp);
			}
			else
				Log += (_T("[INFO]\n")+GetPCAndUserName()+L'\n');
			
			//setlocale(LC_CTYPE, "");
			_fputts((wchar_t*)Log.data(), fp);
			fclose(fp);
		}

		ret = TRUE;
	}
	catch(...)
	{
		ret = FALSE;
	}
	
	return ret;
}



/**************************************************************************
	[Function]
		PathIsDirectoryEX()
	[Details]
		�f�B���N�g���̑��݊m�F(���ǔ�)
		(STL�ł���PathIsDirectory�֐��͌�����Ȃ��Ƃ��ɏ��������d�ɂȂ�̂ŁA�Ǝ��ɍ�������}����)
	[Argument1] path
		�f�B���N�g��
	[Return]
		���� : TRUE    ������Ȃ� : FALSE
**************************************************************************/
BOOL chl::PathIsDirectoryEX(const TCHAR* const path)
{
	BOOL ret = FALSE;
	{// proc
		// get make file path
		const TSTR filePath = GetMakeFilePath(path);
		FILE* fp = NULL;
		if ((fp = _wfopen(filePath.data(), L"w")) != NULL){
			fclose(fp);
			ret = TRUE;
			// �J�����t�@�C���͎ז��ɂȂ�̂ō폜
			if(!DeleteFile(filePath.data()))
				ret = FALSE;
		}
		else
			ret = FALSE;
	}

	return ret;
}


/**************************************************************************
	[Function]
		PathFileExistsEX()
	[Details]
		�t�@�C���̑��݊m�F(���ǔ�)
		(STL�ł���PathFileExists�֐��͌�����Ȃ��Ƃ��ɏ��������d�ɂȂ�̂ŁA�Ǝ��ɍ�������}����)
	[Argument1] path
		�t�@�C���ւ̃t���p�X
	[Return]
		���� : TRUE    ������Ȃ� : FALSE
**************************************************************************/
BOOL chl::PathFileExistsEX(const TCHAR* const path)
{
	BOOL ret = FALSE;
	{// proc
		// get make file path
		FILE* fp = NULL;
		if ((fp = _wfopen(path, L"r")) != NULL){
			fclose(fp);
			ret = TRUE;
		}
		else
			ret = FALSE;
	}

	return ret;
}



/**************************************************************************
	[Function]
		GetMakeFilePath()
	[Details]
		�w��f�B���N�g���̒�����쐬����t�@�C�����̌����������āA������Ȃ��ꍇ�݂̂��̃t�@�C������Ԃ�
	[Argument1] path
		�f�B���N�g���̂�
	[Argument2] opfname
		�쐬�t�@�C�����̐擪��
	[Return]
		���� : �����p�X    ���s : empty
**************************************************************************/
TSTR chl::GetMakeFilePath(const TCHAR* const path, const TCHAR* const opfname)
{
	TSTR dst = _T("");
	{// proc
		const TSTR Path = path;
		const TSTR OpFName = opfname;
		long i  = 1;
		while(1)
		{// filepath fix loop
			const TSTR filePathTemp = Path + _T("\\") + opfname + LngToStr(i) + _T(".txt");
			// if not file
			if(PathFileExists(filePathTemp.data())){
				i++;
				continue;
			}
			else {
				dst = filePathTemp;
				break;
			}
			// loop break point
			if(i > 0x1FF)
				return dst;
		}
	}

	return dst;
}


/**************************************************************************
	[Function]
		GetStrToUnicode()
	[Details]
		�w�肵���������Unicode�𕶎���Ŏ擾����
	[Argument1] src
		�w�蕶����
	[Argument2] size
		������
	[Return]
		���� : �J���}�Ōq����Unicode�����ꂽ������    ���s : Empty
**************************************************************************/
TSTR chl::GetStrToUnicode(const TCHAR* const src, const ULONG& size)
{
	TSTR dst = _T("");
	{// get code
		const wchar_t* str = (const wchar_t*)src;
		for(ULONG i=0; i<size; i++){
			TSTR tmp = Dec2HexStr((ULONG)str[i]);
			tmp.replace(1, _T('\0'), _T(""));
			dst = i!=size-1? dst+_T("0x")+tmp+_T(","):dst+_T("0x")+tmp;
		}
	}
	
	return dst;
}


/**************************************************************************
	[Function]
		PutStrToCode()
	[Details]
		�w�肵���������Unicode���t�@�C���֏o�͂���
	[Argument1] src
		�w�蕶����
	[Argument2] size
		������
	[Return]
		���� : TRUE    ���s : FALSE
**************************************************************************/
BOOL chl::PutStrToCode(const TCHAR* const src, const ULONG& size)
{
	const wchar_t* str = (const wchar_t*)src;
	try
	{
		FILE* fp;
		{// put file open
			TSTR putdir = _T("");
			{// filepath fix loop
				long i  = 1;
				while(1)
				{
					putdir = GetAppMeDir() + _T("\\StrConvCode") + LngToStr(i) + _T(".txt");
					// if not file
					if(PathFileExistsEX(putdir.data())){
						i++;
						continue;
					}
					else{
						break;
					}
					
					// loop break point
					if(i > 10000)
						break;
				}
			}

			if((fp = _wfopen(putdir.data(), L"w,ccs=UNICODE")) == NULL)
				return FALSE;
		}

		TSTR buf = src;
		buf += _T("\n");
		{// get code
			for(ULONG i=0; i<size; i++){
				TSTR tmp = Dec2HexStr((ULONG)str[i]);
				tmp.replace(1, _T('\0'), _T(""));
				buf += (_T("0x")+tmp+_T(","));
			}
		}

		_fputts((wchar_t*)buf.data(), fp);
		fclose(fp);
		
		return TRUE;
	}
	catch(...)
	{
		return FALSE;
	}
}


/**************************************************************************
	[Function]
		PathEscSeqConv()
	[Details]
		'\\'�̕�����ʂ̕�����֕ϊ�����
	[Argument1] src
		�ϊ��Ώ�
	[Argument2] val
		�ϊ��������l(������)
	[Return]
		���� : �ϊ���̕�����    ���s : empty
**************************************************************************/
TSTR chl::PathEscSeqConv(const TCHAR* const src, const TCHAR* const val)
{
	TSTR dst = _T("");
	{// proc
		const std::vector<TSTR> srcspt = Split(src, _T("\\"));
		for(ULONG i=0; i<srcspt.size(); i++){
			dst = i!=srcspt.size()-1? dst+(srcspt[i]+val):dst+srcspt[i];
		}
	}

	return dst;
}


/**************************************************************************
	[Function]
		GetOSVersion()
	[Details]
		OS�o�[�W�������擾����
	[Return]
		���� : �擾�����o�[�W����    ���s : "OS Unknown"
**************************************************************************/
TSTR chl::GetOSVersion(void)
{
	OSVERSIONINFOEX osvi;
	{// �V�X�e�������擾����
		osvi.dwOSVersionInfoSize = sizeof(osvi);
		GetVersionEx((OSVERSIONINFO *)&osvi);
	}

	TSTR dst = _T("");
	switch(osvi.dwPlatformId)
	{// �擾�����V�X�e������U�蕪��
		case VER_PLATFORM_WIN32s:
			dst = _T("Windows 3.1 (Win32s)");
			break;

		case VER_PLATFORM_WIN32_WINDOWS:
			switch(osvi.dwMajorVersion)
			{
				case 0:
					dst = _T("Windows 95");
					break;

				case 10:
					dst = _T("Windows 98");
					break;

				case 90:
					dst = _T("Windows Me");
					break;

				default:
					dst = _T("Windows 95/98/Me");
					break;
			}
			break;

		case VER_PLATFORM_WIN32_NT:
			switch(osvi.dwMajorVersion)
			{
				case 4:
					dst = _T("Windows NT4.0");
					break;

				case 5:
					switch(osvi.dwMinorVersion)
					{
						case 0:
							dst = _T("Windows 2000");
							break;

						case 1:
							dst = _T("Windows XP");
							break;

						case 2:
							dst = _T("Windows Server2003");
							break;
					}
					break;
				
				case 6:
					if(osvi.wProductType == VER_NT_WORKSTATION)
					{// Vista / 7
						switch(osvi.dwMinorVersion)
						{
							case 0:
								dst = _T("Windows Vista");
								break;

							case 1:
								dst = _T("Windows 7");
								break;

							default:
								dst = _T("Windows NT 6.x(unknown)");
								break;
						}
					}
					else
					{// Server 2008 / R2
						switch(osvi.dwMinorVersion)
						{
							case 0:
								dst = _T("Windows Server 2008");
								break;

							case 1:
								dst = _T("Windows Server 2008 R2");
								break;

							default:
								dst = _T("Windows Server 6.x(unknown)");
								break;
						}
					}
					break;

				default:
					dst = _T("Windows NT/2000/XP");
					break;
			}

			break;
		default:
			dst = _T("OS Unknown");
			break;
	}

	return dst;
}





