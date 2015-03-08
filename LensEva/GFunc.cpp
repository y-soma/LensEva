#include "GFunc.h"


/**************************************************************************
	[Function]
		Split()
	[Details]
		指定した文字列をSplit(分割)する
	[Argument1] str
		ぶった切り対象文字列
	[Argument2] delim
		デリミタ
	[Return]
		成功 : 分割した配列   出来ない : 空文字
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
		フルパスからファイル名を取り出す
	[Argument1] fullPath
		フルパスへのポインタ
	[Return]
		成功 : ファイル名	失敗 : 空文字
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
		フルパスからディレクトリ名を取り出す
	[Argument1] fullPath
		フルパスへのポインタ
	[Return]
		成功 : ディレクトリ名	失敗 : 空文字
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
		(自分のいる)ルートディレクトリ情報を取得する
	[Argument1] el = 0
		取り出す要素
		( 0:ディレクトリ 1:フルパス 2:ドライブ 3:実行ファイル名 4:拡張子 )
	[Return]
		成功 : 自分(.exe)のいるディレクトリ	失敗 : Error
	[Remarks]
		引数指定しない場合はディレクトリが返る
		デフォルトの場合、最後の'\\'は取り除いているので、必要に応じて付け加える
**************************************************************************/
TSTR chl::GetAppMeDir(const int& el)
{
	TSTR ret = _T("Error");
	{ // 自分(実行ファイル)のいるディレクトリを取得
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
		数値を文字列に変換する
	[Argument1] val
		変換対象
	[Return]
		成功 : 文字列の変換対象	   失敗 : 空文字
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
		少数を文字列に変換する
	[Argument1] val
		変換対象
	[Return]
		成功 : 文字列の変換対象	   失敗 : 空文字
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
		文字列の数値を少数型に変換する
	[Argument1] str
		変換対象
	[Return]
		少数型の変換対象
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
		文字列の数値を整数型に変換する
	[Argument1] str
		変換対象
	[Return]
		整数型の変換対象
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
		整数を文字列の16進数に変換する
	[Argument1] val
		変換対象
	[Return]
		16進数へ変換した文字列
**************************************************************************/
TSTR chl::Dec2HexStr(const ULONG& val)
{
	wchar_t buf[MAX_PATH*0xFF];
	{// 算出
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
		文字列の数値を文字列の16進数に変換する
	[Argument1] val
		変換対象
	[Return]
		16進数へ変換した文字列
**************************************************************************/
TSTR chl::Dec2HexStr(const TSTR& val)
{
	return Dec2HexStr(StrToLng(val));
}


/**************************************************************************
	[Function]
		GetNumInStr()
	[Details]
		文字列中の数値だけを抽出する
	[Argument1] src
		取り出す対象
	[Return]
		成功 : 抽出連結した数値	  見つからない : 空文字
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
		Initファイル書き込みを行う
	[Argument1] IniName
		書き込むiniファイル名
	[Argument2] section
		指定セクション
	[Argument3] param
		指定パラメータ
	[Argument4] val
		書き込み値
	[Argument5] dir
		指定ディレクトリ
	[Return]
		成功 : TRUE		失敗 : FALSE
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
		Initファイルのパラメータを取得する
	[Argument1] IniName
		iniファイル名
	[Argument2] section
		指定セクション
	[Argument3] param
		指定パラメータ
	[Argument4] dir
		指定ディレクトリ
	[Return]
		成功 : 取得したパラメータ	失敗 : Error  取得できず : default
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
		vector配列内の重複を取り除く(汎用型)
	[Argument1] src
		対象配列
	[Return]
		重複除外後の配列
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
		GUI(エディットボックスなど)に改行を含めた文字列を追加できるように独自の形式に変換する
	[Argument1] src
		変換対象
	[Return]
		成功 : 変換結果	  改行が見つからない : そのまま
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
		GUI(エディットボックスなど)用に変換した改行リテラルが含まれた文字列を元の形式へ戻す
	[Argument1] src
		変換対象
	[Return]
		成功 : 変換結果	  改行が見つからない : そのまま
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
		User、PCの取得
	[Return]
		成功 : User,PC    失敗 : Empty
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
		ログファイル出力
	[Argument1] path
		ディレクトリのみ
	[Argument2] header
		ログに書き込む内容の先頭に付け加える文字列
	[Return]
		成功 : TRUE    失敗 : FALSE
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
			// ※ L"w,ccs=UNICODE" を指定しないとワイド文字列(漢字など)がファイルに正しく出力されない
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
		ディレクトリの存在確認(改良版)
		(STLであるPathIsDirectory関数は見つからないときに処理が激重になるので、独自に高速化を図った)
	[Argument1] path
		ディレクトリ
	[Return]
		発見 : TRUE    見つからない : FALSE
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
			// 開いたファイルは邪魔になるので削除
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
		ファイルの存在確認(改良版)
		(STLであるPathFileExists関数は見つからないときに処理が激重になるので、独自に高速化を図った)
	[Argument1] path
		ファイルへのフルパス
	[Return]
		発見 : TRUE    見つからない : FALSE
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
		指定ディレクトリの中から作成するファイル名の候補を検索して、見つからない場合のみそのファイル名を返す
	[Argument1] path
		ディレクトリのみ
	[Argument2] opfname
		作成ファイル名の先頭名
	[Return]
		成功 : 検索パス    失敗 : empty
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
		指定した文字列のUnicodeを文字列で取得する
	[Argument1] src
		指定文字列
	[Argument2] size
		文字列数
	[Return]
		成功 : カンマで繋いだUnicode化された文字列    失敗 : Empty
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
		指定した文字列のUnicodeをファイルへ出力する
	[Argument1] src
		指定文字列
	[Argument2] size
		文字列数
	[Return]
		成功 : TRUE    失敗 : FALSE
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
		'\\'の部分を別の文字列へ変換する
	[Argument1] src
		変換対象
	[Argument2] val
		変換したい値(文字列)
	[Return]
		成功 : 変換後の文字列    失敗 : empty
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
		OSバージョンを取得する
	[Return]
		成功 : 取得したバージョン    失敗 : "OS Unknown"
**************************************************************************/
TSTR chl::GetOSVersion(void)
{
	OSVERSIONINFOEX osvi;
	{// システム情報を取得する
		osvi.dwOSVersionInfoSize = sizeof(osvi);
		GetVersionEx((OSVERSIONINFO *)&osvi);
	}

	TSTR dst = _T("");
	switch(osvi.dwPlatformId)
	{// 取得したシステム情報を振り分け
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





