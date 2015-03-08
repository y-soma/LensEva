#include "LProc.h"
#include "CExcel.h"

#include <math.h>



CLProc::CLProc(void)
{
	// ���P�[������{��ɐݒ�
	_tsetlocale(LC_ALL, _T("Japanese"));
	Resdt.clear();
	OutDir.empty();
}

CLProc::~CLProc(void)
{
	Resdt.clear();
	OutDir.empty();
}



/* Public Member Function */

/**************************************************************************
	[Function]
		Execution
	[Details]
		�������s
	[Argument1] src
		��荞�݃t�@�C�����
	[Argument2] body
		�{�f�B��
	[Return]
		���� : 1	���s : 0
**************************************************************************/
int CLProc::Execution(const std::vector<chl::FilePInfo>& src, const TSTR& body, const TSTR& outdir)
{
	try
	{
		int ret = 0;
		{
			if(!(ret = ReadFile(src)))
				return ret;
			
			CCExcel ep;
			{// file open
				const TSTR fname = chl::GetAppMeDir() + _T("\\") + TEMPLATE_FILE;
				if(!ep.FileOpen(fname.data())){
					TSTR mes = _T("'"); mes += TEMPLATE_FILE; mes += _T("'���J���܂���ł���");
					MessageBox(NULL, mes.data(), _T("Error"), MB_OKCANCEL | MB_ICONERROR);
					return ret;
				}
			}
			{// �}�N�����s
				TSTR MacroName = TEMPLATE_FILE;
				MacroName += _T("!RUN");
				const TSTR MacroArg = MakeSummaryData(body,outdir);
				// �e�X�g�f�[�^
				//const TCHAR* testdt = _T("������Y,125,45,78/��r�����Y,125,45,78/��r�����Y�̎����l,125,45,78\\�F���x����1,12,25,36,25,56/�F���x����2,12,25,136,25,56/�F���x����3,12,25,36,25,56/�F���x����4,12,25,36,25,56/�F���x����5,12,25,36,25,56/�F���x����6,12,25,36,25,56\\�F���x����1,0.77,1.22,1.23,0.98/�F���x����2,0.77,1.22,1.23,0.98/�F���x����3,0.77,1.22,1.23,0.98/�F���x����3,0.77,1.22,1.23,0.98/�F���x����3,0.77,1.22,1.23,0.98\\E-PL1_PP2_049");
				if(!(ret = ep.MacroExecution(MacroName.data(),MacroArg.data()))){
					TSTR mes = _T("'"); mes += TEMPLATE_FILE; mes += _T("'�̃}�N�������s�ł��܂���ł���");
					MessageBox(NULL, mes.data(), _T("Error"), MB_OKCANCEL | MB_ICONERROR);
					ep.FileClose();
					return ret;
				}
			}
			// ����
			ep.FileClose();
		}
		return ret;
	}
	catch(...)
	{
		return 0;
	}
}


/* Private Member Function */

/*-------------------------------------------------------------------------
	[Function]
		ReadFile
	[Details]
		�t�@�C���ǂݍ���
	[Argument1] src
		��荞�݃t�@�C�����
	[Return]
		���� : 1	���s : 0
-------------------------------------------------------------------------*/
int CLProc::ReadFile(const std::vector<chl::FilePInfo>& src)
{
	try
	{
		std::vector<TSTR> GetCsv(0,_T("")), GetTxt(0,_T(""));
		BOOL csvflg = FALSE;
		for(ULONG i=0; i<src.size(); i++)
		{// �g���q�ɉ����ď�������
			const WCHAR* extW = PathFindExtensionW(src[i].full.data());
			if(!extW)
				return 0;
			
			if(!lstrcmpiW(extW, L".csv"))
			{// csv
				ReadCsv(src[i].full);
				if(OutDir == _T(""))
					OutDir = src[i].dir;
				if(!csvflg)
					csvflg = TRUE;
			}
			else if(!lstrcmpiW(extW, L".txt"))
			{// txt
				if(sp[0].LensName == _T("") || sp[1].LensName == _T("") || sp[2].LensName == _T(""))
					ReadTxt(src[i].full);
			}
		}

		if(!csvflg){
			MessageBox(NULL, _T("CSV�t�@�C�����I������Ă��܂���"), _T("Error"), MB_OKCANCEL | MB_ICONERROR);
			return 0;
		}

		if(sp[0].LensName == _T("") || sp[1].LensName == _T("") || sp[2].LensName == _T(""))
		{// ���̎��_�ŕ������ߗ��f�[�^�������ĂȂ�������f�t�H���g��ǂݍ���
			const TSTR spdef = chl::GetAppMeDir() + _T("\\") + SPECTRAL_DEF;
			ReadTxt(spdef);
		}
		return 1;
	}
	catch(...)
	{
		MessageBox(NULL, _T("�t�@�C���ǂݍ��݂Ɏ��s���܂���"), _T("Error"), MB_OKCANCEL | MB_ICONERROR);
		return 0;
	}
}


/*-------------------------------------------------------------------------
	[Function]
		ReadCsv
	[Details]
		csv�t�@�C��1����F���f�[�^���擾����
	[Argument1] path
		�p�X
-------------------------------------------------------------------------*/
void CLProc::ReadCsv(const TSTR& path)
{
	try
	{
		std::vector<TSTR> lines(0,_T(""));
		{// �t�@�C������S�Ă̍s�f�[�^�擾
			lines= GetLines(path);
			if(lines.size() < 143)
				return;
		}

		LDT buf;
		UINT cnt = 0;
		for(ULONG i=0; i<lines.size(); i++)
		{// �X�Ɏ擾�����f�[�^���d����
			if(cnt > 141)
				break;
			const std::vector<TSTR> line = chl::Split(lines[i].data(), _T(","));
			if(line.size() < 15)
				continue;
			switch(i)
			{
				// ���l�[�����擾���A��������������𔲂��o��
				case 0:
				{
					buf.StdDt.RName = line[4];
					buf.EvaDt.RName = line[11];
					{// ������
						const TSTR SLSrc = PickupLightSrcName(line[7]);
						const TSTR ELSrc = PickupLightSrcName(line[14]);
						buf.LightSrc = SLSrc;
						if(SLSrc == _T("") && ELSrc == _T(""))
							buf.LightSrc = chl::GetFileName(path.data());
						else if(SLSrc == _T("") && ELSrc != _T(""))
							buf.LightSrc = ELSrc;
					}
					break;
				}
				// �X���[
				case 1:
					break;
				// white balance
				case 76:
				{
					const double S_R = chl::StrToDbl(line[7]);
					const double S_G = chl::StrToDbl(line[8]);
					const double S_B = chl::StrToDbl(line[9]);
					const double E_R = chl::StrToDbl(line[14]);
					const double E_G = chl::StrToDbl(line[15]);
					const double E_B = chl::StrToDbl(line[16]);
					if(S_R >= 0.0 && S_G >= 0.0 && S_B >= 0.0){
						buf.StdDt.N5_R = S_R;
						buf.StdDt.N5_G = S_G;
						buf.StdDt.N5_B = S_B;
					}
					if(E_R >= 0.0 && E_G >= 0.0 && E_B >= 0.0){
						buf.EvaDt.N5_R = E_R;
						buf.EvaDt.N5_G = E_G;
						buf.EvaDt.N5_B = E_B;
					}
					// break;  //���F�������߂����̂ŁA������break������default�֗���
				}
				// �F��
				default:
					buf.Dedt.push_back(CalcDE(line[10],line[11],line[12],line[17],line[18],line[19]));
					break;
			}
			cnt++;
		}
		Resdt.push_back(buf);
	}
	catch(...)
	{
		return;
	}
}


/*-------------------------------------------------------------------------
	[Function]
		ReadTxt
	[Details]
		txt�t�@�C��1���番�����ߗ��f�[�^���擾����
	[Argument1] path
		�p�X
-------------------------------------------------------------------------*/
void CLProc::ReadTxt(const TSTR& path)
{
	try
	{
		std::vector<TSTR> lines(0,_T(""));
		{// �t�@�C������S�Ă̍s�f�[�^�擾
			lines= GetLines(path);
			if(lines.size() < 66)
				return;
		}
		
		for(ULONG i=0; i<lines.size(); i++)
		{// �X�Ɏ擾�����f�[�^���d����
			const std::vector<TSTR> line = chl::Split(lines[i].data(), _T("\t"));
			if(line.size() < 3)
				continue;
			
			if(!i)
			{// �����Y��
				sp[0].LensName = line[0];
				sp[1].LensName = line[1];
				sp[2].LensName = line[2];
			}
			else
			{// �g�����Ƃ̃f�[�^
				sp[0].Wdt.push_back(line[0]);
				sp[1].Wdt.push_back(line[1]);
				sp[2].Wdt.push_back(line[2]);
			}
		}
	}
	catch(...)
	{
		return;
	}
}


/*-------------------------------------------------------------------------
	[Function]
		GetLines
	[Details]
		�t�@�C����ǂݍ��ݑS�s�f�[�^���擾����
	[Argument1] path
		�p�X
	[Return]
		���� �F �s���Ƃ̔z��f�[�^	�@���s �F ��z��
-------------------------------------------------------------------------*/
std::vector<TSTR> CLProc::GetLines(const TSTR& path)
{
	std::vector<TSTR> lines(0,_T(""));
	{// �t�@�C������f�[�^��S�ēǂݍ���
		FILE* fp;
		if(!(fp = _wfopen(path.data(), L"r,ccs=UNICODE")))
			return lines;
		while(1)
		{// �S�Ă̍s��ǂݍ���
			TSTR add = _T("");
			{// ���s����菜����1�s���擾
				TCHAR buf[MAX_PATH*0xff] = {0};
				if(!_fgetts(buf,sizeof(buf),fp))
					break;
				add = buf;
				add.replace(add.length()-1,1,_T(""));
			}
			lines.push_back(add);
		}
		fclose(fp);
	}
	return lines;
}


/*-------------------------------------------------------------------------
	[Function]
		PickupLightSrcName
	[Details]
		���l�[��������������𔲂��o��
	[Argument1] rname
		���l�[����
	[Return]
		���� : ������   ������Ȃ� : empty
-------------------------------------------------------------------------*/
TSTR CLProc::PickupLightSrcName(const TSTR& rname)
{
	TSTR dst = _T("");
	{// ����������
		for(UINT i=1500; i<20000; i+=100)
		{// �P���r��������loop
			const TSTR LSrcTmp = chl::LngToStr(i) + _T("K");
			if(rname.find(LSrcTmp.data()) != std::string::npos){
				dst = LSrcTmp;
				break;
			}
		}
	}
	return dst;
}


/*-------------------------------------------------------------------------
	[Function]
		CalcDE
	[Details]
		�F�����v�Z����
	[Argument1] S_L
		������Y��L�l
	[Argument2] S_a
		������Y��a�l
	[Argument3] S_b
		������Y��b�l
	[Argument4] E_L
		�]�������Y��L�l
	[Argument5] E_a
		�]�������Y��a�l
	[Argument6] E_b
		�]�������Y��b�l
	[Return]
		�v�Z����
-------------------------------------------------------------------------*/
double CLProc::CalcDE(const TSTR& S_L, const TSTR& S_a, const TSTR& S_b, const TSTR& E_L, const TSTR& E_a, const TSTR& E_b)
{
	const double dS_L = chl::StrToDbl(S_L);
	const double dS_a = chl::StrToDbl(S_a);
	const double dS_b = chl::StrToDbl(S_b);
	const double dE_L = chl::StrToDbl(E_L);
	const double dE_a = chl::StrToDbl(E_a);
	const double dE_b = chl::StrToDbl(E_b);
	return pow((pow((dE_L-dS_L),2.0)+pow((dE_a-dS_a),2.0)+pow((dE_b-dS_b),2.0)),0.5);
}


/*-------------------------------------------------------------------------
	[Function]
		MakeSummaryData
	[Details]
		�}�N�������p�̂܂Ƃ߃f�[�^�𐶐�����
	[Argument1] body
		�]���{�f�B��
	[Argument2] otudir
		�t�@�C���o�͐�
	[Return]
		�܂Ƃ߃f�[�^
-------------------------------------------------------------------------*/
TSTR CLProc::MakeSummaryData(const TSTR& body, const TSTR& outdir)
{
	TSTR dst = _T("");
	{	
		for(UINT i=0; i<sizeof(sp)/sizeof(sp[0]); i++)
		{// �������ߗ��f�[�^�܂Ƃ�
			const TSTR nameadd = sp[i].LensName+_T(",");
			dst += (!i? nameadd:_T("/")+nameadd);
			for(UINT j=0; j<sp[i].Wdt.size(); j++)
			{
				dst += (j==sp[i].Wdt.size()-1? sp[i].Wdt[j]:sp[i].Wdt[j]+_T(","));
			}
		}
		dst += _T("$");
		
		for(UINT i=0; i<Resdt.size(); i++)
		{// �F���f�[�^�܂Ƃ�
			const TSTR nameadd = Resdt[i].LightSrc+_T(",");
			dst += (!i? nameadd:_T("/")+nameadd);
			for(UINT j=0; j<Resdt[i].Dedt.size(); j++)
			{
				const TSTR detmp = chl::DblToStr(Resdt[i].Dedt[j]);
				dst += (j==Resdt[i].Dedt.size()-1? detmp:detmp+_T(","));
			}
		}
		dst += _T("$");

		for(UINT i=0; i<Resdt.size(); i++)
		{// WB�f�[�^�܂Ƃ�
			const TSTR nameadd = Resdt[i].LightSrc+_T(",");
			dst += (!i? nameadd:_T("/")+nameadd);
			TSTR wbadd = chl::DblToStr(Resdt[i].StdDt.N5_R)+_T(",");
			{
				wbadd += (chl::DblToStr(Resdt[i].StdDt.N5_G)+_T(","));
				wbadd += (chl::DblToStr(Resdt[i].StdDt.N5_B)+_T(","));
				wbadd += (chl::DblToStr(Resdt[i].EvaDt.N5_R)+_T(","));
				wbadd += (chl::DblToStr(Resdt[i].EvaDt.N5_G)+_T(","));
				wbadd += chl::DblToStr(Resdt[i].EvaDt.N5_B);
			}
			dst += wbadd;
		}
		dst += (_T("$")+body);
		if(outdir != _T("") && chl::PathIsDirectoryEX(outdir.data()))
			OutDir = outdir;
		dst += (_T("$")+OutDir);
	}
	return dst;
}

