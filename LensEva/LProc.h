#pragma once

#include "GFunc.h"

// �e���v���[�g
#define TEMPLATE_FILE	_T("LensEva_Template.xls")
#define SPECTRAL_DEF	_T("DefSpectral.txt")


/**************************************************************************
	[Class]
		CLProc
	[Details]
		�����Y�]�����C������
**************************************************************************/
class CLProc
{
private:	
	/**************************************************************************
		[Structure]
			EL_DATA
		[Details]
			�ǂݍ��񂾃����Y1���̃f�[�^���Ǘ�����\����
	**************************************************************************/
	typedef struct EL_DATA{
		// ���l�[����
		TSTR RName;
		// N5�p�b�`��R����
		double N5_R;
		// N5�p�b�`��G����
		double N5_G;
		// N5�p�b�`��B����
		double N5_B;
		
		/* Constructor */
		EL_DATA()
		{
			RName.empty();
			N5_R = 0.0;
			N5_G = 0.0;
			N5_B = 0.0;
		}
	} ELDT;

	/**************************************************************************
		[Structure]
			L_DATA
		[Details]
			����1������̃f�[�^���Ǘ�����\����
	**************************************************************************/
	typedef struct L_DATA{
		// ������
		TSTR LightSrc;
		// ��f�[�^
		ELDT StdDt;
		// �]���f�[�^
		ELDT EvaDt;
		// �F���f�[�^
		std::vector<double> Dedt;

		/* Constructor */
		L_DATA()
		{
			LightSrc.empty();
			Dedt.clear();
		}
	} LDT;

	/**************************************************************************
		[Structure]
			SP_DATA
		[Details]
			�����Y1������̕������ߗ��f�[�^���Ǘ�����\����
	**************************************************************************/
	typedef struct SP_DATA{
		// �����Y��
		TSTR LensName;
		// �g�����Ƃ̃f�[�^
		std::vector<TSTR> Wdt;

		/* Constructor */
		SP_DATA()
		{
			LensName.empty();
			Wdt.clear();
		}
	} SPDT;

private:
	// �ǂݍ��񂾑S�Ă̌��ʂ��i�[���郁���o
	std::vector<LDT> Resdt;
	// �������ߗ����i�[���郁���o
	SPDT sp[3];
	// �o�̓f�B���N�g��
	TSTR OutDir;

public:
	CLProc(void);
	~CLProc(void);

	// ���s
	int Execution(const std::vector<chl::FilePInfo>& src, const TSTR& body, const TSTR& outdir);
	
private:
	// �t�@�C���ǂݍ���
	int ReadFile(const std::vector<chl::FilePInfo>& src);
	// csv�t�@�C��1�ǂݍ���
	void ReadCsv(const TSTR& path);
	// txt�t�@�C��1�ǂݍ���
	void ReadTxt(const TSTR& path);
	// �t�@�C����ǂݍ��ݑS�s�f�[�^���擾����
	std::vector<TSTR> GetLines(const TSTR& path);
	// ���l�[��������������𔲂��o��
	TSTR PickupLightSrcName(const TSTR& rname);
	// �F�����v�Z����
	double CalcDE(const TSTR& S_L, const TSTR& S_a, const TSTR& S_b, const TSTR& E_L, const TSTR& E_a, const TSTR& E_b);
	// �}�N�������p�̂܂Ƃ߃f�[�^�𐶐�����
	TSTR MakeSummaryData(const TSTR& body, const TSTR& outdir);

};
