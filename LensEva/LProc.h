#pragma once

#include "GFunc.h"

// テンプレート
#define TEMPLATE_FILE	_T("LensEva_Template.xls")
#define SPECTRAL_DEF	_T("DefSpectral.txt")


/**************************************************************************
	[Class]
		CLProc
	[Details]
		レンズ評価メイン処理
**************************************************************************/
class CLProc
{
private:	
	/**************************************************************************
		[Structure]
			EL_DATA
		[Details]
			読み込んだレンズ1つ分のデータを管理する構造体
	**************************************************************************/
	typedef struct EL_DATA{
		// リネーム名
		TSTR RName;
		// N5パッチのR平均
		double N5_R;
		// N5パッチのG平均
		double N5_G;
		// N5パッチのB平均
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
			光源1つあたりのデータを管理する構造体
	**************************************************************************/
	typedef struct L_DATA{
		// 光源名
		TSTR LightSrc;
		// 基準データ
		ELDT StdDt;
		// 評価データ
		ELDT EvaDt;
		// 色差データ
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
			レンズ1つあたりの分光透過率データを管理する構造体
	**************************************************************************/
	typedef struct SP_DATA{
		// レンズ名
		TSTR LensName;
		// 波長ごとのデータ
		std::vector<TSTR> Wdt;

		/* Constructor */
		SP_DATA()
		{
			LensName.empty();
			Wdt.clear();
		}
	} SPDT;

private:
	// 読み込んだ全ての結果を格納するメンバ
	std::vector<LDT> Resdt;
	// 分光透過率を格納するメンバ
	SPDT sp[3];
	// 出力ディレクトリ
	TSTR OutDir;

public:
	CLProc(void);
	~CLProc(void);

	// 実行
	int Execution(const std::vector<chl::FilePInfo>& src, const TSTR& body, const TSTR& outdir);
	
private:
	// ファイル読み込み
	int ReadFile(const std::vector<chl::FilePInfo>& src);
	// csvファイル1つ読み込み
	void ReadCsv(const TSTR& path);
	// txtファイル1つ読み込み
	void ReadTxt(const TSTR& path);
	// ファイルを読み込み全行データを取得する
	std::vector<TSTR> GetLines(const TSTR& path);
	// リネーム名から光源名を抜き出す
	TSTR PickupLightSrcName(const TSTR& rname);
	// 色差を計算する
	double CalcDE(const TSTR& S_L, const TSTR& S_a, const TSTR& S_b, const TSTR& E_L, const TSTR& E_a, const TSTR& E_b);
	// マクロ処理用のまとめデータを生成する
	TSTR MakeSummaryData(const TSTR& body, const TSTR& outdir);

};
