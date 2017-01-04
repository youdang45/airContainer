#ifndef _C_PIPE_HOLE_H
#define _C_PIPE_HOLE_H
#include "CGBStandard.h"
#include "CExlHeader.h"

typedef enum holeSize_s {
	ES_INVALID = 0,
	ES_280_380,
	ES_300_400,
	ES_MAX
}holeSize_t;

typedef struct holeInDiameter_s {
	int b;
	int a;
} holeInDiameter_t;

typedef struct pipeConfig_s {
	bool isAddStress;
	bool extendIn;
	SteelNumber_t material;
	float Do;
	float thick;
}pipeConfig_t;

typedef enum AddPressComment_s {
	COMMENT_NONE = 0,
	COMMENT_ADD_PRESS,
	COMMENT_NOT_ADD_PRESS,
}AddPressComment_t;

struct pipeCalcResult_t{
	float minThick;
	float baseThick;
	float minExtInHeight; //TBD: change height to length
	float minExtOutHeight;
	CString GBStr;
	float minWidth;
	float addPressThick;
	AddPressComment_t recommendAddPress;
public:
	pipeCalcResult_t ();
};

typedef struct holeConfig_s {
	SteelNumber_t material;
	holeSize_t    size;
}holeConfig_t;

typedef struct holeCalcResult_s {
	float thickness;
	float extInLenght;
	float extOutLenght;
}holeCalcResult_t;

class CPipeHole {
public:
	static CPipeHole * GetInstance();
	~CPipeHole();
	void addPipeConfig(pipeConfig_t &pipeConfig);
	list <pipeConfig_t> &getPipeConfigList();
	void addPipeCalcResult(pipeCalcResult_t &pipeCalcResult);
	void setHoleConfig(holeConfig_t &holeConfig);
	holeConfig_t &getHoleConfig();
	void setHoleCalcResult(holeCalcResult_t &holeResult);
	list <pipeCalcResult_t> &getPipeCalcResults();
	holeCalcResult_t & getHoleCalcResult();
	void reset();
	bool save();

private:
	CPipeHole();
	void closeExlApp();

private:
	static CPipeHole *m_pInstance;
	list <pipeConfig_t> m_configList;
	list <pipeCalcResult_t> m_pipeCalcResultList;
	holeConfig_t m_holeConfig;
	holeCalcResult_t m_holeResult;
	CApplication m_ecApp;
    CWorkbooks m_ecBooks;
    CWorkbook m_ecBook;
    CWorksheets m_ecSheets;
    CWorksheet m_ecSheet;
	CRange m_range;
};

extern holeInDiameter_t holeSize[];
extern holeSize_t getHoleSizeByStr(CString);
#endif