#ifndef _C_GB_STANDARD_
#define _C_GB_STANDARD_

#include <list>
#include "CExlHeader.h"

using namespace std;

#define GB_STANDARD_FILE _T("许用应力表.xlsx")

typedef enum SteelNumber_e {
    SteelNumberNone = 0,
	Q245R,
	Q345R,
    S30408,
    Q235B,
	E_10,
	E_20,
	E_16Mn,
	SteelNumberMax,
} SteelNumber_t;

typedef struct SteelNumStringMap_s {
    SteelNumber_t steelNum;
	CString       steelNumName;
}SteelNumStringMap_t;

typedef enum GBStandardNumber_e {
	GBStandardNoNone = 0,
	GB_T_8163,
	GB_9948,
	GB_6479,
	GB_T_14976,
    GBStandardNoMax,
} GBStandardNumber_t;

typedef struct GBStandardStrMap_s {
    GBStandardNumber_t GBNum;
	CString            GBNumName;
} GBStandardStrMap_t;

typedef enum TempLevel_e {
	TEMP_LEVEL_INVALID = 0,
	TEMP_LEVEL1,
	TEMP_LEVEL2,
} TempLevel_t;

#define TempLevelMax       18  //温度分级最大数量
typedef enum tempLevel1_e {
    TEMP_LEVEL1_0_20 = 0,
	TEMP_LEVEL1_20_100,
	TEMP_LEVEL1_100_150,
	TEMP_LEVEL1_150_200,
	TEMP_LEVEL1_200_250,
	TEMP_LEVEL1_250_300,
	TEMP_LEVEL1_300_350,
	TEMP_LEVEL1_350_400,
	TEMP_LEVEL1_400_425,
	TEMP_LEVEL1_425_450,
	TEMP_LEVEL1_450_475,
} tempLevel1_t;

typedef struct TEMP_LEVEL1ScopeMap1_s 
{
	unsigned short min;
	unsigned short max;
    tempLevel1_t level;
}tempScopeMap1_t;

typedef enum tempLevel2_e {
    TEMP_LEVEL2_0_20 = 0,
	TEMP_LEVEL2_20_100,
	TEMP_LEVEL2_100_150,
	TEMP_LEVEL2_150_200,
	TEMP_LEVEL2_200_250,
	TEMP_LEVEL2_250_300,
	TEMP_LEVEL2_300_350,
	TEMP_LEVEL2_350_400,
	TEMP_LEVEL2_400_450,
	TEMP_LEVEL2_450_500,
	TEMP_LEVEL2_500_525,
	TEMP_LEVEL2_525_550,
	TEMP_LEVEL2_550_575,
	TEMP_LEVEL2_575_600,
	TEMP_LEVEL2_600_625,
	TEMP_LEVEL2_625_650,
	TEMP_LEVEL2_650_675,
	TEMP_LEVEL2_675_700,
} tempLevel2_t;

typedef struct tempScopeMap2_s 
{
	unsigned short min;
	unsigned short max;
    tempLevel2_t level;
}tempScopeMap2_t;

typedef struct stressScope_s{
	unsigned short beginTemp;
	unsigned short endTemp;
	unsigned short beginStress;
	unsigned short endStress;
}stressScope_t;

class CGBItemLine{
public:
	float thicknessMin;                     //最小壁厚
	float thicknessMax;                     //最大壁厚
    unsigned short Rm;                      //室温强度指标，单位：MPa
	unsigned short ReL;                     //室温强度指标，单位：MPa
	unsigned short stress[TempLevelMax];    //各温度下的许用应力
public:
	CGBItemLine();
	~CGBItemLine();
	bool getStressScope(SteelNumber_t material, short temp, stressScope_t &stressScope);
	unsigned int getLowestTempStress();
};

typedef struct GBItem_s { 
    SteelNumber_t steelNo;                  //钢号
	GBStandardNumber_t GBNo;                //国标号
	list <CGBItemLine> lineList;            //所有行
} GBItem_t;

class CGBStandard
{
public:
    static CGBStandard * GetInstance();
	~CGBStandard();
	bool LoadGBStandard();
	bool LoadPlateStd(CWorksheet &ecSheet);
	bool LoadPipeStd(CWorksheet &ecSheet);
	void closeExlApp();
	bool getPlateGBItem(SteelNumber_t &conMaterial, GBItem_t &item);
	bool getPlateStressScope(SteelNumber_t &material, float Di,
								short temp, stressScope_t &stressScope);
	float getPlateMaxThick(SteelNumber_t &material);
	list <GBItem_t> GetPipeItemList(SteelNumber_t &material);

private:
	CGBStandard();

private:
	CString m_stdFile;
	list<GBItem_t> m_plateGBItemList; //钢板国标条目
    list<GBItem_t> m_pipeGBItemList;  //钢管国标条目
	CApplication m_ecApp;
    CWorkbooks m_ecBooks;
    CWorkbook m_ecBook;
    CWorksheets m_ecSheets;
    CWorksheet m_ecSheet;
	CRange m_usedRange;
	CRange m_columns;
	CRange m_rows;
	CRange m_cell, m_cell1, m_cell2;
	static CGBStandard *m_pInstance;
};

extern const float g_defPhi;
extern SteelNumStringMap_t steelNumStrMap[];
extern GBStandardStrMap_t GBStandardStrMap[];
extern tempScopeMap1_t tempScopeMap1[];
extern tempScopeMap2_t tempScopeMap2[];
extern TempLevel_t steelLevelMap[];

extern float calcSelectedStress(stressScope_t &stressScope, short temp);

extern void setCauterizationRedundant(SteelNumber_t meterailType, float config, float &C_2);
extern bool getThickScopeFromString(CString str, float &min, float &max);
extern SteelNumber_t getSteelNumberByName(CString steelName);
extern GBStandardNumber_t getGBStdNumberByName(CString steelName);
extern CString getSteelNameByNum(SteelNumber_t num);
extern CString getGBStdNameByNum(GBStandardNumber_t gbNum);
extern void plateItemReset(GBItem_t &gbItem);
extern float getPipeMaxThick(GBItem_t &gbItem);

#endif