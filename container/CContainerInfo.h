#ifndef _C_CONTAINER_INFO_H
#define _C_CONTAINER_INFO_H

#include <list>
#include "CGBStandard.h"
#include "CExlHeader.h"
#include "CContainerModel.h"

typedef struct conConfig_s {
	containerType_t conType;
	float volume;
	float pressure;
	short temperature;
	float coefficient;
	SteelNumber_t conMetarial;
	float thickNegWindage;
	float cauterization;
	float DiStep;
	conInstallType_t installType;
	float lengthMin;
	float lengthMax;
	float lengthDiRateMin;
	float lengthDiRateMax;
	unsigned int outputNum;
} conConfig_t;


typedef struct conCaclResultItem_s {
	float diameterIn;           //筒体内径
	float conThick;             //筒壁厚度
	float conCalcThick;         //筒体计算厚度
	float conCalcThickTmp;      //用于接管计算的筒壁厚度
	float conHight;             //筒体高度
	float capThick;             //封头厚度
	float capCalcThick;         //封头计算厚度
	float weight;               //壳体重量
	float totalHight;           //壳体总高
	float capHight;             //封头直边高度
	float conValidThick;        //筒体有效厚度
	float selectedStress;       //筒体设计温度下的许用应力
}conCaclResultItem_t;

class CContainerInfo
{
public:
	~CContainerInfo();
    static CContainerInfo * GetInstance();
	bool setContainerConfig(conConfig_t &config);
	bool getContainerConfig(conConfig_t &config);
	void addContainerItem(conCaclResultItem_t *item);
	void sortContainerResultList();
	list <conCaclResultItem_t> & getContainerResultList();
	void setContainerSelectedNum(int i);
	int getContainerSelectedNum();
	conCaclResultItem_t &getSelectItem();
	unsigned int getOutputNum();
	bool conResultFilter(float height, float Di);
	void reset();
	bool save(CString file);

private:
	void closeExlApp();

private:
	static CContainerInfo *m_pInstance;
	CContainerInfo();
    float m_volume;    //容积
	float m_pressure;  //计算压力
	SteelNumber_t m_conMetarial;  //筒体材料
	unsigned short m_temperature; //设计温度
	float m_coefficient;          //焊接系数
	list <conCaclResultItem_t> m_conCaclResults;
	int m_select;
	bool m_configed;
	CApplication m_ecApp;
    CWorkbooks m_ecBooks;
    CWorkbook m_ecBook;
    CWorksheets m_ecSheets;
    CWorksheet m_ecSheet;
	CRange m_range;
	float m_cauterization;
	containerType_t m_conType;
	conInstallType_t m_conInstType;
	unsigned int m_outputNum;
	float m_lengthMin;
	float m_lengthMax;
	float m_lengthDiRateMin;
	float m_lengthDiRateMax;
	float m_thickNegWindage;

};

extern containerType_t getContainerTypeByStr(CString &str);
extern conInstallType_t getConInstallTypeByStr(CString &str);
extern bool simpleConInputIsValid(float pressure,
								float volume,
								short temperature);

#endif
