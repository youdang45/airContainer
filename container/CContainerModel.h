#ifndef _C_CONTAINER_MODEL_H
#define _C_CONTAINER_MODEL_H

#include "CGBStandard.h"


typedef enum checkReturn_e {
	CHECK_PASS = 0,
	CHECK_FAIL_1,
	CHECK_FAIL_2,
	CHECK_FAIL_3,
} checkReturn_t;

typedef enum containterType_e {
	CONTAINER_TYPE_INVALID = 0,
	CONTAINER_TYPE_GENERAL,
	CONTAINER_TYPE_SIMPLE,
	CONTAINER_TYPE_MAX,
}containerType_t;

typedef enum conInstallType_e {
	CON_INST_TYPE_INVALID = 0,
	CON_INST_TYPE_UPRIGHT,
	CON_INST_TYPE_HORIZANTAL,
	CON_INST_TYPE_MAX,
}conInstallType_t;

class CContainerModel
{
public:
	static CContainerModel * GetInstance();
	~CContainerModel();
	bool calcContainer();
	void setDiStep(int s);
	void setRoundedBase(float b);

private:
	CContainerModel();
	float calcMaxDiameterIn(float volume);
	float calcConThick();
	float calcConHight();
	float calcCapThick();
	float calcConWeight();
	float calcTotalHight();
	float calcCapHight();
	float calcLength(float V, float Di);
	void calcContainerByDiWalk(float thickMax, float thickMin, bool isFirstLine);
	float calcConDeletN(float delet);
	float calcCapDeletN(float delet);
	checkReturn_t checkConDiISOK(float L, float h, float Di, float Delta_1_n,
		                         float Delta_2_n, float thickMax, float thickMin, bool isFirstLine);
	float calcWeight(float l, float h, float Di, float Delta_1n, float Delta_2n);
	bool calcPotentialThick(float L, float Di, float Delta_1n, float Delta_2n, float thickMin,
							bool  isFirstLine, float &Delta_1n_out, float &Delta_2n_out);
	float rounded(float f);


private:
	static CContainerModel *m_pInstance;
	SteelNumber_t m_conMetarial;  
	containerType_t m_conType;   //容器类型
	static const float m_DiMin;  //最小内径
	static const float m_Pi;     //圆周率
	static const float m_Rho_w;  //水的密度
	static const float m_Rho;    //钢板密度
	float m_Phi;                 //筒体焊接系数
	float m_Phi_t;               //接管焊接系数
	float m_V;                   //筒体容积
	float m_P;                   //计算压力
	float m_DiMax;               //最大内径
	float m_Sigma;               //试验温度下的许用应力
	float m_Sigma_T;             //筒体选取的许用应力
	float m_Sigma_Tt;            //接管选取的许用应力
	float m_C_1;                 //钢板厚度负偏差
	float m_C_2;                 //腐蚀裕量
	float m_Rm;                  //(未用)
	float m_ReL;                 //屈服强度
	int m_DiStep;                //内径计算步长
	conInstallType_t m_instType; //安装方式:立式、卧式
	float m_roundedBase;
};

extern CString containerTypeStr[];
extern CString conInstallTypeStr[];

#endif /*_C_CONTAINER_MODEL_H*/
