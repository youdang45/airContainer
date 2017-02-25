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
	containerType_t m_conType;   //��������
	static const float m_DiMin;  //��С�ھ�
	static const float m_Pi;     //Բ����
	static const float m_Rho_w;  //ˮ���ܶ�
	static const float m_Rho;    //�ְ��ܶ�
	float m_Phi;                 //Ͳ�庸��ϵ��
	float m_Phi_t;               //�ӹܺ���ϵ��
	float m_V;                   //Ͳ���ݻ�
	float m_P;                   //����ѹ��
	float m_DiMax;               //����ھ�
	float m_Sigma;               //�����¶��µ�����Ӧ��
	float m_Sigma_T;             //Ͳ��ѡȡ������Ӧ��
	float m_Sigma_Tt;            //�ӹ�ѡȡ������Ӧ��
	float m_C_1;                 //�ְ��ȸ�ƫ��
	float m_C_2;                 //��ʴԣ��
	float m_Rm;                  //(δ��)
	float m_ReL;                 //����ǿ��
	int m_DiStep;                //�ھ����㲽��
	conInstallType_t m_instType; //��װ��ʽ:��ʽ����ʽ
	float m_roundedBase;
};

extern CString containerTypeStr[];
extern CString conInstallTypeStr[];

#endif /*_C_CONTAINER_MODEL_H*/
