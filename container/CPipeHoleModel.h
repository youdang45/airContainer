#ifndef _C_PIPE_HOLE_MODEL_H
#define _C_PIPE_HOLE_MODEL_H

#include "CPipeHole.h"
#include "CContainerModel.h"

class CPipeModel {
public:
	~CPipeModel();
	static CPipeModel * GetInstance();
	bool calcPipes();

private:
	CPipeModel();
	bool calcPerPipe(pipeConfig_t &pipeConf);
	float calcDeltaNT(float Delta_t);
	float calcWidth(float delta_n);
	void calcLenght(float Dop, float delta_n, bool isExtIn);
	bool pipeSubCalc(pipeConfig_t &pipeConf, float Delta_nt,
		                float Delta_1n, pipeCalcResult_t &result);
	bool isDelta_ntValid(float Do, float delta_nt);
	float CPipeModel::pipeResCeil(float v);
	float CPipeModel::pipeJointingCalc(float delta_nt, float delta_1n);


private:
	static CPipeModel *m_pInstance;
	static const float m_Pi;    
	static const float m_Phi_t;  
             
                   
              
	float m_Sigma;                          
    
 
	float m_C_1;                 
	float m_C_2;  
	float m_Dop;
	float m_Delta_1; 
	float m_Sigma_T;             
	float m_Sigma_Tt; 
	float m_Delta_1n;
	float m_Delta_1e;
	float m_P;
	float m_h_1;
	float m_h_2;
	float m_delta_nr;
	float m_Di; 
	float m_sigma_Tt;
	float m_Delta_t;
	float m_Delta_ntMin;
};

class CHoleModel {
public:
	~CHoleModel();
	static CHoleModel * GetInstance();
	bool calcHole();

private:
	CHoleModel();
	bool initThick();
	bool calcLenght(float delta_nr);
	float calcWidth(float delta_nr);
	float calcHoleDeltaN(float delta);
	float holeJointingCalc(float delta_nr, float delta_1n);


private:
	static CHoleModel *m_pInstance;
	holeCalcResult_t m_result;

	float m_delta_nr;//�̽ں��, ������
	float m_P;
	float m_Da;  //�ھ�����
	float m_Db;  //�ھ�����
	float m_Ar;  //��ǿ���
	float m_f_rr; //�̽ڲ���������������Ӧ���ı�ֵ
	float m_Dopr; //�˿׿���ֱ��
	float m_Br;   //��Ч���
	float m_h_1r; //�����˿���Ч��ǿ�߶�
	float m_h_2r; //�����˿���Ч��ǿ�߶�
	float m_C_1;         //�ְ��ȸ�ƫ��
	float m_C_2;         //��ʴԣ��
	holeSize_t m_size;   //�˿׳ߴ�
	float m_Delta_1n;    //Ͳ��������
	float m_Delta_1e;    //Ͳ����Ч���
	float m_Di;		     //�����ھ�
	float m_Sigma_T;  
	float m_Delta_1;
	float m_Phi; 
	float m_thickStep;



};
#endif /*_C_PIPE_HOLE_MODEL_H*/
