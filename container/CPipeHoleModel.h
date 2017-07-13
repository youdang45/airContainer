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

	float m_delta_nr;//短节厚度, 名义厚度
	float m_P;
	float m_Da;  //内径长轴
	float m_Db;  //内径短轴
	float m_Ar;  //补强面积
	float m_f_rr; //短节材料与壳体材料许用应力的比值
	float m_Dopr; //人孔开孔直径
	float m_Br;   //有效宽度
	float m_h_1r; //外伸人孔有效补强高度
	float m_h_2r; //内伸人孔有效补强高度
	float m_C_1;         //钢板厚度负偏差
	float m_C_2;         //腐蚀裕量
	holeSize_t m_size;   //人孔尺寸
	float m_Delta_1n;    //筒体名义厚度
	float m_Delta_1e;    //筒体有效厚度
	float m_Di;		     //壳体内径
	float m_Sigma_T;  
	float m_Delta_1;
	float m_Phi; 
	float m_thickStep;



};
#endif /*_C_PIPE_HOLE_MODEL_H*/
