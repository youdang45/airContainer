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
	float calcDeletNT(float Delet_t);
	float calcWidth(float delet_n);
	void calcLenght(float Dop, float delet_n, bool isExtIn);
	bool pipeSubCalc(pipeConfig_t &pipeConf, float Delet_nt, pipeCalcResult_t &result);
	bool isDelet_ntValid(float Do, float delet_nt);
	float CPipeModel::pipeResCeil(float v);

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
	float m_delet_nr;
	float m_Di; 
	float m_sigma_Tt;
	float m_Delet_t;
	float m_Delet_ntMin;
};

class CHoleModel {
public:
	~CHoleModel();
	static CHoleModel * GetInstance();
	bool calcHole();

private:
	CHoleModel();
	bool initThick();
	bool calcLenght(float delet_nr);
	float calcWidth(float delet_nr);
	float calcHoleDeletN(float delet);

private:
	static CHoleModel *m_pInstance;
	holeCalcResult_t m_result;

	static const float A3r;
	float m_delet_nr;//短节厚度, 名义厚度
	float m_P;
	float m_Da;  //内径长轴
	float m_Db;  //内径短轴
	float m_Ar;  //补强面积
	float m_f_rr; //短节材料与壳体材料许用应力的比值
	float m_Dopr; //人孔开孔直径
	float m_Br;   //有效宽度
	float m_h_1r; //外伸人孔有效补强高度
	float m_h_2r; //内伸人孔有效补强高度
	float m_A_er; //人孔补强面积
	float m_A_1r; //壳体有效厚度减去计算厚度之外的多余面积
	float m_A_2r; //人孔有效厚度减去计算厚度之外的多余面积
	static float m_A_3r; //焊缝金属截面积
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