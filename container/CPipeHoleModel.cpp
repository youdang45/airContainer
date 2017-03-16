#include "stdafx.h"
#include <cmath>
#include "CPipeHoleModel.h"
#include "CContainerInfo.h"
#include "CGBStandard.h"



CPipeModel *CPipeModel::m_pInstance = NULL;
const float CPipeModel::m_Phi_t = 1;


CPipeModel::CPipeModel()
{
	m_C_1 = 0;
	m_C_2 = 0;
}

CPipeModel::~CPipeModel()
{
}

CPipeModel * CPipeModel::GetInstance()  
{
	if (m_pInstance == NULL) {
		m_pInstance = new CPipeModel();
	}
	return m_pInstance;
}

bool CPipeModel::calcPipes()
{
	bool ret = TRUE;
	CPipeHole *pPipeHole = CPipeHole::GetInstance();

	list <pipeConfig_t> & configList = pPipeHole->getPipeConfigList();
	for (list <pipeConfig_t>::iterator iter =  configList.begin();
		iter != configList.end(); iter++) {
		ret = calcPerPipe(*iter);
		if (ret == FALSE) {
			//Add Err Log
			return FALSE;
		}
	}

	return TRUE;
}

bool CPipeModel::calcPerPipe(pipeConfig_t &pipeConf)
{
	bool ret = FALSE;
	pipeCalcResult_t result;
	CGBStandard *gb = CGBStandard::GetInstance();
	CContainerInfo *pconInfo = CContainerInfo::GetInstance();
	ConCaclResultItem &selectedContainer = pconInfo->getSelectItem();
	CPipeHole *pPipeHole = CPipeHole::GetInstance();
	conConfig_t config;
	short temp = 0;
	stressScope_t stressScope; 
	float Delet_nt = 0;
	float Delet_ntMin = 0;
	float maxThick = 0;
	float thick = 0;
	bool found = FALSE;

	pconInfo->getContainerConfig(config);
	m_P = config.pressure;
	temp = config.temperature;
	m_C_1 = config.thickNegWindage;
	m_C_2 = config.cauterization; 
	m_Di = selectedContainer.diameterIn;//筒体内径
	m_Sigma_T = selectedContainer.selectedStress;
	m_Delta_1 = selectedContainer.conCalcThickTmp;
	m_Delta_1n = selectedContainer.conThick;
	m_Delta_1e = selectedContainer.conValidThick;

	list <GBItem_t> gbList = gb->GetPipeItemList(pipeConf.material);
	for (list <GBItem_t>::iterator iter1 = gbList.begin();
		iter1 != gbList.end(); iter1++)
	{
		GBStandardNumber_t gbNo = iter1->GBNo;

		maxThick = getPipeMaxThick(*iter1);
		for (list <CGBItemLine>::iterator iter2 = iter1->lineList.begin();
			iter2 != iter1->lineList.end(); iter2++) {

			memset(&stressScope, 0, sizeof(stressScope));
			ret = iter2->getStressScope(pipeConf.material, temp, stressScope);
			if (ret == FALSE) {
				//Add Error Info.
				return FALSE;
			}
			m_sigma_Tt = calcSelectedStress(stressScope, temp);
			m_Delet_t = m_P * pipeConf.Do/(2*m_sigma_Tt*m_Phi_t + m_P);

			if(pipeConf.isAddStress && 
				(!isDelet_ntValid(pipeConf.Do, pipeConf.thick))) {
				return FALSE;
			}

			m_Delet_ntMin = calcDeletNT(m_Delet_t);
			if (pipeConf.isAddStress) {
				if (m_Delet_ntMin >= pipeConf.thick){
					thick = m_Delet_ntMin;
				} else {
					thick = pipeConf.thick;
					result.baseThick = m_Delet_ntMin;
				}
				ret = pipeSubCalc(pipeConf, thick, gbNo, result);
				if (ret == TRUE) {
					found = TRUE;
					break;
				}
			} else {
				for (Delet_nt = m_Delet_ntMin; Delet_nt < maxThick; Delet_nt += 0.1){
					ret = pipeSubCalc(pipeConf, Delet_nt, gbNo, result);
					if (ret == TRUE) {
						found = TRUE;
						break;
					}
				}
			}
			if (found) {
				break;
			}
		}
	}

	pPipeHole->addPipeCalcResult(result);
}

bool CPipeModel::pipeSubCalc(pipeConfig_t &pipeConf, float Delet_nt,
							GBStandardNumber_t gbNo, pipeCalcResult_t &result)
{
	float  Delet_et = 0;

	m_Dop = pipeConf.Do - 2*(Delet_nt - m_C_2);
	float B = calcWidth(Delet_nt);
	calcLenght(m_Dop, Delet_nt, pipeConf.extendIn);
	Delet_et = Delet_nt - m_C_2;

	
	float rate = (m_sigma_Tt / m_Sigma_T);
	float f_r = ( rate > 1.0 ) ? 1.0 : rate;
	float A = m_Dop*m_Delta_1 + 2 * m_Delta_1 * Delet_et*(1-f_r);
	float A1 = (B-m_Dop)*(m_Delta_1e - m_Delta_1)-2*Delet_et*(m_Delta_1e-m_Delta_1)*(1-f_r);
	float A2 = 2*m_h_1*(Delet_et - m_Delet_t)*f_r+2*m_h_2*(Delet_et-m_C_2)*f_r;
	float A3 = (Delet_nt < 6) ? pow(Delet_nt, 2)/2 : 18;
	float Ae = A1 + A2 + A3;
	if (Ae >= A) {
		m_delet_nr = Delet_nt;
		result.minThick = Delet_nt;
		float v = (m_Di/2.0 - sqrt(pow(m_Di, 2)/4 - pow(pipeConf.Do, 2)/4));
		if (pipeConf.extendIn) {
			result.minExtInHeight = (v > m_h_2) ? v : m_h_2;
		} else {
			result.minExtInHeight = 0;
		}
		result.minExtOutHeight = m_h_1;
		if (!result.GBStr.IsEmpty()) {
			result.GBStr += _T(", ");
		}
		CString str = getGBStdNameByNum(gbNo);
		result.GBStr += str;
		if (pipeConf.isAddStress) {
			result.recommendAddPress = COMMENT_NOT_ADD_PRESS;
		} else {
			if ((Delet_nt/m_Delta_1n) > 2) {
				result.recommendAddPress = COMMENT_ADD_PRESS;
			}
		}
		return TRUE;
	} else {
		if (pipeConf.isAddStress) {
			result.minThick = Delet_nt;
			float v = (m_Di/2.0 - sqrt(pow(m_Di, 2)/4 - pow(pipeConf.Do, 2)/4));
			if (pipeConf.extendIn) {
				result.minExtInHeight = (v > m_h_2) ? v : m_h_2;
			} else {
				result.minExtInHeight = 0;
			}
			result.minExtOutHeight = m_h_1;
			if (!result.GBStr.IsEmpty()) {
				result.GBStr += _T(", ");
			}
			CString str = getGBStdNameByNum(gbNo);
			result.GBStr += str;
			result.minWidth = B/2.0 - pipeConf.Do / 2.0;
			result.addPressThick = (A-Ae) / result.minWidth;
			return TRUE;
		}
	}

	return FALSE;
}

float CPipeModel::calcDeletNT(float Delet_t)
{
	float v = Delet_t + m_C_2;
	int x = int(v*100);
	int ret = x%10;
	if (ret > 0) {
		return float((x/10+1)*10)/100.0;
	} else {
		return float(x)/100.0;
	}
}

float CPipeModel::calcWidth(float delet_n)
{
	float B1 = 2.0*m_Dop;
	float B2 = m_Dop + 2.0*m_Delta_1n + 2*delet_n;

	return (B1 > B2) ? B1 : B2;
}

void CPipeModel::calcLenght(float Dop, float delet_n, bool isExtIn)
{
	m_h_1 = sqrt((Dop*delet_n));
	if (isExtIn) {
		m_h_2 = m_h_1;
	} else {
		m_h_2 = 0;
	}
}

bool CPipeModel::isDelet_ntValid(float Do, float delet_nt)
{
	float v = m_P*Do / (2*m_Sigma_Tt*m_Phi_t+m_P) + m_C_2;

	if (delet_nt >= v) {
		return TRUE;
	} else {
		return FALSE;
	}
}


//----------------------------------------------------------------------------
const float CHoleModel::A3r = 18;
CHoleModel *CHoleModel::m_pInstance = NULL;

CHoleModel::CHoleModel()
{
	m_C_1 = 0.3;
	m_C_2 = 0;
	m_Di = 0;

}

CHoleModel::~CHoleModel()
{
}

CHoleModel * CHoleModel::GetInstance()  
{  
	if(m_pInstance == NULL)  //判断是否第一次调用  
		m_pInstance = new CHoleModel();  
	return m_pInstance;  
}

bool CHoleModel::calcHole()
{
	bool ret = TRUE;
	bool found = FALSE;
	holeCalcResult_t holeResult;
	stressScope_t    stressScope;
	conConfig_t config;
	float sigma_Tr = 0;
	containerType_t conType = CONTAINER_TYPE_INVALID;
	float volume = 0;
	SteelNumber_t conMetarial = SteelNumberNone;
	short temp = 0;
	float minThick = 0, maxThick = 0;
	CGBStandard *gb = CGBStandard::GetInstance();
	CContainerInfo *pconInfo = CContainerInfo::GetInstance();
	ConCaclResultItem &selectedContainer = pconInfo->getSelectItem();
	CPipeHole *pPipeHole = CPipeHole::GetInstance();
	
	memset(&holeResult, 0, sizeof(holeResult));
	memset(&stressScope, 0, sizeof(stressScope));

	pconInfo->getContainerConfig(config);
	conType = config.conType;
	volume = config.volume;
	m_P = config.pressure;
	conMetarial = config.conMetarial;
	temp = config.temperature;
	m_Phi = config.coefficient;
	m_C_1 = config.thickNegWindage;
	m_C_2 = config.cauterization;
	m_Di = selectedContainer.diameterIn;//筒体内径
	m_Sigma_T = selectedContainer.selectedStress;
	m_Delta_1 = selectedContainer.conCalcThickTmp;
	m_Delta_1n = selectedContainer.conThick;
	m_Delta_1e = selectedContainer.conValidThick;

	holeConfig_t & holeConf = pPipeHole->getHoleConfig();

	m_size = holeConf.size;
	ret = initThick();
	if (ret == FALSE) {
		//ADD ERROR log
		return FALSE;
	}

	maxThick = gb->getPlateMaxThick(conMetarial);
	for (float delet_nr = m_delet_nr; delet_nr <= maxThick; delet_nr += 0.1) {
		calcLenght(delet_nr);
		float Br = calcWidth(delet_nr);

		memset(&stressScope, 0, sizeof(stressScope));
		ret = gb->getPlateStressScope(holeConf.material, delet_nr, temp, stressScope);
		if (ret == FALSE) {
			//Add Err log
			return FALSE;
		}
		sigma_Tr = calcSelectedStress(stressScope, temp);

		float rate = (sigma_Tr / m_Sigma_T);
		float f_rr = ( rate > 1.0 ) ? 1.0 : rate;
		float Ar = m_Dopr*m_Delta_1 + 2 * m_Delta_1 * (delet_nr - m_C_1 - m_C_2)*(1-f_rr);
		float A1r = (Br-m_Dopr)*(m_Delta_1e - m_Delta_1)-2*(delet_nr-m_C_1-m_C_2)*(m_Delta_1e-m_Delta_1)*(1-f_rr);
		float A2r = 2*m_h_1r*(delet_nr-m_C_1-m_C_2)*f_rr+2*m_h_2r*(delet_nr-m_C_1-m_C_2-m_C_2)*f_rr;
		float Aer = A1r + A2r + A3r;
		if (Aer >= Ar) {
			m_delet_nr = delet_nr;
			found = TRUE;
			break;
		}
	}

	if (found) {
		holeResult.thickness = m_delet_nr;
		holeResult.extInLenght = m_h_1r;
		float v = m_Di/2.0 - sqrt(pow(m_Di, 2)/4 - pow((holeSize[m_size].a+2*m_delet_nr), 2)/4.0);
		holeResult.extOutLenght = (v > m_h_2r ) ? v : m_h_2r;

		pPipeHole->setHoleCalcResult(holeResult);
		return TRUE;
	} else {
		return FALSE;
	}	
}

bool CHoleModel::initThick()
{

	if (m_P <= 0.85) {
		m_delet_nr = 20;
	} else if (m_P > 0.85 && m_P <= 1.5) {
		switch (m_size) {
			case ES_280_380:
				m_delet_nr = 22;
				break;
			case ES_300_400:
				m_delet_nr = 25;
				break;
			default:
				m_delet_nr = 0;
				break;
		}
	} else {
		//MSG("使用人孔时，计算压力应不大于1.5MPa")
		return FALSE;
	}

	return TRUE;
}

bool CHoleModel::calcLenght(float delet_nr)
{
	m_Dopr = holeSize[m_size].b + 2*(m_C_1 + m_C_2);
	m_h_2r = m_h_1r = sqrt((m_Dopr*delet_nr));

	return TRUE;
}

float CHoleModel::calcWidth(float delet_nr)
{
	float Br1 = 2.0*m_Dopr;
	float Br2 = m_Dopr + 2.0*m_Delta_1n + 2*delet_nr;

	return (Br1 > Br2) ? Br1 : Br2;
}

float CHoleModel::calcHoleDeletN(float delet)
{
	float ret = 0;
	int n = 0;
	float sum = 0;

	sum = (delet + m_C_1 + m_C_2);
	n = sum / 0.5;
	ret = sum - n * 0.5;
	if (ret) {
		return (n + 1)*0.5;
	} else {
		return sum;
	}
}

