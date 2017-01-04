#include "stdafx.h"
#include <cmath>
#include "CContainerModel.h"
#include "CContainerInfo.h"


CString containerTypeStr[CONTAINER_TYPE_MAX] = {
	"Invalid",
	"压力容器",
	"简单压力容器"
};


CString conInstallTypeStr[CON_INST_TYPE_MAX] = {
	"Invalid",
	"立式",
	"卧式"
};

const float CContainerModel::m_Pi = 3.14;
const float CContainerModel::m_DiMin = 150;
const float CContainerModel::m_Rho_w = 1000;
const float CContainerModel::m_Rho = 7850;


CContainerModel *CContainerModel::m_pInstance = NULL;
CContainerModel * CContainerModel::GetInstance()  
{  
    if(m_pInstance == NULL)  //判断是否第一次调用  
        m_pInstance = new CContainerModel();  
    return m_pInstance;  
}


CContainerModel::CContainerModel()
{
	m_conMetarial = SteelNumberNone;  
	m_conType = CONTAINER_TYPE_INVALID;
	m_Phi = g_defPhi;           
	m_Phi_t = g_defPhi;             
	m_V = 0;             
	m_P = 0;             
	m_DiMax = 0;         
	m_Sigma = 0;         
	m_Sigma_T = 0;       
	m_Sigma_Tt = 0;      
	m_C_1 = 0.3;         
	m_C_2 = 0;           
	m_Rm = 0;            
	m_ReL = 0;
	
}

CContainerModel::~CContainerModel()
{
}

bool CContainerModel::calcContainer()
{
	bool             ret = FALSE;
	CContainerInfo * pContainerInfo = CContainerInfo::GetInstance();
	CGBStandard    * pGBStandard    = CGBStandard::GetInstance();
	conConfig_t      config;
	short            temp = 0;        //温度
	GBItem_t         gbItem;
	stressScope_t    stressScope;
	float            phi = 0;
	float            C_2 = 0;

	ret = pContainerInfo->getContainerConfig(config);
	m_V = config.volume;
	m_P = config.pressure;
	m_conMetarial = config.conMetarial;
	m_conType = config.conType;
	temp = config.temperature;
	phi = config.coefficient;
	m_C_2 = config.cauterization;
	m_instType = config.installType;

	m_Phi = phi;
	m_Phi_t = phi;
	m_DiMax = calcMaxDiameterIn(m_V);

	ret = pGBStandard->getPlateGBItem(m_conMetarial, gbItem);
	if (ret == FALSE) {
		//Add Error Info.
		return FALSE;
	}

	for (list<CGBItemLine>::iterator iter = gbItem.lineList.begin();
		iter != gbItem.lineList.end(); iter++) {
		m_ReL = (*iter).ReL;

		memset(&stressScope, 0, sizeof(stressScope));
		ret = (*iter).getStressScope(m_conMetarial, temp, stressScope);
		if (ret == FALSE) {
			//Add Error Info.
			return FALSE;
		}
		m_Sigma = (*iter).getLowestTempStress();
		m_Sigma_T = calcSelectedStress(stressScope, temp);

		calcContainerByDiWalk((*iter).thicknessMax, (*iter).thicknessMin);
	}
	pContainerInfo->sortContainerResultList();
}	

void CContainerModel::calcContainerByDiWalk(float thickMax, float thickMin)
{
	conCaclResultItem_t conRestItem;
	CContainerInfo * pContainerInfo = CContainerInfo::GetInstance();
	float Delta_1 = 0;         //筒体计算厚度
	float Delta_1n = 0;        //筒体名义厚度
	float Delta_1_tmp = 0;     //用于接管/人孔计算
	float Delta_2 = 0;         //封头计算厚度
	float Delta_2n = 0;        //封头名义厚度
	float h = 0;               //封头直边高度
	float L = 0;               //圆筒长度
	checkReturn_t checkPass = CHECK_PASS;
	float weight = 0;

	for(float Di = m_DiMin; Di < m_DiMax; Di += m_DiStep){
		L = calcLength(m_V, Di);
		if (Di <= 2000) {
			h = 25;
		} else {
			h = 40;
		}
		Delta_1 = m_P * Di / (2 * m_Sigma_T * m_Phi - m_P);
		Delta_2 = m_P * Di / (2 * m_Sigma_T * m_Phi - 0.5 * m_P);
		Delta_1n = calcConDeletN(Delta_1);
		Delta_2n = calcCapDeletN(Delta_2);
		checkPass = checkConDiISOK(L, h, Di, Delta_1n, Delta_2n, thickMax, thickMin);
		switch(checkPass) {
			case CHECK_FAIL_1:
				return;
			case CHECK_FAIL_2:
				continue;
			default:
			{
				//check passed
				break;
			}				
		}
		bool filtRet = TRUE;
		filtRet = pContainerInfo->conResultFilter(h, Di);
		if (filtRet == FALSE) {
			continue;
		}

		Delta_1_tmp = m_P * Di / (2 * m_Sigma_T * 1.0 - m_P);
		weight = calcWeight(L, h, Di, Delta_1n, Delta_2n);

		conRestItem.diameterIn = Di;		
		conRestItem.conThick = Delta_1n;
		conRestItem.conCalcThickTmp = Delta_1_tmp;
		conRestItem.conCalcThick = Delta_1;
		conRestItem.conHight= (L - 2*h); 		
		conRestItem.capThick = Delta_2n; 
		conRestItem.capCalcThick = Delta_2;
		conRestItem.weight = weight;			
		conRestItem.totalHight = (L + Di/2);		
		conRestItem.capHight = h;
		conRestItem.selectedStress = m_Sigma_T;
		conRestItem.conValidThick = Delta_1n - m_C_1 - m_C_2;
		pContainerInfo->addContainerItem(&conRestItem);
	}

	return;
}


float CContainerModel::calcMaxDiameterIn(float V)
{
	float Di = 0;
	Di = pow(float(12*m_V/m_Pi), float(1.0/3.0))*1000;
	if (Di > 5000) {
		return 5000;
	} else {
		return Di;
	}
}

float CContainerModel::calcLength(float V, float Di)
{
	return (m_V*pow((float)10, 9)-2*m_Pi/24*pow(Di, 3))*4/m_Pi/pow(Di, 2);
}

float CContainerModel::calcWeight(float l, float h, float Di, float Delta_1n, float Delta_2n)
{
	float part1 = 0;
	float part2 = 0;

	part1 = m_Rho*(l-2*h)*m_Pi*Delta_1n*(Di+Delta_1n)*pow(10.0, -9);
	part2 = pow(Di, 2)/3.0 + 5.0/6.0 *Di*Delta_2n + 2.0/3 * pow(Delta_2n, 2) + (Di+Delta_2n)*h;


	return (part1 + 2*m_Rho*m_Pi*Delta_2n*part2*pow(10.0, -9));
}

float CContainerModel::calcConDeletN(float delet)
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

float CContainerModel::calcCapDeletN(float delta)
{
	float ret = 0;
	int n = 0;
	float sum = 0;

	sum = 1.12 * (delta + m_C_1) + m_C_2;
	n = sum / 0.5;
	ret = sum - n * 0.5;
	if (ret) {
		return (n + 1)*0.5;
	} else {
		return sum;
	}
}

checkReturn_t CContainerModel::checkConDiISOK(float L, float h, float Di, float Delta_1n,
						                      float Delta_2n, float thickMax, float thickMin)
{
	if ((Delta_2n > thickMax) || (Delta_1n > thickMax)) {
		return CHECK_FAIL_1;
	}

	//Check for Q235B
	if ((m_conType == CONTAINER_TYPE_GENERAL) 
		&& (m_conMetarial == Q235B)
		&& (Delta_1n > 16)){
		return CHECK_FAIL_1;
	}

	if ((Delta_2n <= thickMin) || (Delta_1n <= thickMin)) {
		return CHECK_FAIL_2;
	}

	if (L < 2*h) {
		return CHECK_FAIL_2;
	}

	float limit = 0;
	if (m_conType == CONTAINER_TYPE_GENERAL) {
		limit = (m_conMetarial == S30408) ? 2 : 3;

	} else if (m_conType == CONTAINER_TYPE_SIMPLE) {
		limit = (m_conMetarial == S30408) ? 1 : 2;
	}
	if (((Delta_1n - m_C_2) < limit) || ((Delta_2n - m_C_2) < limit)){
		return CHECK_FAIL_2;
	}


	if ((Delta_2n - m_C_2 - m_C_1) < 0.15/100*Di) {
		return CHECK_FAIL_2;
	}

	float testPress = 0;
	float assumedPress = 0;
	float value = 0;
	float tmp1 = 0, tmp2 = 0;

	testPress = 1.25 * m_P * m_Sigma / m_Sigma_T;
	if (m_instType == CON_INST_TYPE_UPRIGHT) {
		assumedPress = (L + Di/2 + Delta_2n + 200) * m_Rho_w * 9.8 * pow(10.0, -9);
	} else if (m_instType == CON_INST_TYPE_HORIZANTAL){
		assumedPress = (Di + Delta_2n + 200) * m_Rho_w * 9.8 * pow(10.0, -9);
	}
	tmp1 = Di + Delta_1n - m_C_1 - m_C_2;
	tmp2 = Delta_1n - m_C_1 - m_C_2;
	value = (testPress+assumedPress)*tmp1/(2*tmp2*m_Phi);

	if (value > (0.9 * m_ReL)) {
		return CHECK_FAIL_2;
	}

	return CHECK_PASS;
}

void CContainerModel::setDiStep(int s)
{
	m_DiStep = s;
}
