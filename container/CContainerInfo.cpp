#include "stdafx.h"
#include <algorithm>
#include "CContainerInfo.h"

CContainerInfo *CContainerInfo::m_pInstance = NULL;

bool containerSort(conCaclResultItem_t &s1, conCaclResultItem_t &s2)
{
	return s1.weight < s2.weight;
}

CContainerInfo::CContainerInfo()
{
	m_select = 0;
	m_lengthMax = 0;
	m_configed = FALSE;
}

CContainerInfo::~CContainerInfo()
{
}

CContainerInfo * CContainerInfo::GetInstance()  
{  
    if(m_pInstance == NULL)  //�ж��Ƿ��һ�ε���  
        m_pInstance = new CContainerInfo();  
    return m_pInstance;  
}

bool CContainerInfo::setContainerConfig(conConfig_t &config)
{
	//TBD: Add args check here
	m_conType = config.conType;
	m_volume = config.volume;
	m_pressure = config.pressure;
	m_conMetarial = config.conMetarial;
	m_temperature = config.temperature;
	m_coefficient = config.coefficient;
	m_thickNegWindage = config.thickNegWindage;
	m_cauterization = config.cauterization;
	m_conInstType = config.installType;
	m_outputNum = config.outputNum;
	m_lengthMin = config.lengthMin;   
	m_lengthMax = config.lengthMax;
	m_lengthDiRateMin = config.lengthDiRateMin;
	m_lengthDiRateMax = config.lengthDiRateMax;
	m_configed = TRUE;

	return TRUE;
}

bool CContainerInfo::getContainerConfig(conConfig_t &config)
{
	if (!m_configed) {
		return FALSE;
	}

	config.conType = m_conType;
	config.volume = m_volume;
	config.pressure = m_pressure;
	config.conMetarial = m_conMetarial;
	config.temperature = m_temperature;
	config.coefficient = m_coefficient;
	config.cauterization = m_cauterization;
	config.installType = m_conInstType;
	config.thickNegWindage = m_thickNegWindage;


	return TRUE;
}

void CContainerInfo::addContainerItem(conCaclResultItem_t *item)
{
	m_conCaclResults.push_back(*item);
}

void CContainerInfo::sortContainerResultList()
{
	m_conCaclResults.sort(containerSort);
}

list <conCaclResultItem_t> & CContainerInfo::getContainerResultList()
{
	return m_conCaclResults;
}

void CContainerInfo::setContainerSelectedNum(int i)
{
	m_select = i;
}

int CContainerInfo::getContainerSelectedNum()
{
	return m_select;
}

conCaclResultItem_t &CContainerInfo::getSelectItem()
{
	list <conCaclResultItem_t>::iterator iter;

	for (int i = 0; i < m_select; i++)
	{
		if (i != 0)
			iter++;
		else
			iter = m_conCaclResults.begin();
	}

	return *iter;
}

unsigned int CContainerInfo::getOutputNum()
{
	return m_outputNum;
}


void CContainerInfo::reset()
{
	m_conMetarial = SteelNumberNone;
	m_pressure = 0;
	m_temperature = 0;
	m_volume = 0;
	m_coefficient = 0;
	m_select = -1;
	m_conCaclResults.clear();
	m_configed = FALSE;
}

bool CContainerInfo::conResultFilter(float length, float Di)
{
	if (m_lengthMax > 0.00001) {
		if ((length > m_lengthMax) || (length < m_lengthMin)) {
			return FALSE;
		}
	}

	float rate = 0;
	rate = length / Di;

	if ((rate > m_lengthDiRateMax) || (rate < m_lengthDiRateMin)) {
		return FALSE;
	}

	return TRUE;
}

bool CContainerInfo::save(CString file)
{
	bool ret = FALSE;
	CString str;

	COleException pe;  
    if (!m_ecApp.CreateDispatch(_T("Excel.Application"), &pe))  
    {  
        AfxMessageBox(_T("Application����ʧ�ܣ���ȷ����װ��word 2000�����ϰ汾!"), MB_OK|MB_ICONWARNING);  
        pe.ReportError();  
        throw &pe;  
        return FALSE;  
    }
	m_ecApp.put_AlertBeforeOverwriting(FALSE);
	m_ecApp.put_DisplayAlerts(FALSE);

	m_ecBooks = m_ecApp.get_Workbooks();  
    if (!m_ecBooks.m_lpDispatch)   
    {  
        AfxMessageBox(_T("WorkBooks����ʧ��!"), MB_OK|MB_ICONWARNING);  
        return FALSE;  
    } 

	COleVariant VOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);  
    m_ecBook = m_ecBooks.Add(VOptional);
	if(!m_ecBook.m_lpDispatch)   
    {  
        AfxMessageBox(_T("�ļ���ʧ�ܻ�ȡʧ��!"), MB_OK|MB_ICONWARNING);  
        return FALSE;
    } 

	m_ecSheets = m_ecBook.get_Sheets();  
    if(!m_ecSheets.m_lpDispatch)   
    {
        AfxMessageBox(_T("WorkSheetsΪ��!"), MB_OK|MB_ICONWARNING);  
        return FALSE;
    }

	short firstSheet = 1;
	m_ecSheet = m_ecSheets.get_Item(COleVariant(firstSheet));  
    if(!m_ecSheet.m_lpDispatch)   
    {  
        AfxMessageBox(_T("WorkSheetΪ��!"), MB_OK|MB_ICONWARNING);  
        return FALSE;  
    }
	m_ecSheet.put_Name(_T("���޼�����"));

	m_range = m_ecSheet.get_Cells();
    if(!m_range.m_lpDispatch)   
    {  
        AfxMessageBox(_T("m_rangeΪ��!"), MB_OK|MB_ICONWARNING);  
        return FALSE;  
    }

	//��д������Ϣ
	CRange range = m_ecSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("I1")));
	range.Merge(VOptional);
	m_range.put_Item(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("����������Ϣ:")));

	m_range.put_Item(COleVariant((long)2),COleVariant((long)1),COleVariant(_T("��������")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)2),COleVariant(_T("�ݻ�(m3)")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)3),COleVariant(_T("����ѹ��(MPa)")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)4),COleVariant(_T("����¶�(��)"))); 
	m_range.put_Item(COleVariant((long)2),COleVariant((long)5),COleVariant(_T("���ӽ�ͷϵ��")));	
	m_range.put_Item(COleVariant((long)2),COleVariant((long)6),COleVariant(_T("����")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)7),COleVariant(_T("��ȸ�ƫ��(mm)")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)8),COleVariant(_T("��ʴԣ��(mm)")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)9),COleVariant(_T("��װ��ʽ")));

	m_range.put_Item(COleVariant((long)3),COleVariant((long)1),COleVariant(containerTypeStr[m_conType]));
	str.Format(_T("%.2f"), m_volume);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)2),COleVariant(str));
	str.Format(_T("%.2f"), m_pressure);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)3),COleVariant(str));
	str.Format(_T("%d"), m_temperature);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)4),COleVariant(str)); 
	str.Format(_T("%.2f"), m_coefficient);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)5),COleVariant(str));	 
	m_range.put_Item(COleVariant((long)3),COleVariant((long)6),COleVariant(steelNumStrMap[m_conMetarial].steelNumName));
	str.Format(_T("%.2f"), m_thickNegWindage);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)7),COleVariant(str));
	str.Format(_T("%.2f"), m_cauterization);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)8),COleVariant(str));
	m_range.put_Item(COleVariant((long)3),COleVariant((long)9),COleVariant(conInstallTypeStr[m_conInstType]));	



	range = m_ecSheet.get_Range(COleVariant(_T("A6")),COleVariant(_T("H6")));
	range.Merge(VOptional);
	m_range.put_Item(COleVariant((long)6),COleVariant((long)1),COleVariant(_T("����������:")));

    m_range.put_Item(COleVariant((long)7),COleVariant((long)1),COleVariant(_T("���")));
	m_range.put_Item(COleVariant((long)7),COleVariant((long)2),COleVariant(_T("Ͳ���ھ���mm��")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)3),COleVariant(_T("Ͳ��ں�mm��")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)4),COleVariant(_T("Ͳ��߶ȣ�mm��")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)5),COleVariant(_T("��ͷ�ں�mm��")));
	m_range.put_Item(COleVariant((long)7),COleVariant((long)6),COleVariant(_T("��ͷֱ�߸߶ȣ�mm��")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)7),COleVariant(_T("�����ܸߣ�mm��")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)8),COleVariant(_T("����������kg��")));

	int i = 8;
	for (list <conCaclResultItem_t>::iterator iter = m_conCaclResults.begin();
		iter != m_conCaclResults.end(); iter++) {

		str.Format(_T("%d"), (i-7));
		m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(str));
		str.Format(_T("%.2f"), iter->diameterIn);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(str));
		str.Format(_T("%.2f(%.2f)"), iter->conThick, iter->conCalcThick);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(str));
		str.Format(_T("%d"), (int)iter->conHight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(str));
		str.Format(_T("%.2f(%.2f)"), iter->capThick, iter->capCalcThick);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(str));
		str.Format(_T("%d"), (int)iter->capHight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)6),COleVariant(str));
		str.Format(_T("%d"), (int)iter->totalHight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)7),COleVariant(str));
		str.Format(_T("%d"), (int)iter->weight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)8),COleVariant(str));

		i++;
	}



    m_ecBook.SaveAs(  
        COleVariant(file),
        COleVariant((long)51),//�ò���ʹ�ñ����ʱ��ʹ�ü���ģʽ��������ʹ��excel2003�򿪵�ʱ�򵯳���ʾ��  
        VOptional,  
        VOptional,  
        VOptional,  
        VOptional,  
        0,  
        VOptional,  
        VOptional,  
        VOptional,  
        VOptional,  
        VOptional  
        );


	closeExlApp();

	AfxMessageBox(_T("���޼���������ɹ�!"), MB_OK|MB_ICONINFORMATION);  

	return TRUE;
}

void CContainerInfo::closeExlApp()
{
	COleVariant vTrue((short)TRUE),      
                vFalse((short)FALSE),  
                vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);  
 
    m_ecApp.Quit();
    if(m_range.m_lpDispatch)  
        m_range.ReleaseDispatch();  
    if(m_ecSheet.m_lpDispatch)  
        m_ecSheet.ReleaseDispatch();  
    if(m_ecSheets.m_lpDispatch)  
        m_ecSheets.ReleaseDispatch();  
    if(m_ecBook.m_lpDispatch)  
        m_ecBook.ReleaseDispatch();  
    if(m_ecBooks.m_lpDispatch)  
        m_ecBooks.ReleaseDispatch();  
    if(m_ecApp.m_lpDispatch)  
        m_ecApp.ReleaseDispatch();  
}

containerType_t getContainerTypeByStr(CString &str)
{
	for (int i = CONTAINER_TYPE_INVALID + 1; i < CONTAINER_TYPE_MAX; i++) {
		if (str == containerTypeStr[i]) {
			return (containerType_t )i;
		}
	}
	return CONTAINER_TYPE_INVALID;
}

conInstallType_t getConInstallTypeByStr(CString &str)
{
	for (int i = CON_INST_TYPE_INVALID + 1; i < CON_INST_TYPE_MAX; i++) {
		if (str == conInstallTypeStr[i]) {
			return (conInstallType_t )i;
		}
	}
	return CON_INST_TYPE_INVALID;
}


bool simpleConInputIsValid(float pressure,
								float volume,
								short temperature)
{
	if (pressure > 1.6) {
		return FALSE;
	}

	if (volume > 1.0) {
		return FALSE;
	}

	if (pressure*volume > 1.0) {
		return FALSE;
	}

	if (temperature < -20.0 || temperature > 150.0) {
		return FALSE;
	}

	return TRUE;
}
	

