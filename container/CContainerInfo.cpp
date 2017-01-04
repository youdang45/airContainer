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
	m_heightMax = 0;
	m_configed = FALSE;
}

CContainerInfo::~CContainerInfo()
{
}

CContainerInfo * CContainerInfo::GetInstance()  
{  
    if(m_pInstance == NULL)  //判断是否第一次调用  
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
	m_heightMin = config.heightMin;   
	m_heightMax = config.heightMax;
	m_heightDiRateMin = config.heightDiRateMin;
	m_heightDiRateMax = config.heightDiRateMax;

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

bool CContainerInfo::conResultFilter(float height, float Di)
{
	if (m_heightMax > 0.00001) {
		if ((height > m_heightMax) || (height < m_heightMin)) {
			return FALSE;
		}
	}

	float rate = 0;
	rate = height / Di;

	if ((rate > m_heightDiRateMax) || (rate < m_heightDiRateMin)) {
		return FALSE;
	}

	return TRUE;
}

bool CContainerInfo::save()
{
	bool ret = FALSE;
	CString file;
	CString str;

	TCHAR chpath[100] = {0}; 
	GetModuleFileName(NULL, chpath,sizeof(chpath));
	(strrchr(chpath, _T('\\')))[1] = 0;
	file = chpath;
	file += OUTPUT_FILE;

	COleException pe;  
    if (!m_ecApp.CreateDispatch(_T("Excel.Application"), &pe))  
    {  
        AfxMessageBox(_T("Application创建失败，请确保安装了word 2000或以上版本!"), MB_OK|MB_ICONWARNING);  
        pe.ReportError();  
        throw &pe;  
        return FALSE;  
    }
	m_ecApp.put_AlertBeforeOverwriting(FALSE);

	m_ecBooks = m_ecApp.get_Workbooks();  
    if (!m_ecBooks.m_lpDispatch)   
    {  
        AfxMessageBox(_T("WorkBooks创建失败!"), MB_OK|MB_ICONWARNING);  
        return FALSE;  
    } 

	COleVariant VOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);  
    m_ecBook = m_ecBooks.Add(VOptional);
	if(!m_ecBook.m_lpDispatch)   
    {  
        AfxMessageBox(_T("文件打开失败获取失败!"), MB_OK|MB_ICONWARNING);  
        return FALSE;
    } 

	m_ecSheets = m_ecBook.get_Sheets();  
    if(!m_ecSheets.m_lpDispatch)   
    {
        AfxMessageBox(_T("WorkSheets为空!"), MB_OK|MB_ICONWARNING);  
        return FALSE;
    }

	short firstSheet = 1;
	m_ecSheet = m_ecSheets.get_Item(COleVariant(firstSheet));  
    if(!m_ecSheet.m_lpDispatch)   
    {  
        AfxMessageBox("WorkSheet为空!", MB_OK|MB_ICONWARNING);  
        return FALSE;  
    }
	m_ecSheet.put_Name(_T("储罐计算结果"));

	CRange range = m_ecSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("I1")));
	range.Merge(VOptional);

	m_range = m_ecSheet.get_Cells();
    if(!m_range.m_lpDispatch)   
    {  
        AfxMessageBox("m_range为空!", MB_OK|MB_ICONWARNING);  
        return FALSE;  
    }

	//填写配置信息
	m_range.put_Item(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("容器配置信息:")));

	m_range.put_Item(COleVariant((long)2),COleVariant((long)1),COleVariant(_T("容器类型")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)2),COleVariant(_T("容积(m3)")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)3),COleVariant(_T("计算压力(MPa)")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)4),COleVariant(_T("设计温度(℃)"))); 
	m_range.put_Item(COleVariant((long)2),COleVariant((long)5),COleVariant(_T("焊接接头系数")));	
	m_range.put_Item(COleVariant((long)2),COleVariant((long)6),COleVariant(_T("材料")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)7),COleVariant(_T("厚度负偏差(mm)")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)8),COleVariant(_T("腐蚀裕量(mm)")));
	m_range.put_Item(COleVariant((long)2),COleVariant((long)9),COleVariant(_T("安装形式")));

	m_range.put_Item(COleVariant((long)3),COleVariant((long)1),COleVariant(_T(containerTypeStr[m_conType])));
	str.Format(_T("%.2f"), m_volume);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)2),COleVariant(str));
	str.Format(_T("%.2f"), m_pressure);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)3),COleVariant(str));
	str.Format(_T("%d"), m_temperature);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)4),COleVariant(str)); 
	str.Format(_T("%.2f"), m_coefficient);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)5),COleVariant(str));	 
	m_range.put_Item(COleVariant((long)3),COleVariant((long)6),COleVariant(_T(steelNumStrMap[m_conMetarial].steelNumName)));
	str.Format(_T("%.2f"), m_thickNegWindage);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)7),COleVariant(str));
	str.Format(_T("%.2f"), m_cauterization);
	m_range.put_Item(COleVariant((long)3),COleVariant((long)8),COleVariant(str));
	m_range.put_Item(COleVariant((long)3),COleVariant((long)9),COleVariant(_T(conInstallTypeStr[m_conInstType])));	



	range = m_ecSheet.get_Range(COleVariant(_T("A6")),COleVariant(_T("G6")));
	range.Merge(VOptional);
	m_range.put_Item(COleVariant((long)6),COleVariant((long)1),COleVariant(_T("容器计算结果:")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)1),COleVariant(_T("序号")));
	m_range.put_Item(COleVariant((long)7),COleVariant((long)2),COleVariant(_T("筒体内径（mm）")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)3),COleVariant(_T("筒体壁厚（mm）")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)4),COleVariant(_T("筒体高度（mm）")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)5),COleVariant(_T("封头壁厚（mm）")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)6),COleVariant(_T("壳体总高（mm）")));
    m_range.put_Item(COleVariant((long)7),COleVariant((long)7),COleVariant(_T("壳体重量（kg）")));

	int i = 8;
	for (list <conCaclResultItem_t>::iterator iter = m_conCaclResults.begin();
		iter != m_conCaclResults.end(); iter++) {

		str.Format(_T("%d"), (i-7));
		m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(str));
		str.Format(_T("%.2f"), iter->diameterIn);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(str));
		str.Format(_T("%.2f"), iter->conThick);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(str));
		str.Format(_T("%.2f"), iter->conHight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(str));
		str.Format(_T("%.2f"), iter->capThick);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(str));
		str.Format(_T("%.2f"), iter->totalHight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)6),COleVariant(str));
		str.Format(_T("%.2f"), iter->weight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)7),COleVariant(str));

		i++;
	}



    m_ecBook.SaveAs(  
        COleVariant(file),
        COleVariant((long)51),//该参数使得保存的时候使用兼容模式，避免了使用excel2003打开的时候弹出提示框  
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

	AfxMessageBox("储罐计算结果保存成功!", MB_OK|MB_ICONINFORMATION);  

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

	if (temperature < -20 || temperature > 150) {
		return FALSE;
	}

	return TRUE;
}
	

