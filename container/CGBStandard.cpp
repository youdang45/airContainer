#include "stdafx.h"
#include "CGBStandard.h"
#include "CExlHeader.h"

static const float g_defPhi = 0.85;


SteelNumStringMap_t steelNumStrMap[] = 
{
	{SteelNumberNone, _T("None")},
	{Q245R, _T("Q245R")},
	{Q345R, _T("Q345R")},
	{S30408, _T("S30408")},
	{Q235B, _T("Q235B")},
	{E_10, _T("10")},
	{E_20, _T("20")},
	{E_16Mn, _T("16Mn")},
};

GBStandardStrMap_t GBStandardStrMap[] = 
{
	{GBStandardNoNone, _T("未知标准")},
	{GB_T_8163, _T("GB/T 8163")},
	{GB_9948, _T("GB 9948")},
	{GB_6479, _T("GB 6479")},
	{GB_T_14976, _T("GB/T 14976")},
};

tempScopeMap1_t tempScopeMap1[] = 
{
	{0,    20, TEMP_LEVEL1_0_20},
	{20,  100, TEMP_LEVEL1_20_100},
	{100, 150, TEMP_LEVEL1_100_150},
	{150, 200, TEMP_LEVEL1_150_200},
	{200, 250, TEMP_LEVEL1_200_250},
	{250, 300, TEMP_LEVEL1_250_300},
	{300, 350, TEMP_LEVEL1_300_350},
	{350, 400, TEMP_LEVEL1_350_400},
	{400, 425, TEMP_LEVEL1_400_425},
	{425, 450, TEMP_LEVEL1_425_450},
	{450, 475, TEMP_LEVEL1_450_475}
};

tempScopeMap2_t tempScopeMap2[] = 
{
	{0,    20, TEMP_LEVEL2_0_20},
	{20,  100, TEMP_LEVEL2_20_100},
	{100, 150, TEMP_LEVEL2_100_150},
	{150, 200, TEMP_LEVEL2_150_200},
	{200, 250, TEMP_LEVEL2_200_250},
	{250, 300, TEMP_LEVEL2_250_300},
	{300, 350, TEMP_LEVEL2_300_350},
	{350, 400, TEMP_LEVEL2_350_400},
	{400, 450, TEMP_LEVEL2_400_450},
	{450, 500, TEMP_LEVEL2_450_500},
	{500, 525, TEMP_LEVEL2_500_525},
	{525, 550, TEMP_LEVEL2_525_550},
	{550, 575, TEMP_LEVEL2_550_575},
	{575, 600, TEMP_LEVEL2_575_600},
	{600, 625, TEMP_LEVEL2_600_625},
	{625, 650, TEMP_LEVEL2_625_650},
	{650, 675, TEMP_LEVEL2_650_675},
	{675, 700, TEMP_LEVEL2_675_700},
};

TempLevel_t steelLevelMap[] =
{
    TEMP_LEVEL_INVALID,
	TEMP_LEVEL1,   //[Q245R]  
	TEMP_LEVEL1,   //[Q345R]  
    TEMP_LEVEL2,   //[S30408] 
	TEMP_LEVEL1,   //[E_10]   
	TEMP_LEVEL1,   //[E_20]   
	TEMP_LEVEL1,   //[E_16Mn]
};

unsigned short maxTemp[] = {0, 475, 475, 700, 475, 475, 475};

CGBItemLine::CGBItemLine()
{
}

CGBItemLine::~CGBItemLine()
{
}

bool CGBItemLine::getStressScope(SteelNumber_t  material,
								 short temp,
								 stressScope_t &stressScope)
{
	if (material == S30408) {
		for (int i = 0; i < sizeof(tempScopeMap2)/sizeof(tempScopeMap2_t); i++) {
			if ((temp > tempScopeMap2[i].min) && (temp <= tempScopeMap2[i].max)) {
				stressScope.beginTemp = tempScopeMap2[i].min;
				stressScope.endTemp = tempScopeMap2[i].max;
				tempLevel2_t tempLevel = tempScopeMap2[i].level;
				stressScope.beginStress = stress[tempLevel];
				stressScope.endStress = stress[tempLevel];
				return TRUE;
			}
		}
	} else {
		for (int i = 0; i < sizeof(tempScopeMap1)/sizeof(tempScopeMap1_t); i++) {
			if ((temp > tempScopeMap1[i].min) && (temp <= tempScopeMap1[i].max)) {
				stressScope.beginTemp = tempScopeMap1[i].min;
				stressScope.endTemp = tempScopeMap1[i].max;
				tempLevel1_t tempLevel = tempScopeMap1[i].level;
				stressScope.beginStress = stress[tempLevel-1];
				stressScope.endStress = stress[tempLevel];
				return TRUE;
			}
		}
	}

	return FALSE;
}

unsigned int CGBItemLine::getLowestTempStress()
{
	return stress[0];
}

CGBStandard *CGBStandard::m_pInstance = NULL;
CGBStandard::CGBStandard()
{
	TCHAR chpath[100] = {0}; 
	GetModuleFileName(NULL, chpath,sizeof(chpath));
	(strrchr(chpath, _T('\\')))[1] = 0;
	m_stdFile = chpath;
	m_stdFile += GB_STANDARD_FILE;
}

CGBStandard::~CGBStandard()
{
}

CGBStandard * CGBStandard::GetInstance()  
{  
    if(m_pInstance == NULL)  //判断是否第一次调用  
        m_pInstance = new CGBStandard();  
    return m_pInstance;  
}

bool CGBStandard::LoadGBStandard()
{
	bool ret = FALSE;
	CString fileName = m_stdFile;

    if (!AfxOleInit())  
    {
       AfxMessageBox(_T("注册COM出错!"));  
       return FALSE;  
    }

	COleException pe;  
    if (!m_ecApp.CreateDispatch(_T("Excel.Application"), &pe))  
    {  
        AfxMessageBox(_T("Application创建失败，请确保安装了word 2000或以上版本!"), MB_OK|MB_ICONWARNING);  
        pe.ReportError();  
        throw &pe;  
        return FALSE;  
    }
	m_ecBooks = m_ecApp.get_Workbooks();  
    if (!m_ecBooks.m_lpDispatch)   
    {  
        AfxMessageBox(_T("WorkBooks创建失败!"), MB_OK|MB_ICONWARNING);  
        return FALSE;  
    } 

	COleVariant VOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);  
    m_ecBook = m_ecBooks.Open(fileName,
	                      VOptional, VOptional, VOptional, VOptional,
						  VOptional, VOptional, VOptional, VOptional,
						  VOptional, VOptional, VOptional, VOptional,
						  VOptional, VOptional);  
	if(!m_ecBook.m_lpDispatch)   
    {  
        AfxMessageBox(_T("WorkSheet获取失败!"), MB_OK|MB_ICONWARNING);  
        return FALSE;  
    } 

	m_ecSheets = m_ecBook.get_Sheets();  
    if(!m_ecSheets.m_lpDispatch)   
    {  
        AfxMessageBox(_T("WorkSheets为空!"), MB_OK|MB_ICONWARNING);  
        return FALSE;  
    }
    CString sheetName;
	for (short i = 1; i < m_ecSheets.get_Count(); i++) 
	{
		m_ecSheet = m_ecSheets.get_Item(COleVariant(i));
		sheetName = m_ecSheet.get_Name();
		if (sheetName == "钢板") {
			//
			ret = LoadPlateStd(m_ecSheet);
		} else if (sheetName == "钢管") {
			ret = LoadPipeStd(m_ecSheet);
		}

		if (FALSE == ret){
			//Add Error Info
			return FALSE;
		}

	}

	closeExlApp();
	return TRUE;
}

bool CGBStandard::LoadPlateStd(CWorksheet &ecSheet)
{
	//GBItem_t item;
	COleVariant value;

	//获取使用区域范围
	m_usedRange = ecSheet.get_UsedRange();


	int columnCnt = 0;
	int rowCnt = 0;

	m_columns = m_usedRange.get_Columns();
	columnCnt = m_columns.get_Count();

	m_rows = m_usedRange.get_Rows();
	rowCnt = m_rows.get_Count();

	GBItem_t gbItem;
	CString tmpStr;
	bool itemValid = FALSE;
	gbItem.GBNo = GBStandardNoNone;
	for (int i = 1; i <= rowCnt; i++) {

		//标题行，跳过两行
		m_cell1 = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)1)).pdispVal;
		value = m_cell1.get_Value2();
		if (value.vt == VT_BSTR){
			CString vStr;
			vStr =  value.bstrVal;
			if (vStr == "钢号") {
			    i++;
			    continue;
			}
		}

		//空行自动跳过
		if (value.vt == VT_EMPTY) { //第一格为空
			m_cell2 = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)2)).pdispVal;
			value = m_cell2.get_Value2();
			if (value.vt == VT_EMPTY) {//第二格也为空
				continue;
			}
		}

		//读取钢号
		m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)1)).pdispVal;
		value = m_cell.get_Value2();
		if (value.vt == VT_BSTR) {
			if (itemValid) {
				m_plateGBItemList.push_back(gbItem);
				plateItemReset(gbItem);
			}
			itemValid = TRUE;
			tmpStr = value.bstrVal;
			gbItem.steelNo = getSteelNumberByName(tmpStr);
			if (SteelNumberNone == gbItem.steelNo)
			{
				//Add Error Msg
				return FALSE;
			}
		}

		CGBItemLine gbItemLine;
		bool ret = FALSE;
		float minThick = 0;
		float maxThick = 0;


		//读取厚度
		m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)2)).pdispVal;
		value = m_cell.get_Value2();
		tmpStr = value.bstrVal;
		ret = getThickScopeFromString(tmpStr, minThick, maxThick);
		if (FALSE == ret) {
			//Add Error Info
			return FALSE;
		}

		gbItemLine.thicknessMin = minThick;
		gbItemLine.thicknessMax = maxThick;

		//读取Rm
		m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)3)).pdispVal;
		value = m_cell.get_Value2(); //Warning!!
		gbItemLine.Rm = value.dblVal;

		//读取ReL
		m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)4)).pdispVal;
		value = m_cell.get_Value2();
		gbItemLine.ReL = value.dblVal;

		//读取许用应力
		for (int j = 5; j <= columnCnt; j++) {
			m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)j)).pdispVal;
			value = m_cell.get_Value2();
			if (value.vt == VT_EMPTY) {
				break;
			}
			gbItemLine.stress[j-5] = value.dblVal;
		}
		gbItem.lineList.push_back(gbItemLine);
	}
	
	if (itemValid) {
		m_plateGBItemList.push_back(gbItem);
	}

	return TRUE;
}

bool CGBStandard::LoadPipeStd(CWorksheet &ecSheet)
{
	//GBItem_t item;
	COleVariant value;

	//获取使用区域范围
	m_usedRange = ecSheet.get_UsedRange();


	int columnCnt = 0;
	int rowCnt = 0;

	m_columns = m_usedRange.get_Columns();
	columnCnt = m_columns.get_Count();

	m_rows = m_usedRange.get_Rows();
	rowCnt = m_rows.get_Count();
	GBItem_t gbItem;
	CString tmpStr;
	bool itemValid = FALSE;

	for (int i = 1; i <= rowCnt; i++) {

		//标题行，跳过两行
		m_cell1 = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)1)).pdispVal;
		value = m_cell1.get_Value2();
		if (value.vt == VT_BSTR){
			CString vStr;
			vStr =  value.bstrVal;
			if (vStr == "钢号") {
			    i++;
			    continue;
			}
		}

		//空行自动跳过
		if (value.vt == VT_EMPTY) { //第一格为空
			m_cell2 = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)3)).pdispVal;
			value = m_cell2.get_Value2();
			if (value.vt == VT_EMPTY) {//第二格也为空
				continue;
			}
		}

		//读取钢号
		m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)1)).pdispVal;
		value = m_cell.get_Value2();
		if (value.vt != VT_EMPTY) {
			//将上一条有效项放入list
			if (itemValid) {
				m_pipeGBItemList.push_back(gbItem);
				plateItemReset(gbItem);
				itemValid = FALSE;
			}

			itemValid = TRUE;
			if (value.vt == VT_BSTR) {
				tmpStr = value.bstrVal;
			} else if (value.vt == VT_R8){
				int tmp = 0;
				tmp = value.dblVal;
				tmpStr.Format(_T("%d"),tmp);
			}
			gbItem.steelNo = getSteelNumberByName(tmpStr);
			if (SteelNumberNone == gbItem.steelNo)
			{
				//Add Error Msg
				return FALSE;
			}

			//读取国标号
			m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)2)).pdispVal;
			value = m_cell.get_Value2();
			if (value.vt == VT_BSTR) {
				tmpStr = value.bstrVal;
				gbItem.GBNo = getGBStdNumberByName(tmpStr);
				if (SteelNumberNone == gbItem.steelNo)
				{
					//Add Error Msg
					return FALSE;
				}
			}
		}


		CGBItemLine gbItemLine;
		bool ret = FALSE;

		//读取厚度
		float minThick = 0;
		float maxThick = 0;

		m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)3)).pdispVal;
		value = m_cell.get_Value2();
		tmpStr = value.bstrVal;
		ret = getThickScopeFromString(tmpStr, minThick, maxThick);
		if (FALSE == ret) {
			//Add Error Info
			return FALSE;
		}

		gbItemLine.thicknessMin = minThick;
		gbItemLine.thicknessMax = maxThick;

		//读取Rm
		m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)4)).pdispVal;
		value = m_cell.get_Value2(); //Warning!!
		gbItemLine.Rm = value.dblVal;

		//读取ReL
		m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)5)).pdispVal;
		value = m_cell.get_Value2();
		gbItemLine.ReL = value.dblVal;

		//读取许用应力
		for (int j = 6; j <= columnCnt; j++) {
			m_cell = m_usedRange.get_Item(COleVariant((long)i),COleVariant((long)j)).pdispVal;
			value = m_cell.get_Value2();
			if (value.vt == VT_EMPTY) {
				break;
			}
			gbItemLine.stress[j-6] = value.dblVal;
		}
		gbItem.lineList.push_back(gbItemLine);
	}
	
	if (itemValid) {
		m_pipeGBItemList.push_back(gbItem);
	}

	return TRUE;
}

void CGBStandard::closeExlApp()
{
	COleVariant vTrue((short)TRUE),      
                vFalse((short)FALSE),  
                vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);  
 
    m_ecApp.Quit();
	if (m_cell2.m_lpDispatch)
		m_cell2.ReleaseDispatch();
	if (m_cell1.m_lpDispatch)
		m_cell1.ReleaseDispatch();
	if (m_cell.m_lpDispatch)
		m_cell.ReleaseDispatch();
	if (m_rows.m_lpDispatch)
		m_rows.ReleaseDispatch();
	if (m_columns.m_lpDispatch)
		m_columns.ReleaseDispatch();
    if(m_usedRange.m_lpDispatch)  
        m_usedRange.ReleaseDispatch();  
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


bool CGBStandard::getPlateGBItem(SteelNumber_t &conMaterial, GBItem_t &item)
{
	for (list<GBItem_t>::iterator iter = m_plateGBItemList.begin();
		iter != m_plateGBItemList.end(); iter++ ) {
		if ((*iter).steelNo == conMaterial) {
			item = (*iter);
			return TRUE;
		}
	}
	return FALSE;
}

bool CGBStandard::getPlateStressScope(SteelNumber_t &material, float Di,
						short temp, stressScope_t &stressScope)
{
	bool ret = TRUE;

	for (list<GBItem_t>::iterator iter1 = m_plateGBItemList.begin();
		iter1 != m_plateGBItemList.end(); iter1++) {
		if (iter1->steelNo == material) {
			for (list<CGBItemLine>::iterator iter2 = iter1->lineList.begin();
				iter2 != iter1->lineList.end(); iter2++) {
				if ((iter2->thicknessMin < Di)
					&& (Di <= iter2->thicknessMax)) {
					ret = iter2->getStressScope(material, temp, stressScope);
					return ret;
				}
			}
		}
	}

	return FALSE;
}

float CGBStandard::getPlateMaxThick(SteelNumber_t &material)
{
	float thick = 0;
	for (list<GBItem_t>::iterator iter1 = m_plateGBItemList.begin();
		iter1 != m_plateGBItemList.end(); iter1++) {
			if (iter1->steelNo == material) {
				list<CGBItemLine>::iterator iter2 = iter1->lineList.end();
				iter2--;
				thick = iter2->thicknessMax;
				break;
			}
	}
	return thick;
}

list <GBItem_t> CGBStandard::GetPipeItemList(SteelNumber_t &material)
{
	list <GBItem_t> pipeGBList;

	for (list<GBItem_t>::iterator iter = m_pipeGBItemList.begin();
		iter != m_pipeGBItemList.end(); iter++) {
		if (iter->steelNo == material) {
			pipeGBList.push_back(*iter);

		}
	}
	return pipeGBList;
}

float calcSelectedStress(stressScope_t &sS, short temp)
{
	float rate = (sS.endTemp - temp )/(sS.endTemp - sS.beginTemp);
	return (sS.endStress + rate*(sS.beginStress - sS.endStress));
}

SteelNumber_t getSteelNumberByName(CString steelName)
{
	for (int i = 0; i < sizeof(steelNumStrMap)/sizeof(SteelNumStringMap_t); i++) {
		if (steelNumStrMap[i].steelNumName == steelName) {
			return steelNumStrMap[i].steelNum;
		}
	}
	return SteelNumberNone;
}

CString getSteelNameByNum(SteelNumber_t num)
{
	for (int i = 0; i < sizeof(steelNumStrMap)/sizeof(SteelNumStringMap_t); i++) {
		if (steelNumStrMap[i].steelNum == num) {
			return steelNumStrMap[i].steelNumName;
		}
	}
	return _T("");

}

GBStandardNumber_t getGBStdNumberByName(CString gbName)
{
	for (int i = 0; i < sizeof(GBStandardStrMap)/sizeof(GBStandardStrMap_t); i++) {
		if (GBStandardStrMap[i].GBNumName == gbName) {
			return GBStandardStrMap[i].GBNum;
		}
	}
	return GBStandardNoNone;
}

CString getGBStdNameByNum(GBStandardNumber_t gbNum)
{
	for (int i = 0; i < sizeof(GBStandardStrMap)/sizeof(GBStandardStrMap_t); i++) {
		if (GBStandardStrMap[i].GBNum == gbNum) {
			return GBStandardStrMap[i].GBNumName;
		}
	}
	return _T("");
}


bool getThickScopeFromString(CString str, float &min, float &max)
{
	int nLen = str.GetLength();
	bool splitFound = FALSE;
    CString strNumMin;
	CString strNumMax;

    for(int i=0; i<nLen; i++)
    {
		if (!splitFound) {
			if((str.GetAt(i) >= '0' && str.GetAt(i) <= '9')
				|| (str.GetAt(i) == '.'))
				strNumMin += str.GetAt(i);
			if (str.GetAt(i) == '~'){
				splitFound = TRUE;
			}
		} else {
			if((str.GetAt(i) >= '0' && str.GetAt(i) <= '9')
				|| (str.GetAt(i) == '.'))
				strNumMax += str.GetAt(i);
		}
	}

	if (!splitFound) {
		strNumMax = strNumMin;
		strNumMin = "0";
	}
	if ((strNumMin == "") || (strNumMax == "")) {
		//Add Error Log
		return FALSE;
	}

	min =  atof(strNumMin);
	max =  atof(strNumMax);

	return TRUE;
}

void plateItemReset(GBItem_t &gbItem)
{
	gbItem.GBNo = GBStandardNoNone;
	gbItem.steelNo = SteelNumberNone;
	gbItem.lineList.clear();
}

float getPipeMaxThick(GBItem_t &gbItem)
{
	list <CGBItemLine>::iterator iter = gbItem.lineList.end();
	iter--;

	return iter->thicknessMax;
}