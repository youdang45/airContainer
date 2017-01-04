#include "stdafx.h"
#include <algorithm>
#include "CExlHeader.h"
#include "CPipeHole.h"
#include "CContainerInfo.h"

#define PH_OUTPUT_FILE "�ӹ��˿׼�����.xlsx"

holeInDiameter_t holeSize[] = 
{
	{0,	  0},
	{280, 380},
	{300, 400}
};

CString holeSizeStr[] = 
{
	"Invalid",
	"280*380",
	"300*400"
};
pipeCalcResult_t::pipeCalcResult_t ()
{
	minThick = 0;
	baseThick = 0;
	minExtInHeight = 0;
	minWidth = 0;
	addPressThick = 0;
	recommendAddPress = COMMENT_NONE;
}

CPipeHole *CPipeHole::m_pInstance = NULL;

CPipeHole::CPipeHole()
{
	m_holeConfig.material = SteelNumberNone;
}

CPipeHole::~CPipeHole()
{
}

CPipeHole * CPipeHole::GetInstance()  
{  
	if(m_pInstance == NULL)  //�ж��Ƿ��һ�ε���  
		m_pInstance = new CPipeHole();  
	return m_pInstance;  
}

void CPipeHole::addPipeConfig(pipeConfig_t &pipeConfig)
{
	pipeConfig_t pc = pipeConfig;
	m_configList.push_back(pc);
}

list <pipeConfig_t> &CPipeHole::getPipeConfigList()
{
	return m_configList;
}

void CPipeHole::addPipeCalcResult(pipeCalcResult_t &pipeCalcResult)
{
	pipeCalcResult_t pr = pipeCalcResult;
	m_pipeCalcResultList.push_back(pr);
}

void CPipeHole::setHoleConfig(holeConfig_t &holeConfig)
{
	m_holeConfig = holeConfig;
}
holeConfig_t &CPipeHole::getHoleConfig()
{
	return m_holeConfig;
}
void CPipeHole::setHoleCalcResult(holeCalcResult_t &holeResult)
{
	m_holeResult = holeResult;
}
list <pipeCalcResult_t> &CPipeHole::getPipeCalcResults()
{
	return m_pipeCalcResultList;
}

holeCalcResult_t &CPipeHole::getHoleCalcResult()
{
	return m_holeResult;
}

void CPipeHole::reset()
{
	m_configList.clear();
	m_pipeCalcResultList.clear();
	memset(&m_holeConfig, 0, sizeof(m_holeConfig));
	memset(&m_holeResult, 0, sizeof(m_holeResult));
}

bool CPipeHole::save()
{
	bool ret = FALSE;
	CContainerInfo *containerInfo = CContainerInfo::GetInstance();
	conConfig_t config;
	conCaclResultItem_t &conItem = containerInfo->getSelectItem();
	CString file;
	CString str;
	CRange rangeTmp;
	int i = 1;
	CString row, left, right;

	TCHAR chpath[100] = {0}; 
	GetModuleFileName(NULL, chpath,sizeof(chpath));
	(strrchr(chpath, _T('\\')))[1] = 0;
	file = chpath;
	file += PH_OUTPUT_FILE;

	COleException pe;
    if (!m_ecApp.CreateDispatch(_T("Excel.Application"), &pe))  
    {  
        AfxMessageBox(_T("Application����ʧ�ܣ���ȷ����װ��word 2000�����ϰ汾!"), MB_OK|MB_ICONWARNING);  
        pe.ReportError();  
        throw &pe;  
        return FALSE;  
    }
	m_ecApp.put_AlertBeforeOverwriting(FALSE);

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

	short sndSheet = 1;
	m_ecSheet = m_ecSheets.get_Item(COleVariant(sndSheet));  
    if(!m_ecSheet.m_lpDispatch)   
    {  
        AfxMessageBox("WorkSheetΪ��!", MB_OK|MB_ICONWARNING);  
        return FALSE;  
    }
	m_ecSheet.put_Name(_T("�����˿׼�����"));


	m_range = m_ecSheet.get_Cells();
    if(!m_range.m_lpDispatch)   
    {  
        AfxMessageBox("rangeΪ��!", MB_OK|MB_ICONWARNING);  
        return FALSE;  
    }


	//��д������Ϣ
	row.Format(_T("%d"), i);
	left = _T("A") + row;
	right = _T("H")+row;
	rangeTmp = m_ecSheet.get_Range(COleVariant(left),COleVariant(right));
	rangeTmp.Merge(VOptional);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(_T("��ѡ������Ϣ:")));

	i++;
	m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(_T("�ݻ�(m3)")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(_T("����ѹ��(MPa)")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(_T("����")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(_T("����¶�(��)")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(_T("���ӽ�ͷϵ��")));

	i++;
	ret = containerInfo->getContainerConfig(config);
	if (ret == FALSE) {
		//Add Error Msg;
		exit(1);
	}
	str.Format(_T("%.2f"), config.volume);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(str));
	str.Format(_T("%.2f"), config.pressure);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(str));
	str = getSteelNameByNum(config.conMetarial);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(str));
	str.Format(_T("%d"), config.temperature);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(str));
	str.Format(_T("%.2f"), config.coefficient);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(str));

	i++;
	m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(_T("���")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(_T("Ͳ���ھ���mm��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(_T("Ͳ��ں�mm��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(_T("Ͳ��߶ȣ�mm��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(_T("��ͷ�ں�mm��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)6),COleVariant(_T("��ͷֱ�߸߶ȣ�mm��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)7),COleVariant(_T("�����ܸߣ�mm��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)8),COleVariant(_T("����������kg��")));

	i++;
	int selectNo = containerInfo->getContainerSelectedNum();
	str.Format(_T("%d"), (selectNo));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(str));
	str.Format(_T("%.2f"), conItem.diameterIn);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(str));
	str.Format(_T("%.2f"), conItem.conThick);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(str));
	str.Format(_T("%.2f"), conItem.conHight);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(str));
	str.Format(_T("%.2f"), conItem.capThick);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(str));
	str.Format(_T("%.2f"), conItem.weight);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)6),COleVariant(str));
	str.Format(_T("%.2f"), conItem.totalHight);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)7),COleVariant(str));
	str.Format(_T("%.2f"), conItem.capHight);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)8),COleVariant(str));


	i += 2;
	i++;
	row.Format(_T("%d"), i);
	left = _T("A") + row;
	right = _T("M")+row;
	rangeTmp = m_ecSheet.get_Range(COleVariant(left),COleVariant(right));
	rangeTmp.Merge(VOptional);
	m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(_T("�ӹ���Ϣ:")));
	i++;
    m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(_T("���")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(_T("���ò�ǿ")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(_T("�ӹܲ���")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(_T("�ӹ��⾶��mm��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(_T("�ӹ�Ͷ�Ϻ�ȣ�mm��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)6),COleVariant(_T("��������")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)7),COleVariant(_T("�ӹ���СͶ�Ϻ��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)8),COleVariant(_T("�ӹ���С����߶�")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)9),COleVariant(_T("�ӹ���С���쳤��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)10),COleVariant(_T("��ѡ�ñ�׼��")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)11),COleVariant(_T("��ǿ��С���")));
    m_range.put_Item(COleVariant((long)i),COleVariant((long)12),COleVariant(_T("��ǿȦ��С���")));
	m_range.put_Item(COleVariant((long)i),COleVariant((long)13),COleVariant(_T("��ע")));


	i++;
	int j = 1;
	list <pipeConfig_t>::iterator iter1 = m_configList.begin();
	list <pipeCalcResult_t>::iterator iter2 = m_pipeCalcResultList.begin();
	for (;
		(iter1 != m_configList.end()) && (iter2 != m_pipeCalcResultList.end());
		iter1++, iter2++) {

		str.Format(_T("%d"), j);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(str));
		str = (iter1->isAddStress) ? _T("��") : _T("��");
		m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(str));
		str = steelNumStrMap[iter1->material].steelNumName;
		m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(str));
		str.Format(_T("%.2f"), iter1->Do);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(str));
		str.Format(_T("%.2f"), iter1->thick);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(str));
		str = (iter1->extendIn) ? _T("��") : _T("��");
		m_range.put_Item(COleVariant((long)i),COleVariant((long)6),COleVariant(str));
		str.Format(_T("%.2f"), iter2->minThick);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)7),COleVariant(str));
		str.Format(_T("%.2f"), iter2->minExtInHeight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)8),COleVariant(str));
		str.Format(_T("%.2f"), iter2->minExtOutHeight);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)9),COleVariant(str));
		str = iter2->GBStr;
		m_range.put_Item(COleVariant((long)i),COleVariant((long)10),COleVariant(str));
		str.Format(_T("%.2f"), iter2->minWidth);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)11),COleVariant(str));
		str.Format(_T("%.2f"), iter2->addPressThick);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)12),COleVariant(str));
		switch(iter2->recommendAddPress) {
			case COMMENT_ADD_PRESS:
				str = _T("������ò�ǿȦ���в�ǿ");
				break;
			case COMMENT_NOT_ADD_PRESS:
				str = _T("����Ҫ���в�ǿ");
				break;
			default:
				str = "";
				break;
		}
		m_range.put_Item(COleVariant((long)i),COleVariant((long)13),COleVariant(str));

		i++;
		j++;
	}

	if (m_holeConfig.material != SteelNumberNone) {
		i += 2;//���п������ڷָ��˿���Ϣ
		row.Format(_T("%d"), i);
		left = _T("A") + row;
		right = _T("M")+row;
		rangeTmp = m_ecSheet.get_Range(COleVariant(left),COleVariant(right));
		rangeTmp.Merge(VOptional);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(_T("�˿���Ϣ:")));


		i++;
		m_range.put_Item(COleVariant((long)i),COleVariant((long)1),COleVariant(_T("�˿׳ߴ�")));
		m_range.put_Item(COleVariant((long)i),COleVariant((long)2),COleVariant(_T("�˿ײ���")));
		m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(_T("�˿�������(mm)")));
		m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(_T("�˿���С���볤��")));
		m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(_T("�˿���С�������")));

		i++;
		m_range.put_Item(COleVariant((long)i),COleVariant((long)1),
						COleVariant(holeSizeStr[m_holeConfig.size]));
		m_range.put_Item(COleVariant((long)i),COleVariant((long)2),
						COleVariant(steelNumStrMap[m_holeConfig.material].steelNumName));
		str.Format(_T("%.2f"), m_holeResult.thickness);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)3),COleVariant(str));
		str.Format(_T("%.2f"), m_holeResult.extInLenght);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)4),COleVariant(str));
		str.Format(_T("%.2f"), m_holeResult.extOutLenght);
		m_range.put_Item(COleVariant((long)i),COleVariant((long)5),COleVariant(str));


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

	AfxMessageBox("�ӹ��˿׼���������ɹ�!", MB_OK|MB_ICONINFORMATION);  

	return TRUE;
}


void CPipeHole::closeExlApp()
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
holeSize_t getHoleSizeByStr(CString str)
{
	if (str == "280*380") {
		return ES_280_380;
	} else if (str == "300*400") {
		return ES_300_400;
	} else {
		return ES_INVALID;
	}
}
