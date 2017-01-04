// container2Dlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "CExlHeader.h"
#include "container.h"
#include "containerDlg.h"
#include "CGBStandard.h"
#include "CContainerInfo.h"
#include "CPipeHole.h"
#include "CPipeHoleModel.h" //To be updated

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// Ccontainer2Dlg �Ի���




Ccontainer2Dlg::Ccontainer2Dlg(CWnd* pParent /*=NULL*/)
	: CDialog(Ccontainer2Dlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_exit = TRUE;
}

void Ccontainer2Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST, m_resultList);
	DDX_Control(pDX, IDC_EDIT_V, m_volumeEdit);
	DDX_Control(pDX, IDC_EDIT_P, m_pressEdit);
	DDX_Control(pDX, IDC_COMBO_M, m_materialCombo);
	DDX_Control(pDX, IDC_EDIT_T, m_temperatureEdit);
	DDX_Control(pDX, IDC_COMBO_C, m_coefCombo);
	DDX_Control(pDX, IDC_COMBO_T, m_containerType);
	DDX_Control(pDX, IDC_COMBO_IT, m_installType);
	DDX_Control(pDX, IDC_COMBO_STEP, m_thickStep);
	DDX_Control(pDX, IDC_EDIT_HEIGHT_MIN, m_heightMin);
	DDX_Control(pDX, IDC_EDIT_HEIGHT_MAX, m_heightMax);
	DDX_Control(pDX, IDC_EDIT_HDR_MIN, m_heightDiRateMin);
	DDX_Control(pDX, IDC_EDIT_OUTPUT_NUM, m_outputNum);
	DDX_Control(pDX, IDC_EDIT_HDR_MAX, m_heightDiRateMax);
	DDX_Control(pDX, IDC_EDIT_CA, m_cauterization);
}

BEGIN_MESSAGE_MAP(Ccontainer2Dlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDCANCEL, &Ccontainer2Dlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_RESET, &Ccontainer2Dlg::OnBnClickedReset)
	ON_BN_CLICKED(IDC_CAN_CACL, &Ccontainer2Dlg::OnBnClickedConCacl)
	ON_BN_CLICKED(ID_PIPE_CACL, &Ccontainer2Dlg::OnBnClickedPipeCacl)
	ON_CBN_SELCHANGE(IDC_COMBO_M, &Ccontainer2Dlg::OnCbnSelchangeComboM)
	ON_BN_CLICKED(IDSAVE, &Ccontainer2Dlg::OnBnClickedSave)
END_MESSAGE_MAP()


// Ccontainer2Dlg ��Ϣ�������

BOOL Ccontainer2Dlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	initCombo();
	setConConfigInfo();
    initResultList();
	setDefaultConfig();
	fillContainerResultList();

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void Ccontainer2Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void Ccontainer2Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR Ccontainer2Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//�Լ�������г�ʼ
void Ccontainer2Dlg::initResultList()
{
	LVCOLUMN lvColumn;
	CString strText;

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = 80;
    strText = _T("���");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_resultList.InsertColumn(0, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
	strText = _T("Ͳ���ھ���mm��");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_resultList.InsertColumn(1, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("Ͳ��ں�mm��");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_resultList.InsertColumn(2, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("Ͳ��߶ȣ�mm��");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_resultList.InsertColumn(3, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("��ͷ�ں�mm��");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_resultList.InsertColumn(4, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth + 50;
    strText = _T("��ͷֱ�߸߶ȣ�mm��");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_resultList.InsertColumn(5, &lvColumn);
	
	lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("�����ܸߣ�mm��");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_resultList.InsertColumn(6, &lvColumn);
	
	lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("����������kg��");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_resultList.InsertColumn(7, &lvColumn);

    m_resultList.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

	return;
}

void Ccontainer2Dlg::initCombo()
{
	m_materialCombo.AddString(_T("Q245R"));
	m_materialCombo.AddString(_T("Q345R"));
	m_materialCombo.AddString(_T("S30408"));
	m_materialCombo.AddString(_T("Q235B"));

	m_coefCombo.AddString(_T("0.85"));
	m_coefCombo.AddString(_T("1.00"));

	m_containerType.AddString(_T(containerTypeStr[CONTAINER_TYPE_GENERAL]));
	m_containerType.AddString(_T(containerTypeStr[CONTAINER_TYPE_SIMPLE]));

	m_installType.AddString(_T(conInstallTypeStr[CON_INST_TYPE_UPRIGHT]));
	m_installType.AddString(_T(conInstallTypeStr[CON_INST_TYPE_HORIZANTAL]));

	m_thickStep.AddString(_T("0.5"));
	m_thickStep.AddString(_T("0.25"));

}

void Ccontainer2Dlg::OnBnClickedCancel()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	OnCancel();
}

void Ccontainer2Dlg::OnBnClickedPipeCacl()
{
	CString str;
	CContainerInfo *containerInfo;
	POSITION ps = m_resultList.GetFirstSelectedItemPosition();
	int selectLine = m_resultList.GetNextSelectedItem(ps);
	str = m_resultList.GetItemText(selectLine, 0);
	containerInfo = CContainerInfo::GetInstance();
	containerInfo->setContainerSelectedNum(_ttoi (str));	

	OnCancel();
	m_exit = FALSE;

}

void Ccontainer2Dlg::OnBnClickedReset()
{
	GetDlgItem(IDC_EDIT_V)->SetWindowText(_T(""));
	GetDlgItem(IDC_EDIT_P)->SetWindowText(_T(""));
	GetDlgItem(IDC_EDIT_T)->SetWindowText(_T(""));
	m_coefCombo.SetCurSel(0);
	m_resultList.DeleteAllItems();

	resetConfig();
}

void Ccontainer2Dlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	OnOK();
}


//����Ͳ��
void Ccontainer2Dlg::OnBnClickedConCacl()
{
	CContainerInfo *containerInfo;
	containerInfo = CContainerInfo::GetInstance();

	m_resultList.DeleteAllItems();
	resetConfig();

	containerType_t conType = CONTAINER_TYPE_INVALID;
	float volume = 0;           //�ݻ�
	float pressure = 0;         //����ѹ��
	short temperature = 0;      //����¶�
	float coefficient = 0;      //����ϵ��
	SteelNumber_t conMetarial = SteelNumberNone;  //Ͳ�����
	float thickNegWindage = 0;  //��ȸ�ƫ��
	float cauterization = 0;    //��ʴԣ��
	float thickStep = 0;        //�ں񲽳�
	float DiStep = 0;           //�ھ����㲽��
	conInstallType_t installType = CON_INST_TYPE_INVALID;
	conConfig_t config;
	float heightMin = 0;
	float heightMax = 0;
	float heightDiRateMin = 0;
	float heightDiRateMax = 0;
	unsigned int outputNum = 0;
	CString str;	
	bool checkRet = TRUE;

	m_containerType.GetLBText(m_containerType.GetCurSel(),str);
	conType = getContainerTypeByStr(str);

	GetDlgItemText(IDC_EDIT_V, str);
	volume =  atof(str);

	GetDlgItemText(IDC_EDIT_P, str);
	pressure =  atof(str);

	GetDlgItemText(IDC_EDIT_T, str);
	temperature = _ttoi(str);

	GetDlgItemText(IDC_COMBO_C, str);
	coefficient = atof(str);

	m_materialCombo.GetLBText(m_materialCombo.GetCurSel(),str);
	conMetarial = getSteelNumberByName(str);

	GetDlgItemText(IDC_EDIT_TNW, str);
	thickNegWindage = atof(str);

	GetDlgItemText(IDC_EDIT_CA, str);
	cauterization = atof(str);

	GetDlgItemText(IDC_COMBO_STEP, str);
	thickStep = atof(str);

	GetDlgItemText(IDC_EDIT_STEP, str);
	DiStep = atof(str);

	GetDlgItemText(IDC_COMBO_IT, str);
	installType = getConInstallTypeByStr(str);

	GetDlgItemText(IDC_EDIT_HEIGHT_MIN, str);
	heightMin = atof(str);

	GetDlgItemText(IDC_EDIT_HEIGHT_MAX, str);
	heightMax = atof(str);

	GetDlgItemText(IDC_EDIT_LDR_MIN, str);
	heightDiRateMin = atof(str);

	GetDlgItemText(IDC_EDIT_HDR_MAX, str);
	heightDiRateMax = atof(str);
	
	GetDlgItemText(IDC_EDIT_OUTPUT_NUM, str);
	outputNum = atoi(str);

	if (conType == CONTAINER_TYPE_SIMPLE) {
		checkRet = simpleConInputIsValid(pressure, volume, temperature);
		if (FALSE == checkRet) {
			MessageBox("�������������������!");
			return;
		}
	} else {
		if ((conMetarial == Q235B) &&
			(temperature < 20 || temperature > 300 || pressure > 1.6)){
			MessageBox("Q235B�������ѹ��������¶�����!");
			return;
		}
	}

	config.conType = conType; 
	config.volume = volume; 
	config.pressure = pressure; 
	config.temperature = temperature; 
	config.coefficient = coefficient; 
	config.conMetarial = conMetarial; 
	config.thickNegWindage = thickNegWindage;
	config.cauterization = cauterization; 
	config.thickStep = thickStep; 
	config.DiStep = DiStep; 
	config.installType = installType; 
	config.heightMin = heightMin; 
	config.heightMax = heightMax; 
	config.heightDiRateMin = heightDiRateMin;
	config.heightDiRateMax = heightDiRateMax;
	config.outputNum = outputNum; 

	containerInfo->setContainerConfig(config);

	CContainerModel * p_conModel = CContainerModel::GetInstance();
	p_conModel->setDiStep(DiStep);

	//To be updated
	CHoleModel * p_holeModel = CHoleModel::GetInstance();
	p_holeModel->setThickStep(thickStep);

	p_conModel->calcContainer();

	fillContainerResultList();

}

void Ccontainer2Dlg::resetConfig()
{
	CContainerInfo *pConInfo = CContainerInfo::GetInstance();
	pConInfo->reset();

	CPipeHole *pPH = CPipeHole::GetInstance();
	pPH->reset();
}

void Ccontainer2Dlg::setConConfigInfo()
{
	CContainerInfo *containerInfo;
	containerInfo = CContainerInfo::GetInstance();
	containerType_t ct = CONTAINER_TYPE_INVALID;
	conConfig_t config;
	float v = 0, p = 0, c = 0, ca = 0;
	SteelNumber_t m = SteelNumberNone;
	short t = 0;
	CString str;

	bool ret = containerInfo->getContainerConfig(config);

	if (ret == TRUE) {
		ct = config.conType;
		v = config.volume;
		p = config.pressure;
		m = config.conMetarial;
		t = config.temperature;
		c = config.coefficient;
		ca = config.cauterization;

		m_containerType.SetCurSel(ct-1);

		str.Format(_T("%.2f"), v);
		GetDlgItem(IDC_EDIT_V)->SetWindowText(str); 

		str.Format(_T("%.2f"), p);
		GetDlgItem(IDC_EDIT_P)->SetWindowText(str); 

		str.Format(_T("%d"), t);
		GetDlgItem(IDC_EDIT_T)->SetWindowText(str);

		str.Format(_T("%.2f"), c);
		GetDlgItem(IDC_EDIT_P)->SetWindowText(str); 

		if (ca > 0.00001) {
			str.Format(_T("%.2f"), ca);
			GetDlgItem(IDC_EDIT_CA)->SetWindowText(str);
		}

		m_materialCombo.SetCurSel(m-1);

	} else {
		setDefaultConfig();
	}
}


void Ccontainer2Dlg::setDefaultConfig()
{
	CString str;

	str.Format(_T("1.0"));
	GetDlgItem(IDC_EDIT_CA)->SetWindowText(str); 

	str.Format(_T("0.3"));
	GetDlgItem(IDC_EDIT_TNW)->SetWindowText(str); 

	str.Format(_T("0"));
	GetDlgItem(IDC_EDIT_HEIGHT_MIN)->SetWindowText(str); 
	GetDlgItem(IDC_EDIT_HDR_MIN)->SetWindowText(str); 

	str.Format(_T("10"));
	GetDlgItem(IDC_EDIT_HDR_MAX)->SetWindowText(str); 

	str.Format(_T("50"));
	GetDlgItem(IDC_EDIT_OUTPUT_NUM)->SetWindowText(str);

	str.Format(_T("20"));
	GetDlgItem(IDC_EDIT_STEP)->SetWindowText(str);

	m_coefCombo.SetCurSel(0);
	m_containerType.SetCurSel(0);
	m_installType.SetCurSel(0);
	m_thickStep.SetCurSel(0);

}

void Ccontainer2Dlg::fillContainerResultList()
{
	CString str;
	CContainerInfo *containerInfo;
	containerInfo = CContainerInfo::GetInstance();

	list <conCaclResultItem_t> & resultList = containerInfo->getContainerResultList();
	unsigned int outputNum = containerInfo->getOutputNum();
	int max_size = (resultList.size() > outputNum) ? outputNum : resultList.size();

	int i = 0;
	for (list<conCaclResultItem_t>::iterator iter = resultList.begin();
		i < max_size; i++, iter++) {
		str.Format(_T("%d"), (i+1));
		m_resultList.InsertItem(i, str);

		str.Format(_T("%.2f"), iter->diameterIn);
		m_resultList.SetItemText(i, 1, str);

		str.Format(_T("%.2f(%.2f)"), iter->conThick, iter->conCalcThick);
		m_resultList.SetItemText(i, 2, str);

		str.Format(_T("%.2f"), iter->conHight);
		m_resultList.SetItemText(i, 3, str);

		str.Format(_T("%.2f(%.2f)"), iter->capThick, iter->capCalcThick);
		m_resultList.SetItemText(i, 4, str);

		str.Format(_T("%.2f"), iter->capHight);
		m_resultList.SetItemText(i, 5, str);

		str.Format(_T("%.2f"), iter->totalHight);
		m_resultList.SetItemText(i, 6, str);

		str.Format(_T("%.2f"), iter->weight);
		m_resultList.SetItemText(i, 7, str);

	}


}

void Ccontainer2Dlg::OnCbnSelchangeComboM()
{
	SteelNumber_t conMetarial = SteelNumberNone;
	CString str;

	m_materialCombo.GetLBText(m_materialCombo.GetCurSel(),str);
	conMetarial = getSteelNumberByName(str);

	if (conMetarial == S30408) {
		str.Format(_T("0.0"));		
	} else {
		str.Format(_T("1.0"));
	}

	GetDlgItem(IDC_EDIT_CA)->SetWindowText(str);

	return;
}

void Ccontainer2Dlg::OnBnClickedSave()
{
	CContainerInfo *containerInfo = CContainerInfo::GetInstance();

	containerInfo->save();
}
