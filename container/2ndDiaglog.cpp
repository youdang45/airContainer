// container2Dlg.cpp : 实现文件
//

#include "stdafx.h"
#include "container.h"
#include "2ndDiaglog.h"
#include "CContainerInfo.h"
#include "CGBStandard.h"
#include "CPipeHole.h"
#include "CPipeHoleModel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// C2ndDiaglog 对话框
void floatToStr(float v, CString & str) {
	if(-0.000001<v && v<0.000001) {
		str = _T("0");
	} else {
		str.Format(_T("%.2f"), v);
	}
}



C2ndDiaglog::C2ndDiaglog(CWnd* pParent /*=NULL*/)
	: CDialog(C2ndDiaglog::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_exit = TRUE;
}

void C2ndDiaglog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_ENTRY, m_resultEntry);
	DDX_Control(pDX, IDC_LIST2, m_connectConfig);
	DDX_Control(pDX, IDC_LIST1, m_connectResults);
	DDX_Control(pDX, IDC_EDIT_VS, m_volumeShow);
	DDX_Control(pDX, IDC_EDIT_PS, m_pressShow);
	DDX_Control(pDX, IDC_EDIT_MS, m_materialShow);
	DDX_Control(pDX, IDC_EDIT_TS, m_temperatureShow);
	DDX_Control(pDX, IDC_EDIT_CS, m_coefShow);
	DDX_Control(pDX, IDC_EDIT_PNUM, m_pipeNumEdit);
	DDX_Control(pDX, IDC_CHECK, m_holeCheck);
	DDX_Control(pDX, IDC_COMBO_SIZE, m_holeSizeCombo);
	DDX_Control(pDX, IDC_COMBO_HOLE_M, m_holeMaterialCombo);
}

BEGIN_MESSAGE_MAP(C2ndDiaglog, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_LAST_PAGE, &C2ndDiaglog::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_CHECK, &C2ndDiaglog::OnBnClickedCheck)
	ON_EN_CHANGE(IDC_EDIT_PNUM, &C2ndDiaglog::OnEnChangeEditPnum)
 	ON_NOTIFY(NM_CLICK, IDC_LIST_ENTRY, &C2ndDiaglog::OnClickListEntry)
	ON_EN_KILLFOCUS(IDC_EDIT_CREATEID, &C2ndDiaglog::OnKillfocusEdit)
	ON_CBN_KILLFOCUS(IDC_COMBOX_CREATEID, &C2ndDiaglog::OnKillfocusCcomboBox)

	ON_NOTIFY(NM_CLICK, IDC_LIST2, &C2ndDiaglog::OnClickList2)
	ON_BN_CLICKED(IDC_CACL, &C2ndDiaglog::OnBnClickedCacl)
	ON_BN_CLICKED(IDSAVE, &C2ndDiaglog::OnBnClickedSave)
END_MESSAGE_MAP()


// Ccontainer2Dlg 消息处理程序

BOOL C2ndDiaglog::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	editCreated = false;
	comboCreated = false;
    initConResultEntry();
	initConnectHeader();
	initConnectResultHeader();
	initHoleSizeCombo();
	initHoleMaterialCombo();
	setContainerConfig();
	m_fileName.Format(_T("%g立方_%g兆帕 %s"), m_volume, m_pressure, PIPE_HOLE_FILE_NAME);


	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void C2ndDiaglog::OnSysCommand(UINT nID, LPARAM lParam)
{
	CDialog::OnSysCommand(nID, lParam);
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void C2ndDiaglog::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR C2ndDiaglog::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//对计算表格进行初始
void C2ndDiaglog::initConResultEntry()
{
	initConResultEntryHeader();
	setConResultEntry();
}

void C2ndDiaglog::initConResultEntryHeader()
{
	LVCOLUMN lvColumn;
	CString strText;

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = 80;
    strText = _T("序号");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_resultEntry.InsertColumn(0, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
	strText = _T("筒体内径（mm）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_resultEntry.InsertColumn(1, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("筒体壁厚（mm）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_resultEntry.InsertColumn(2, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("筒体高度（mm）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_resultEntry.InsertColumn(3, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("封头壁厚（mm）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_resultEntry.InsertColumn(4, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth + 30;
    strText = _T("封头直边高度（mm）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_resultEntry.InsertColumn(5, &lvColumn);
	
	lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("壳体总高（mm）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_resultEntry.InsertColumn(5, &lvColumn);
	
	lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth;
    strText = _T("壳体重量（kg）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_resultEntry.InsertColumn(5, &lvColumn);

    this->m_resultEntry.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

	return;
}

void C2ndDiaglog::setConResultEntry()
{
	CContainerInfo *containerInfo;
	containerInfo = CContainerInfo::GetInstance();


	ConCaclResultItem &conItem = containerInfo->getSelectItem();
	int selectNo = containerInfo->getContainerSelectedNum();

	CString str;
	str.Format(_T("%d"), (selectNo));
	m_resultEntry.InsertItem(0, str);

	str.Format(_T("%.2f"), conItem.diameterIn);
	m_resultEntry.SetItemText(0, 1, str);

	str.Format(_T("%.2f"), conItem.conThick);
	m_resultEntry.SetItemText(0, 2, str);

	str.Format(_T("%.2f"), conItem.conHight);
	m_resultEntry.SetItemText(0, 3, str);

	str.Format(_T("%.2f"), conItem.capThick);
	m_resultEntry.SetItemText(0, 4, str);

	str.Format(_T("%.2f"), conItem.capHight);
	m_resultEntry.SetItemText(0, 7, str);

	str.Format(_T("%.2f"), conItem.totalHight);
	m_resultEntry.SetItemText(0, 6, str);

	str.Format(_T("%.2f"), conItem.weight);
	m_resultEntry.SetItemText(0, 5, str);
}

//对计算表格进行初始
void C2ndDiaglog::initConnectHeader()
{
	LVCOLUMN lvColumn;
	CString strText;

#define ConnUsualWidth (UsualWidth+20)

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = 80;
    strText = _T("序号");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_connectConfig.InsertColumn(0, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = ConnUsualWidth;
	strText = _T("采用补强");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_connectConfig.InsertColumn(1, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = ConnUsualWidth;
    strText = _T("接管材质");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_connectConfig.InsertColumn(2, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = ConnUsualWidth;
    strText = _T("接管外径（mm）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_connectConfig.InsertColumn(3, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = ConnUsualWidth;
    strText = _T("接管投料厚度（mm）");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    m_connectConfig.InsertColumn(4, &lvColumn);

	lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = ConnUsualWidth;
    strText = _T("允许内伸");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
	m_connectConfig.InsertColumn(5, &lvColumn);

    m_connectConfig.SetExtendedStyle(LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);

	return;
}

//对计算表格进行初始
void C2ndDiaglog::initConnectResultHeader()
{
	LVCOLUMN lvColumn;
	CString strText;

#define ConnResultWidth (UsualWidth+60)

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = 50;
    strText = _T("序号");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_connectResults.InsertColumn(0, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = ConnResultWidth-60;
	strText = _T("接管名义厚度");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_connectResults.InsertColumn(1, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = ConnResultWidth-60;
    strText = _T("接管最小内伸高度");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_connectResults.InsertColumn(2, &lvColumn);

	lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
	lvColumn.fmt = LVCFMT_CENTER;
	lvColumn.cx = ConnResultWidth-60;
	strText = _T("接管最小外伸长度");
	lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
	this->m_connectResults.InsertColumn(3, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = ConnResultWidth + 20;
    strText = _T("标准号");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_connectResults.InsertColumn(4, &lvColumn);

    lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth - 30;
    strText = _T("补强最小宽度");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_connectResults.InsertColumn(5, &lvColumn);

	lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
    lvColumn.fmt = LVCFMT_CENTER;
    lvColumn.cx = UsualWidth - 10;
    strText = _T("补强圈最小厚度");
    lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
    this->m_connectResults.InsertColumn(6, &lvColumn);

	lvColumn.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
	lvColumn.fmt = LVCFMT_CENTER;
	lvColumn.cx = ConnResultWidth - 10;
	strText = _T("备注");
	lvColumn.pszText = (LPTSTR)(LPCTSTR)strText;
	this->m_connectResults.InsertColumn(7, &lvColumn);

    this->m_connectResults.SetExtendedStyle(LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);

	return;
}

void C2ndDiaglog::initHoleSizeCombo()
{
	m_holeSizeCombo.EnableWindow(FALSE);
	m_holeSizeCombo.AddString(_T("280*380"));
	m_holeSizeCombo.AddString(_T("300*400"));
}

void C2ndDiaglog::initHoleMaterialCombo()
{
	m_holeMaterialCombo.EnableWindow(FALSE);
	m_holeMaterialCombo.AddString(_T("Q245R"));
	m_holeMaterialCombo.AddString(_T("Q345R"));
	m_holeMaterialCombo.AddString(_T("S30408"));
}

void C2ndDiaglog::OnBnClickedCancel()
{
	OnCancel();
	m_exit = FALSE;
}

void C2ndDiaglog::OnBnClickedCheck()
{
	if (m_holeCheck.GetCheck()) {
		GetDlgItem(IDC_COMBO_SIZE)->EnableWindow(TRUE);
		GetDlgItem(IDC_COMBO_HOLE_M)->EnableWindow(TRUE);
		GetDlgItem(IDC_STATIC_HT)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT_HT)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_HINL)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT_HINL)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_HOUTL)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT_HOUTL)->ShowWindow(SW_SHOW);
	} else {
		GetDlgItem(IDC_COMBO_SIZE)->EnableWindow(FALSE);
		GetDlgItem(IDC_COMBO_HOLE_M)->EnableWindow(FALSE);
		GetDlgItem(IDC_STATIC_HT)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT_HT)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_HINL)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT_HINL)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_HOUTL)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT_HOUTL)->ShowWindow(SW_HIDE);
	}
}

void C2ndDiaglog::setContainerConfig()
{
	bool ret = FALSE;
	CContainerInfo *containerInfo;
	containerInfo = CContainerInfo::GetInstance();
	conConfig_t config;
	CString str;

	ret = containerInfo->getContainerConfig(config);
	if (ret == FALSE) {
		//Add Error Msg;
		exit(1);
	}

	str.Format(_T("%.2f"), config.volume);
	GetDlgItem(IDC_EDIT_VS)->SetWindowText(str); 

	str.Format(_T("%.2f"), config.pressure);
	GetDlgItem(IDC_EDIT_PS)->SetWindowText(str); 

	str = getSteelNameByNum(config.conMetarial);
	GetDlgItem(IDC_EDIT_MS)->SetWindowText(str);
	
	str.Format(_T("%u"), config.temperature);
	GetDlgItem(IDC_EDIT_TS)->SetWindowText(str); 

	str.Format(_T("%.2f"), config.coefficient);
	GetDlgItem(IDC_EDIT_CS)->SetWindowText(str); 

	m_volume = config.volume;
	m_pressure = config.pressure;
}

void C2ndDiaglog::OnEnChangeEditPnum()
{
	// TODO:  如果该控件是 RICHEDIT 控件，则它将不会
	// 发送该通知，除非重写 CDialog::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	CString str;
	GetDlgItem(IDC_EDIT_PNUM)->GetWindowText(str);
	m_connectConfig.DeleteAllItems();
	int num = _ttoi(str);
	for (int i = 0; i < num; i++) {
		str.Format(_T("%d"), (i+1));
		m_connectConfig.InsertItem(i, str);

		str.Format(_T("否"));
		m_connectConfig.SetItemText(i, 1, str);

		GetDlgItemText(IDC_EDIT_MS, str);
		if(str == steelNumStrMap[S30408].steelNumName) {
			str.Format(steelNumStrMap[S30408].steelNumName);
			m_connectConfig.SetItemText(i, 2, str);
		} else {
			str.Format(steelNumStrMap[E_20].steelNumName);
			m_connectConfig.SetItemText(i, 2, str);
		}
		str.Format(_T("\\"));
		m_connectConfig.SetItemText(i, 4, str);

		str.Format(_T("是"));
		m_connectConfig.SetItemText(i, 5, str);		
	}
}

void C2ndDiaglog::OnClickListEntry(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	NM_LISTVIEW  *pEditCtrl = (NM_LISTVIEW *)pNMHDR;
	CContainerInfo *containerInfo = CContainerInfo::GetInstance();

	if (pEditCtrl->iItem==-1)//点击到非工作区
	{
		if (editCreated == true)//如果之前创建了编辑框就销毁掉
		{
			destroyEdit(&m_resultEntry, &m_edit, e_Item, e_SubItem);//销毁单元格编辑框对象
			editCreated = false;
		}
		if (comboCreated == true)//如果之前创建了下拉列表框就销毁掉
		{
			destroyCcombobox(&m_resultEntry, &m_comBox, e_Item, e_SubItem);
			comboCreated = false;
		}
	}
	else if (pEditCtrl->iSubItem != 0)//如果是第一列
	{
		//Do nothing
	}
	else//如果不是第一列
	{
		if (editCreated == true)
		{
			destroyEdit(&m_resultEntry, &m_edit, e_Item, e_SubItem);
			editCreated = false;
		}
		if (comboCreated == true)
		{
			if (!(e_Item == pEditCtrl->iItem && e_SubItem == pEditCtrl->iSubItem))
			{
				destroyCcombobox(&m_resultEntry, &m_comBox, e_Item, e_SubItem);
				comboCreated = false;
				createCcombobox(&m_resultEntry, pEditCtrl, &m_comBox, e_Item, e_SubItem, comboCreated);

				list <ConCaclResultItem> & resultList = containerInfo->getContainerResultList();
				int max = (resultList.size() > 20) ? 20 : resultList.size();
				for (int i = 0; i < max; i++) {
					CString str;
					str.Format(_T("%d"), (i+1));
					m_comBox.AddString(str);
				}
				m_comBox.ShowDropDown();
			}
			else
			{
				m_comBox.SetFocus();
			}
		}
		else
		{
			e_Item = pEditCtrl->iItem;
			e_SubItem = pEditCtrl->iSubItem;
			createCcombobox(&m_resultEntry, pEditCtrl, &m_comBox, e_Item, e_SubItem, comboCreated);

			list <ConCaclResultItem> & resultList = containerInfo->getContainerResultList();
			unsigned int outputNum = containerInfo->getOutputNum();
			int max = (resultList.size() > outputNum) ? outputNum : resultList.size();
			for (int i = 0; i < max; i++) {
				CString str;
				str.Format(_T("%d"), (i+1));
				m_comBox.AddString(str);
			}
			m_comBox.ShowDropDown();//自动下拉
		}
	}
	*pResult = 0;
}

void C2ndDiaglog::createEdit(CListCtrl *list, NM_LISTVIEW  *pEditCtrl, CEdit *edit,
							 int &Item, int &SubItem, bool &created)
{
	Item = pEditCtrl->iItem;
	SubItem = pEditCtrl->iSubItem;
	edit->Create(ES_AUTOHSCROLL | WS_CHILD | ES_LEFT | ES_WANTRETURN,
		CRect(0, 0, 0, 0), this, IDC_EDIT_CREATEID);
	created = true;
	edit->SetFont(this->GetFont(), FALSE);
	edit->SetParent(list);
	CRect  EditRect;
	list->GetSubItemRect(e_Item, e_SubItem, LVIR_LABEL, EditRect);
	EditRect.SetRect(EditRect.left+1, EditRect.top+1, 
		EditRect.left + list->GetColumnWidth(e_SubItem)-1, EditRect.bottom-1);
	CString strItem = list->GetItemText(e_Item, e_SubItem);
	edit->SetWindowText(strItem);
	edit->MoveWindow(&EditRect);
	edit->ShowWindow(SW_SHOW);
	edit->SetFocus();
	edit->SetSel(-1);
}

void C2ndDiaglog::destroyEdit(CListCtrl *list,CEdit *edit, int &Item, int &SubItem)
{
	CString str;

	edit->GetWindowText(str);
	list->SetItemText(Item, SubItem, str);
	edit->DestroyWindow();
	editCreated = FALSE;
}

void C2ndDiaglog::createCcombobox(CListCtrl *list, NM_LISTVIEW  *pEditCtrl, CComboBox *combo,
								  int &Item, int &SubItem, bool &created)
{
	Item = pEditCtrl->iItem;
	SubItem = pEditCtrl->iSubItem;
	created = true;
	combo->Create(WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWN | CBS_OEMCONVERT | CBS_AUTOHSCROLL,
				CRect(0, 0, 0, 0), this, IDC_COMBOX_CREATEID);
	combo->SetFont(this->GetFont(), FALSE);
	combo->SetParent(list);
	CRect  Rect;
	list->GetSubItemRect(e_Item, e_SubItem, LVIR_LABEL, Rect);

	Rect.SetRect(Rect.left + 1, Rect.top + 1, 
				Rect.left + list->GetColumnWidth(e_SubItem) - 1, 
				(Rect.top +5*Rect.Height()));

	CString strItem = list->GetItemText(e_Item, e_SubItem);
	combo->SetWindowText(strItem);
	combo->MoveWindow(&Rect);
	combo->ShowWindow(SW_SHOW);
}

void C2ndDiaglog::destroyCcombobox(CListCtrl *list, CComboBox* combo, int &Item, int &SubItem)
{
	CString str;

	combo->GetWindowText(str);   
	combo->ShowWindow(SW_HIDE);
	if (list->GetDlgCtrlID() == IDC_LIST_ENTRY) {
		CContainerInfo *containerInfo;
		containerInfo = CContainerInfo::GetInstance();

		m_resultEntry.DeleteAllItems();
		containerInfo->setContainerSelectedNum(_ttoi(str));
		setConResultEntry();
	} else {
		list->SetItemText(Item, SubItem, str);
		if (SubItem == 1){
			CString strTmp;
			if (str == _T("否")){
				strTmp = _T("\\");
				list->SetItemText(Item, 4, strTmp);
			} else {
				strTmp = _T("");
				list->SetItemText(Item, 4, strTmp);
			}
		}
	}
	combo->DestroyWindow();//销毁对象，有创建就要有销毁，不然会报错
	comboCreated = FALSE;
}

void C2ndDiaglog::OnKillfocusCcomboBox()
{


	if (editCreated == true)//如果之前创建了编辑框就销毁掉
	{
		CListCtrl * list = (CListCtrl *)m_edit.GetParent();
		destroyEdit(list, &m_edit, e_Item, e_SubItem);//销毁单元格编辑框对象
		editCreated = false;
	}
	if (comboCreated == true)//如果之前创建了下拉列表框就销毁掉
	{
		CListCtrl * list = (CListCtrl *)m_comBox.GetParent();
		destroyCcombobox(list, &m_comBox, e_Item, e_SubItem);
		comboCreated = false;
	}


}

void C2ndDiaglog::OnKillfocusEdit()
{
	if (editCreated == true)//如果之前创建了编辑框就销毁掉
	{
		CListCtrl * list = (CListCtrl *)m_edit.GetParent();
		destroyEdit(list, &m_edit, e_Item, e_SubItem);//销毁单元格编辑框对象
		editCreated = false;
	}
	if (comboCreated == true)//如果之前创建了下拉列表框就销毁掉
	{
		CListCtrl * list = (CListCtrl *)m_comBox.GetParent();
		destroyCcombobox(list, &m_comBox, e_Item, e_SubItem);
		comboCreated = false;
	}
}

void C2ndDiaglog::OnOK()
{
}

void C2ndDiaglog::OnClickList2(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	NM_LISTVIEW  *pEditCtrl = (NM_LISTVIEW *)pNMHDR;
	//CListCtrl *list = (CListCtrl *)GetDlgItem(pNMHDR->idFrom);
	CListCtrl *list = NULL;

	if ((pEditCtrl->iItem == -1) || (pEditCtrl->iSubItem == 0))//点击到非工作区
	{
		if (editCreated == true)//如果之前创建了编辑框就销毁掉
		{
			list = (CListCtrl *)m_edit.GetParent();
			destroyEdit(list, &m_edit, e_Item, e_SubItem);//销毁单元格编辑框对象
			editCreated = false;
		}
		if (comboCreated == true)//如果之前创建了下拉列表框就销毁掉
		{
			list = (CListCtrl *)m_comBox.GetParent();
			destroyCcombobox(list, &m_comBox, e_Item, e_SubItem);
			comboCreated = false;
		}
	}
	else if (pEditCtrl->iSubItem == 3 || pEditCtrl->iSubItem == 4)
	{
		if (comboCreated == true)//如果之前创建了编辑框就销毁掉
		{
			list = (CListCtrl *)m_comBox.GetParent();
			destroyCcombobox(list, &m_comBox, e_Item, e_SubItem);
			comboCreated = false;
		}
		if (editCreated == true)
		{
			if (!(e_Item == pEditCtrl->iItem && e_SubItem == pEditCtrl->iSubItem))//如果点中的单元格不是之前创建好的
			{
				list = (CListCtrl *)m_edit.GetParent();
				destroyEdit(list, &m_edit, e_Item, e_SubItem);
				editCreated = false;
				if (pEditCtrl->iSubItem == 4) {
					CString str = list->GetItemText(pEditCtrl->iItem, 1);
					if (str == _T("否")) {
						//Do nothing
					} else {
						createEdit(&m_connectConfig, pEditCtrl, &m_edit, e_Item, e_SubItem, editCreated);//创建编辑框
					}
				} else {
					createEdit(&m_connectConfig, pEditCtrl, &m_edit, e_Item, e_SubItem, editCreated);//创建编辑框
				}

			}
			else//点中的单元格是之前创建好的
			{
				m_edit.SetFocus();//设置为焦点 
			}
		}
		else
		{
			e_Item = pEditCtrl->iItem;//将点中的单元格的行赋值给“刚编辑过的行”以便后期处理
			e_SubItem = pEditCtrl->iSubItem;//将点中的单元格的行赋值给“刚编辑过的行”以便后期处理
			if (pEditCtrl->iSubItem == 4) {				
				CString str = m_connectConfig.GetItemText(pEditCtrl->iItem, 1);
				if (str == _T("否")) {
					//Do nothing
				} else {
					createEdit(&m_connectConfig, pEditCtrl, &m_edit, e_Item, e_SubItem, editCreated);//创建编辑框
				}
			} else {
				createEdit(&m_connectConfig, pEditCtrl, &m_edit, e_Item, e_SubItem, editCreated);//创建编辑框
			}
		}
	}
	else //其他三列
	{

		if ((editCreated == true) || (comboCreated == true))
		{
			if (!(e_Item == pEditCtrl->iItem && e_SubItem == pEditCtrl->iSubItem))
			{
				list = (CListCtrl *)m_comBox.GetParent();
				destroyCcombobox(list, &m_comBox, e_Item, e_SubItem);
				comboCreated = false;
				createCcombobox(&m_connectConfig, pEditCtrl, &m_comBox, e_Item, e_SubItem, comboCreated);

				if(pEditCtrl->iSubItem == 1 || pEditCtrl->iSubItem == 5) {
					m_comBox.AddString(_T("是"));
					m_comBox.AddString(_T("否"));
				} else {
					m_comBox.AddString(_T("10"));
					m_comBox.AddString(_T("20"));
					m_comBox.AddString(_T("S30408"));
				}
				m_comBox.ShowDropDown();
			}
			else
			{
				m_comBox.SetFocus();
			}
		}
		else
		{
			e_Item = pEditCtrl->iItem;
			e_SubItem = pEditCtrl->iSubItem;
			createCcombobox(&m_connectConfig, pEditCtrl, &m_comBox, e_Item, e_SubItem, comboCreated);

			if(pEditCtrl->iSubItem == 1 || pEditCtrl->iSubItem == 5) {
				m_comBox.AddString(_T("是"));
				m_comBox.AddString(_T("否"));
			} else {
				m_comBox.AddString(_T("10"));
				m_comBox.AddString(_T("20"));
				m_comBox.AddString(_T("S30408"));
			}
			m_comBox.ShowDropDown();//自动下拉
		}
	}
	*pResult = 0;
}

void C2ndDiaglog::OnBnClickedCacl()
{
	bool hole_ret = FALSE;
	CPipeHole *pPipeHole = CPipeHole::GetInstance();
	CPipeModel *pPipeModel = CPipeModel::GetInstance();
	CHoleModel *pHoleModel = CHoleModel::GetInstance();

	//清空原有信息
	pPipeHole->reset();
	m_connectResults.DeleteAllItems();

	getSavePipeConfig();
	pPipeModel->calcPipes();	
	showPipeCalcResult();

	if (m_holeCheck.GetCheck()) {
		getSaveHoleConfig();

		hole_ret = pHoleModel->calcHole();
		if(hole_ret == TRUE) {
			showHoleCalcResult();
		} else {
			//showHoleErr(); TBD
		}
	}
}

void C2ndDiaglog::getSavePipeConfig(){
	CPipeHole *pPipeHole = CPipeHole::GetInstance();
	pipeConfig_t pipeConfig;

	for (int i = 0; i < m_connectConfig.GetItemCount(); i++) {
		CString str;
		str = m_connectConfig.GetItemText(i, 1) ;
		pipeConfig.isAddStress = ( str == _T("是"));

		str = m_connectConfig.GetItemText(i, 2); 
		pipeConfig.material = getSteelNumberByName(str);

		str = m_connectConfig.GetItemText(i, 3);
		pipeConfig.Do =  _wtof(str);

		if (pipeConfig.isAddStress) {
			str = m_connectConfig.GetItemText(i, 4);
			pipeConfig.thick =  _wtof(str);
		} else {
			pipeConfig.thick = 0;
		}

		str = m_connectConfig.GetItemText(i, 5);
		pipeConfig.extendIn = (str == _T("是"));

		pPipeHole->addPipeConfig(pipeConfig);
	}
}

void C2ndDiaglog::getSaveHoleConfig()
{
	holeConfig_t holeConf;
	CString str;
	CPipeHole *pPipeHole = CPipeHole::GetInstance();

	m_holeMaterialCombo.GetWindowText(str);
	holeConf.material = getSteelNumberByName(str);

	m_holeSizeCombo.GetWindowText(str);
	holeConf.size = getHoleSizeByStr(str);

	pPipeHole->setHoleConfig(holeConf);
}

void C2ndDiaglog::showPipeCalcResult()
{
	CString str;
	CPipeHole *pPipeHole = CPipeHole::GetInstance();

	list<pipeCalcResult_t> &pipeResults = pPipeHole->getPipeCalcResults();
	int i = 0;
	for (list<pipeCalcResult_t>::iterator iter = pipeResults.begin();
		iter != pipeResults.end(); iter++, i++) {
		str.Format(_T("%d"), (i+1));
		m_connectResults.InsertItem(i, str);

		floatToStr(iter->minDesignThick, str);
		m_connectResults.SetItemText(i, 1, str);

		floatToStr(iter->minExtInHeight, str);
		m_connectResults.SetItemText(i, 2, str);

		floatToStr(iter->minExtOutHeight, str);
		m_connectResults.SetItemText(i, 3, str);

		m_connectResults.SetItemText(i, 4, iter->GBStr);

		floatToStr(iter->minWidth, str);
		m_connectResults.SetItemText(i, 5, str);

		floatToStr(iter->addPressThick, str);
		m_connectResults.SetItemText(i, 6, str);

		switch(iter->recommendAddPress) {
			case COMMENT_ADD_PRESS:
				str = _T("建议采用补强圈进行补强");
				break;
			case COMMENT_NOT_ADD_PRESS:
				str = _T("不需要进行补强");
				break;
			default:
				str = "";
				break;
		}
		m_connectResults.SetItemText(i, 7, str);

		
	}
}

void C2ndDiaglog::showHoleCalcResult()
{
	CString str;
	CPipeHole *pPipeHole = CPipeHole::GetInstance();

	holeCalcResult_t &holeResult = pPipeHole->getHoleCalcResult();
	bool calcRet = FALSE;

	str.Format(_T("%.2f"), holeResult.thickness);
	GetDlgItem(IDC_EDIT_HT)->SetWindowText(str);

	str.Format(_T("%.2f"), holeResult.extInLenght);
	GetDlgItem(IDC_EDIT_HINL)->SetWindowText(str);

	str.Format(_T("%.2f"), holeResult.extOutLenght);
	GetDlgItem(IDC_EDIT_HOUTL)->SetWindowText(str);
}

void C2ndDiaglog::OnBnClickedSave()
{
	CString fileName;
	CFileDialog fileDlg(FALSE, (LPCTSTR)_T("xls"), (LPCTSTR)m_fileName, 
					OFN_OVERWRITEPROMPT, (LPCTSTR)_T("Excel Workbook(*.xls)"), this);
	if ( fileDlg.DoModal() == IDOK ) {
		fileName = fileDlg.GetPathName();
		CPipeHole *phInfo = CPipeHole::GetInstance();
		phInfo->save(fileName);
	} 
}