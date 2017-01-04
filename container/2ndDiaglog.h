// 2ndDiaglog.h : 头文件
//

#pragma once
#include "afxcmn.h"
#include "afxwin.h"


// C2ndDiaglog 对话框
class C2ndDiaglog : public CDialog
{
// 构造
public:
	C2ndDiaglog(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_2ND_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedCheck();
	afx_msg void OnEnChangeEditPnum();
	afx_msg void OnClickListEntry(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnClickList2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButton2();

private:
	void setResultEntry();
	void initConResultEntry();
	void initConResultEntryHeader();
	void setConResultEntry();
	void initConnectHeader();
	void initConnectResultHeader();
	void initPipeMaterialCombo();
	void initHoleSizeCombo();
	void initHoleMaterialCombo();
	void setContainerConfig();
	afx_msg void OnKillfocusEdit();
	afx_msg void OnKillfocusCcomboBox();
	afx_msg void OnOK();
	void createEdit(CListCtrl *list, NM_LISTVIEW  *pEditCtrl,
					CEdit *edit, int &Item, int &SubItem, bool &created);
	void destroyEdit(CListCtrl *list, CEdit* edit, int &Item, int &SubItem);
	void createCcombobox(CListCtrl *list, NM_LISTVIEW  *pEditCtrl, CComboBox *combo,
						int &Item, int &SubItem, bool &created);
	void destroyCcombobox(CListCtrl *list, CComboBox* combo, int &Item, int &SubItem);
	void getSavePipeConfig();
	void getSaveHoleConfig();
	void showPipeCalcResult();
	void showHoleCalcResult();
	void showHoleErr();

private:
	// 选取的容器参数
	CListCtrl m_resultEntry;
	// 接头配置
	CListCtrl m_connectConfig;
	// 接管计算结果
	CListCtrl m_connectResults;
	// 容积
	CEdit m_volumeShow;
	// 压力
	CEdit m_pressShow;
	// 容器材料
	CEdit m_materialShow;
	// 设计温度
	CEdit m_temperatureShow;
	// 焊接系数
	CEdit m_coefShow;
	// 接管个数
	CEdit m_pipeNumEdit;
	// 是否需要人孔
	CButton m_holeCheck;
	// 人孔尺寸
	CComboBox m_holeSizeCombo;
	// 人孔材料
	CComboBox m_holeMaterialCombo;
	int e_Item;
	int e_SubItem;
	CEdit m_edit;
	bool editCreated;
	CComboBox m_comBox;
	bool comboCreated;

public:
	bool m_exit;
	afx_msg void OnBnClickedSave();
};