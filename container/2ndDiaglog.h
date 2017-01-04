// 2ndDiaglog.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"
#include "afxwin.h"


// C2ndDiaglog �Ի���
class C2ndDiaglog : public CDialog
{
// ����
public:
	C2ndDiaglog(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_2ND_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
	// ѡȡ����������
	CListCtrl m_resultEntry;
	// ��ͷ����
	CListCtrl m_connectConfig;
	// �ӹܼ�����
	CListCtrl m_connectResults;
	// �ݻ�
	CEdit m_volumeShow;
	// ѹ��
	CEdit m_pressShow;
	// ��������
	CEdit m_materialShow;
	// ����¶�
	CEdit m_temperatureShow;
	// ����ϵ��
	CEdit m_coefShow;
	// �ӹܸ���
	CEdit m_pipeNumEdit;
	// �Ƿ���Ҫ�˿�
	CButton m_holeCheck;
	// �˿׳ߴ�
	CComboBox m_holeSizeCombo;
	// �˿ײ���
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