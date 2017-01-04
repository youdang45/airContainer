// container2Dlg.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"
#include "CContainerModel.h"
#include "CResults.h"
#include "afxwin.h"

// Ccontainer2Dlg �Ի���
class Ccontainer2Dlg : public CDialog
{
// ����
public:
	Ccontainer2Dlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_CONTAINER_DIALOG };

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
	CListCtrl m_resultList;
	bool m_exit;

private:
	void initResultList();
	void initCombo();
	void setDefaultConfig();
	void fillContainerResultList();
	void setConConfigInfo();
	void resetConfig();

public:
	afx_msg void OnBnClickedCancel();
	// �ݻ��༭��
	CEdit m_volumeEdit;
	// ����ѹ��
	CEdit m_pressEdit;
	// ����
	CComboBox m_materialCombo;
	// ����¶�
	CEdit m_temperatureEdit;
	// ����ϵ��
	CComboBox m_coefCombo;
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedReset();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedConCacl();
	afx_msg void OnBnClickedPipeCacl();
	// ��������
	CComboBox m_containerType;
	// ��װ��ʽ
	CComboBox m_installType;
	// �ְ�ں񲽳�
	CComboBox m_thickStep;
	// ��С�߶�
	CEdit m_heightMin;
	// ���߶�
	CEdit m_heightMax;
	// �����ھ�����
	CEdit m_heightDiRateMin;
	// ����������
	CEdit m_outputNum;
	CEdit m_heightDiRateMax;
	// ��ʴԣ��
	CEdit m_cauterization;
	// �ְ��ȸ�ƫ��
	CEdit m_thickNegWindage;
	afx_msg void OnCbnSelchangeComboM();
	afx_msg void OnBnClickedSave();
};
