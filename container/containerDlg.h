// container2Dlg.h : 头文件
//

#pragma once
#include "afxcmn.h"
#include "CContainerModel.h"
#include "CResults.h"
#include "afxwin.h"

// Ccontainer2Dlg 对话框
class Ccontainer2Dlg : public CDialog
{
// 构造
public:
	Ccontainer2Dlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_CONTAINER_DIALOG };

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
	// 容积编辑框
	CEdit m_volumeEdit;
	// 计算压力
	CEdit m_pressEdit;
	// 材料
	CComboBox m_materialCombo;
	// 设计温度
	CEdit m_temperatureEdit;
	// 焊接系数
	CComboBox m_coefCombo;
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedReset();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedConCacl();
	afx_msg void OnBnClickedPipeCacl();
	// 容器类型
	CComboBox m_containerType;
	// 安装形式
	CComboBox m_installType;
	// 钢板壁厚步长
	CComboBox m_thickStep;
	// 最小高度
	CEdit m_heightMin;
	// 最大高度
	CEdit m_heightMax;
	// 长度内径比例
	CEdit m_heightDiRateMin;
	// 输出最大数量
	CEdit m_outputNum;
	CEdit m_heightDiRateMax;
	// 腐蚀裕量
	CEdit m_cauterization;
	// 钢板厚度负偏差
	CEdit m_thickNegWindage;
	afx_msg void OnCbnSelchangeComboM();
	afx_msg void OnBnClickedSave();
};
