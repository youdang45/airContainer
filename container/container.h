// container2.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������

#define UsualWidth 150
// Ccontainer2App:
// �йش����ʵ�֣������ container2.cpp
//

class Ccontainer2App : public CWinApp
{
private:
	int m_iSteps;
	CDialog* m_pCurrentDlg;
public:
	Ccontainer2App();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern Ccontainer2App theApp;