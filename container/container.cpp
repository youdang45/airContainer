// container2.cpp : 定义应用程序的类行为。
//

#include "stdafx.h"
#include "container.h"
#include "containerDlg.h"
#include "2ndDiaglog.h"
#include "CGBStandard.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// Ccontainer2App

BEGIN_MESSAGE_MAP(Ccontainer2App, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// Ccontainer2App 构造

Ccontainer2App::Ccontainer2App()
{
	// TODO: 在此处添加构造代码，
	// 将所有重要的初始化放置在 InitInstance 中
	m_pCurrentDlg = NULL;
	m_iSteps = IDD_CONTAINER_DIALOG;
}


// 唯一的一个 Ccontainer2App 对象

Ccontainer2App theApp;


// Ccontainer2App 初始化

BOOL Ccontainer2App::InitInstance()
{

	//加载国标标准信息
	CGBStandard *gbStd;
	bool loadRet = FALSE;
	gbStd = CGBStandard::GetInstance();
	loadRet = gbStd->LoadGBStandard();
	if (loadRet != TRUE) {
		return FALSE;
	}

	int nResponse;
	while (m_iSteps) //!=0
	{
		nResponse=-1;
		switch (m_iSteps)
		{
			case IDD_CONTAINER_DIALOG://main dialog 
			{
				Ccontainer2Dlg dlg;
				m_pCurrentDlg=&dlg;
				nResponse = dlg.DoModal();					
				switch (nResponse)
				{
					case IDCANCEL:// DialogA
						if (dlg.m_exit)
							m_iSteps = DIALOG_NO;
						else
							m_iSteps = IDD_2ND_DIALOG;
					break;
					case IDOK:
						m_iSteps = DIALOG_NO;
					break;
				}
			}
			break;
			case IDD_2ND_DIALOG:
			{
				C2ndDiaglog sndDlg;
				m_pCurrentDlg=&sndDlg;
				nResponse = sndDlg.DoModal();
				switch (nResponse)
				{
					case IDCANCEL:
						if (sndDlg.m_exit)
							m_iSteps = DIALOG_NO;
						else
							m_iSteps = IDD_CONTAINER_DIALOG;
					break;
					case IDOK:
						m_iSteps=DIALOG_NO;
					break;
				}
			}
// any other dialog
			break;
			default:
				m_iSteps=DIALOG_NO;
		}// end switch (m_iSteps)
		m_pCurrentDlg=0;
	}// end while m_iSteps
	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
