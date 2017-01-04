// container2.cpp : ����Ӧ�ó��������Ϊ��
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


// Ccontainer2App ����

Ccontainer2App::Ccontainer2App()
{
	// TODO: �ڴ˴���ӹ�����룬
	// ��������Ҫ�ĳ�ʼ�������� InitInstance ��
	m_pCurrentDlg = NULL;
	m_iSteps = IDD_CONTAINER_DIALOG;
}


// Ψһ��һ�� Ccontainer2App ����

Ccontainer2App theApp;


// Ccontainer2App ��ʼ��

BOOL Ccontainer2App::InitInstance()
{

	//���ع����׼��Ϣ
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
