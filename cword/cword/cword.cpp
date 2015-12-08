// cword.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "cword.h"
#include "cwordDlg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CCwordApp

BEGIN_MESSAGE_MAP(CCwordApp, CWinApp)
	//{{AFX_MSG_MAP(CCwordApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCwordApp construction

CCwordApp::CCwordApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance

}

/////////////////////////////////////////////////////////////////////////////
// The one and only CCwordApp object

CCwordApp theApp;
WordApp::WordApp()
{
}

WordApp::~WordApp()
{
	COleVariant vTrue((short)TRUE),
                vFalse((short)FALSE),
                vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	oSel.ReleaseDispatch();
	oIsh.ReleaseDispatch();
	tables.ReleaseDispatch();
	oDoc.ReleaseDispatch();
	oDocs.ReleaseDispatch();
	oWord.ReleaseDispatch();
	

}


/////////////////////////////////////////////////////////////////////////////
// CCwordApp initialization

BOOL CCwordApp::InitInstance()
{
	AfxEnableControlContainer();
	if(!AfxOleInit())
	{AfxMessageBox("oleInit error");
	return false;
	}

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	//  of your final executable, you should remove from the following
	//  the specific initialization routines you do not need.

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	CCwordDlg dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with OK
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
BOOL WordApp::CreateApp()
{
		if(!oWord.CreateDispatch(TEXT("Word.Application")))
	{
		AfxMessageBox("服务创建失败", MB_OK | MB_SETFOREGROUND);
		return false;
	}
	AfxMessageBox("创建成功");
	return true;

}
BOOL WordApp::CreateDocument()
{
	COleVariant vTrue((short)TRUE),
                vFalse((short)FALSE),
                vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	oDocs = oWord.GetDocuments();

	//准备调用documents::add

	//参数的类型通过观察从word2007录制的宏,在readme.txt里

	oDoc = oDocs.Add(vOpt,vOpt,vOpt,vOpt);
	return true;

}
void WordApp::ShowApp()
{
	oWord.SetVisible(true);//显示word
}

void WordApp::WriteText()
{
	oSel = oWord.GetSelection();
	//调用函数typetext
	oSel.TypeText("hello,test\n\nby wanglz");
	//释放
	oSel.ReleaseDispatch();
}
BOOL WordApp::SaveDocumentAs()
{
	COleVariant vTrue((short)TRUE),
                vFalse((short)FALSE),
                vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

///////////////参考宏saveas()对应_document，_application中getactivedocument()//
	oDoc = oWord.GetActiveDocument();
	oDoc.SaveAs(
		COleVariant("d:\\wanglztest.doc"),
		COleVariant((short) 0),//?
		vFalse,
		COleVariant(""),
		vTrue,
		COleVariant(""),
		vFalse,
		vFalse,
		vFalse,
		vFalse,
		vFalse,
		vOpt, 
		vOpt, 
		vOpt, 
		vOpt,
		vOpt

		);
	oDoc.ReleaseDispatch();
	//oDocs.ReleaseDispatch();
	//oWord.ReleaseDispatch();
	AfxMessageBox("请检查d盘根目录是否产生wanglztest.doc");
	return true;
}
BOOL WordApp::SaveDocument()
{
	oDoc = oWord.GetActiveDocument();
	oDoc.Save();
	return true;

}

BOOL WordApp::Quit()
{
	COleVariant vTrue((short)TRUE),
                vFalse((short)FALSE),
                vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	oWord.Quit(vOpt,vOpt,vOpt);
	return true;
}
BOOL WordApp::OpenDocument()
{
	COleVariant vTrue((short)TRUE),
				vFalse((short)FALSE),
				vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	oDocs = oWord.GetDocuments();
	oDoc =oDocs.Open(	COleVariant("d:\\wanglztest.doc"),
	vFalse,
	vFalse,
	vFalse,
	COleVariant(""),
	COleVariant(""),
	vFalse,
	COleVariant(""),
	COleVariant(""),
	COleVariant((short) 0),
	COleVariant(""),vOpt,vOpt,vOpt,vOpt,vOpt);
	return true;

}
void WordApp::InsertPicture()
{
	COleVariant vTrue((short)TRUE),
				vFalse((short)FALSE),
				vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	oSel = oWord.GetSelection();
	oIsh =oSel.GetInlineShapes();
	oIsh.AddPicture(
		"d:\\1.png",
		vFalse,
		vTrue,
		vOpt
		);
	oIsh.ReleaseDispatch();

}
void WordApp::InserTable()
{
	COleVariant vTrue((short)TRUE),
				vFalse((short)FALSE),
				vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	oSel = oWord.GetSelection();
	oDoc = oWord.GetActiveDocument();
	Tables tables =oDoc.GetTables();
	tables.Add(oSel.GetRange(),2,3,COleVariant((short) 1),COleVariant((short) 0));
	oSel.TypeText("姓名");
	oSel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)0));
	oSel.TypeText("学号");
	oSel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)0));
	oSel.TypeText("专业");
	oSel.MoveDown(COleVariant((short)4),COleVariant((short)1),COleVariant((short)0));
	oSel.TypeText("wanglz");
	//oSel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)0));
	//oSel.MoveDown(COleVariant((short)1),COleVariant((short)1),COleVariant((short)0));
	//oSel.MoveLeft(COleVariant((short)1),COleVariant((short)2),COleVariant((short)0));
	//oSel.TypeText("wanglz");
	//oSel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)0));
	
	oSel.ReleaseDispatch();
	tables.ReleaseDispatch();
}