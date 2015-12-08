// cword.h : main header file for the CWORD application
//

#if !defined(AFX_CWORD_H__8F6CBF43_29A9_45BF_8B2A_67939376C489__INCLUDED_)
#define AFX_CWORD_H__8F6CBF43_29A9_45BF_8B2A_67939376C489__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols
#include "msword.h"

/////////////////////////////////////////////////////////////////////////////
// CCwordApp:
// See cword.cpp for the implementation of this class
//

class CCwordApp : public CWinApp
{
public:
	CCwordApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCwordApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CCwordApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

class WordApp
{
	private:
	_Application oWord;//定义Word提供的应用程序对象;
	_Document oDoc;
	//_Document sDoc;
	Documents oDocs;//定义Word提供的文档对象;
	Selection oSel;//输入点
	InlineShapes oIsh;
	Tables tables;
public:
	WordApp();
	virtual ~WordApp();
public:
	BOOL CreateApp();
	BOOL CreateDocument();
	BOOL OpenDocument();
	BOOL SaveDocument();
	BOOL SaveDocumentAs();
	BOOL Quit();
	void ShowApp();
	void WriteText();
	void InsertPicture();
	void InserTable();
	//void InsertTable
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CWORD_H__8F6CBF43_29A9_45BF_8B2A_67939376C489__INCLUDED_)
