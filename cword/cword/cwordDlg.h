// cwordDlg.h : header file
//

#if !defined(AFX_CWORDDLG_H__B02DD0CC_4B05_4C00_8480_14A87D65AE2E__INCLUDED_)
#define AFX_CWORDDLG_H__B02DD0CC_4B05_4C00_8480_14A87D65AE2E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CCwordDlg dialog

class CCwordDlg : public CDialog
{
// Construction
public:
	CCwordDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CCwordDlg)
	enum { IDD = IDD_CWORD_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCwordDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CCwordDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	afx_msg void OnNewWord();
	virtual void OnCancel();
	afx_msg void OnPicture();
	afx_msg void OnBiaoge();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CWORDDLG_H__B02DD0CC_4B05_4C00_8480_14A87D65AE2E__INCLUDED_)
