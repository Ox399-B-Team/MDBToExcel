#pragma once
#include <vector>
#include <algorithm>
// MdbToXlDlg.h: 헤더 파일
//#include "PictureEx.h"
#include "XLAutomation.h"
#include "XLEzAutomation.h"
#include <typeinfo>
#include <afxdb.h>
#include "CPictureEX.h"

#pragma once
using namespace std;
class CMdbToXlDlgAutoProxy;

typedef struct FieldINFO
{
	CString strFieldName;	//원본 이름
	CString strExcelName;	//엑셀 이름(변경됨)
	int nFieldIdcnt;		//필드 인덱스
	int nExcelIdcnt;		//엑셀 인덱스
} FieldINFO;

// CMdbToXlDlg 대화 상자
class CMdbToXlDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CMdbToXlDlg);
	friend class CMdbToXlDlgAutoProxy;

	// 생성입니다.
public:
	CMdbToXlDlg(CWnd* pParent = nullptr);	// 표준 생성자입니다.
	virtual ~CMdbToXlDlg();

	// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_MDBTOXL_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.

// 구현입니다.
protected:
	CMdbToXlDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnCbnSelchangeTable();
	afx_msg void OnLvnItemchangedFieldlist(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnBnClickedbtnfileload();
	afx_msg void OnBnClickedbtninput();
	afx_msg void OnBnClickedbtnsave();
	afx_msg void OnBnClickedbtnupdate();
	afx_msg void OnNMClickExcellist(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnBnClickedbtnallselect();
	afx_msg void OnBnClickedbtncancel();
	afx_msg void OnNMClickFieldlist(NMHDR* pNMHDR, LRESULT* pResult);

	//--------------------------------------
	// 에러 관련
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnOK();
	virtual void OnCancel();
	//-----------------------------------
	// MDB 관련
	CDatabase m_DB;
	CRecordset* m_pRecordset;

	BOOL m_bConn;

	vector<FieldINFO> m_VstTrueFieldValue;	//실제 데이터 원본
	vector<FieldINFO> m_VstExcelValue;		//Excellist 데이터 //실제 excel에 나올 데이터
	FieldINFO m_stFieldInfo;				//이름과 번호 구조체

	CString m_strPathName;
	CString m_strFileName;

	CString m_strTable;
	CString m_strPW;

	void CloseDBConn(BOOL bToDisconn, BOOL bToClose);
	void CallDBTable();

	//-----------------------------------
	// 엑셀 관련
	CFileDialog* m_pExcelDlg;		// excel 다이얼로그 포인터 변수
	CString m_strExcelPathName;
	HANDLE m_hExcelThread;

	CXLEzAutomation *m_pExcelServer;
	
	//HANDLE m_hSaveCancleEvent;		//강사 확인 후 열 것!!
	//-----------------------------------
	// GUI 관련
	int m_nXlRowNum;				// excel 행 증가	
	int m_nidx;						// 마우스 포커스 유지 변수

	CString strClip(CString str);
	CString m_strEdit;
	CComboBox m_ctrlTable;
	CListCtrl m_ctrlFieldList;
	CListCtrl m_ctrlExcelList;
	CProgressCtrl m_ctrlProgress;

	//--------------------------------------
	// 로고 관련
	CPictureEx m_ctrlROGO;

};

