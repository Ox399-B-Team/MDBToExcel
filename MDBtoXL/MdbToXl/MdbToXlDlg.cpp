// MdbToXlDlg.cpp: 구현 파일
//
#include "pch.h"
#include "framework.h"
#include "MdbToXl.h"
#include "MdbToXlDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"
//---데이터베이스 관련 헤더파일 추가
#include <afxdb.h>

//엑셀 관련 헤더파일 추가
#include <afxdisp.h>
#undef _WINDOWS_
#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// 응용 프로그램 정보에 사용되는 CAboutDlg 대화 상자입니다.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 지원입니다.

// 구현입니다.
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()

// CMdbToXlDlg 대화 상자

IMPLEMENT_DYNAMIC(CMdbToXlDlg, CDialogEx);

CMdbToXlDlg::CMdbToXlDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_MDBTOXL_DIALOG, pParent)
	, m_strPW(_T(""))
	, m_strTable(_T(""))
	, m_strEdit(_T(""))
	, m_strFileName(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = nullptr;
	m_pRecordset = NULL;

	m_hExcelThread = NULL; // 스레드 : 핸들 초기화
	m_strExcelPathName = _T(""); //스레드 : 저장 파일명 초기화 
	m_nTotalRecordCount = 0;
}

void CMdbToXlDlg::CloseDBConn(BOOL bToDisconn = FALSE, BOOL bToClose = FALSE)
{
	if (m_bConn)
	{
		m_nXlRowNum = 0;
		m_stFieldInfo.nExcelIdcnt = 0;
		m_stFieldInfo.nExcelIdcnt = 0;
		m_stFieldInfo.strExcelName = _T("");
		m_stFieldInfo.strFieldName = _T("");

		m_pRecordset->Close();
		if (bToDisconn)
		{
			delete m_pRecordset;
			m_DB.Close();
		}
		if (!bToClose)
		{
			m_ctrlExcelList.DeleteAllItems();
			m_ctrlFieldList.DeleteAllItems();
			SetDlgItemText(IDC_UPDATE_NAME, _T(""));

			m_VstExcelValue.clear();
			m_VstTrueFieldValue.clear();
		}
		m_bConn = FALSE;
	}
}

void CMdbToXlDlg::CallDBTable()
{
	HSTMT hStmt;
	SQLLEN lLen;
	CString strUnicode;

	char pcName[256];
	int nIdx = 0;

	int a = 0;
	CString outputName = _T("");
	BOOL bswitch = FALSE;

	SQLAllocStmt(m_DB.m_hdbc, &hStmt);
	if (SQLTables(hStmt, NULL, 0, NULL, 0, NULL, 0, _T("TABLE"), SQL_NTS) != SQL_ERROR)
	{ /* OK */
		if (SQLFetch(hStmt) != SQL_NO_DATA_FOUND)
		{ /* Data found */
			while (!SQLGetData(hStmt, 3, SQL_C_CHAR, pcName, 256, &lLen))
			{ /* We have a name */
				if (pcName[0] != _T('~'))
				{
					// Do something with the name here
					strUnicode = pcName; // Char to CString

					if (strUnicode == _T("SV_DVVAL"))
					{
						outputName = strUnicode;
						bswitch = TRUE;
						a = nIdx;
					}
					m_ctrlTable.InsertString(nIdx, strUnicode);
					nIdx++;
				}
				SQLFetch(hStmt);
			}
		}
	}
	SQLFreeStmt(hStmt, SQL_CLOSE);
	if (bswitch)
		m_ctrlTable.SetCurSel(a);
	else
		m_ctrlTable.SetCurSel(0);

}

void CMdbToXlDlg::FirstInput(CString strTable)
{
	if (m_strTable == strTable)
	{
		m_ctrlFieldList.SetCheck(2);
		m_ctrlFieldList.SetCheck(3);
		m_ctrlFieldList.SetCheck(7);
		m_ctrlFieldList.SetCheck(4);
		m_ctrlFieldList.SetCheck(5);

		m_ctrlExcelList.SetItemText(0, 0, _T("SVID"));
		m_VstExcelValue.at(0).strExcelName = _T("SVID");

		m_ctrlExcelList.SetItemText(1, 0, _T("VID NAME"));
		m_VstExcelValue.at(1).strExcelName = _T("VID NAME");

		m_ctrlExcelList.SetItemText(2, 0, _T("VID Description"));
		m_VstExcelValue.at(2).strExcelName = _T("VID Description");

		m_ctrlExcelList.SetItemText(3, 0, _T("FORMAT"));
		m_VstExcelValue.at(3).strExcelName = _T("FORMAT");

		m_ctrlExcelList.SetItemText(4, 0, _T("Type Size"));
		m_VstExcelValue.at(4).strExcelName = _T("Type Size");
	}	
}

CMdbToXlDlg::~CMdbToXlDlg()
{
	// 이 대화 상자에 대한 자동화 프록시가 있을 경우 이 대화 상자에 대한
	//  후방 포인터를 null로 설정하여
	//  대화 상자가 삭제되었음을 알 수 있게 합니다.
	if (m_pAutoProxy != nullptr)
		m_pAutoProxy->m_pDialog = nullptr;

	//스레드 동작 중일 시 강제 종료
	if (m_hExcelThread != NULL)
	{
		//m_pExcelServer->ReleaseExcel();
		//delete m_pExcelServer;		
	
		KillProcess(_T("EXCEL.EXE")); //Excel 프로세스 종료 함수
		TerminateThread(m_hExcelThread, 0);
		CloseHandle(m_hExcelThread);	
		m_hExcelThread = NULL;
	}

	CloseDBConn(m_bConn, TRUE);
}

void CMdbToXlDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Control(pDX, IDC_TABLE, m_ctrlTable);
	DDX_Control(pDX, IDC_FieldList, m_ctrlFieldList);
	DDX_Control(pDX, IDC_ExcelList, m_ctrlExcelList);
	DDX_Control(pDX, IDC_DataSaveProgress, m_ctrlProgress);
	DDX_Text(pDX, IDC_FILENAME, m_strFileName);
	DDX_Text(pDX, IDC_PASSWORD, m_strPW);
	DDX_Text(pDX, IDC_UPDATE_NAME, m_strEdit);
	DDX_CBString(pDX, IDC_TABLE, m_strTable);
	DDX_Control(pDX, IDC_LOGO, m_ctrlROGO);
	DDX_Control(pDX, IDC_STATIC_TEST, TEST);
}

BEGIN_MESSAGE_MAP(CMdbToXlDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(btnFileLoad, &CMdbToXlDlg::OnBnClickedbtnfileload)
	ON_BN_CLICKED(btnInput, &CMdbToXlDlg::OnBnClickedbtninput)
	ON_CBN_SELCHANGE(IDC_TABLE, &CMdbToXlDlg::OnCbnSelchangeTable)
	ON_BN_CLICKED(btnSave, &CMdbToXlDlg::OnBnClickedbtnsave)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_FieldList, &CMdbToXlDlg::OnLvnItemchangedFieldlist)
	ON_BN_CLICKED(btnUpdate, &CMdbToXlDlg::OnBnClickedbtnupdate)
	ON_NOTIFY(NM_CLICK, IDC_ExcelList, &CMdbToXlDlg::OnNMClickExcellist)
	ON_BN_CLICKED(btnAllSelect, &CMdbToXlDlg::OnBnClickedbtnallselect)
	ON_BN_CLICKED(btnCancel, &CMdbToXlDlg::OnBnClickedbtncancel)
	ON_NOTIFY(NM_CLICK, IDC_FieldList, &CMdbToXlDlg::OnNMClickFieldlist)
	ON_BN_CLICKED(IDC_DATASAVE_XJEM, &CMdbToXlDlg::OnBnClickedDatasaveXjem)
END_MESSAGE_MAP()

// CMdbToXlDlg 메시지 처리기

BOOL CMdbToXlDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 시스템 메뉴에 "정보..." 메뉴 항목을 추가합니다.
	// IDM_ABOUTBOX는 시스템 명령 범위에 있어야 합니다.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	// TODO: 여기에 추가 초기화 작업을 추가합니다.
	GetDlgItem(btnInput)->EnableWindow(FALSE);
	GetDlgItem(btnAllSelect)->EnableWindow(FALSE);
	GetDlgItem(btnCancel)->EnableWindow(FALSE);
	GetDlgItem(btnUpdate)->EnableWindow(FALSE);
	GetDlgItem(btnSave)->EnableWindow(FALSE);
	GetDlgItem(IDC_UPDATE_NAME)->EnableWindow(FALSE);
	GetDlgItem(IDC_FILENAME)->EnableWindow(FALSE);
	GetDlgItem(IDC_PASSWORD)->EnableWindow(FALSE);

	m_ctrlFieldList.SetExtendedStyle(m_ctrlFieldList.GetExtendedStyle() |
		LVS_EX_GRIDLINES | LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT);
	m_ctrlFieldList.ModifyStyle(0, LVS_SHOWSELALWAYS);

	m_ctrlFieldList.InsertColumn(0, _T("Field Name"), LVCFMT_RIGHT, 205);

	m_ctrlExcelList.SetExtendedStyle(m_ctrlExcelList.GetExtendedStyle() |
		LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);
	m_ctrlExcelList.ModifyStyle(0, LVS_SHOWSELALWAYS);
	m_ctrlExcelList.InsertColumn(0, _T("Field Name"), LVCFMT_RIGHT, 215);

	m_bConn = FALSE;
	m_nTotalRecordCount = 0;

	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

void CMdbToXlDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else if ((nID & 0xFFF0) == SC_CLOSE) // ESC버튼은 닫힐 수 있도록 하기위한 조건
	{
		EndDialog(IDCANCEL);
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 애플리케이션의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CMdbToXlDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CMdbToXlDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 컨트롤러에서 해당 개체 중 하나를 계속 사용하고 있을 경우
//  사용자가 UI를 닫을 때 자동화 서버를 종료하면 안 됩니다.  이들
//  메시지 처리기는 프록시가 아직 사용 중인 경우 UI는 숨기지만,
//  UI가 표시되지 않아도 대화 상자는
//  남겨 둡니다.

BOOL CMdbToXlDlg::CanExit()
{
	// 프록시 개체가 계속 남아 있으면 자동화 컨트롤러에서는
	//  이 애플리케이션을 계속 사용합니다.  대화 상자는 남겨 두지만
	//  해당 UI는 숨깁니다.
	if (m_pAutoProxy != nullptr)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}
	return TRUE;
}



// 확장자를 뺀 파일명 복사
CString CMdbToXlDlg::strClip(CString str)
{
	//파일 Full Path를 복사
	TCHAR szTmp[4096];
	StrCpy(szTmp, str);
	CString strTmp;

	CString strResult = _T("");

	// 확장자를 뺀 파일명 복사
	strTmp = PathFindFileName(szTmp);
	ZeroMemory(szTmp, 4096);
	StrCpy(szTmp, strTmp);
	PathRemoveExtension(szTmp);
	strResult = szTmp;

	return strResult;
}

void CMdbToXlDlg::OnBnClickedbtnfileload()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	SetDlgItemText(IDC_SAVE_PROGRESS, _T(""));
	CFileDialog fileDialog(TRUE, _T("mdb"), NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT
		| OFN_FILEMUSTEXIST, _T("mdb File(*.mdb) |*.mdb|모든파일(*.*)|*.*|"));

	int iReturn = fileDialog.DoModal();
	if (iReturn == IDOK)
	{
		m_strPathName = fileDialog.GetPathName();
		m_strFileName = strClip(fileDialog.GetFileName());

		UpdateData(0);//컨트롤 <-- 변수 // 파일이름 띄우기

		SetDlgItemText(IDC_PASSWORD, _T(""));

		////////////////////////////////////////////////////////
		// 비밀번호가 없을 시 연결 

		m_ctrlTable.ResetContent();

		CloseDBConn(m_bConn);

		m_nXlRowNum = 0;

		m_stFieldInfo.nFieldIdcnt = 0;
		m_stFieldInfo.nExcelIdcnt = 0;
		m_stFieldInfo.strExcelName = _T("");
		m_stFieldInfo.strFieldName = _T("");

		TRY
		{
			//DB연결
			// //B1594C47 비밀번호
			CString strtemp;
			strtemp.Format(_T("Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=%s;"),
						m_strPathName);
			m_DB.OpenEx(strtemp, CDatabase::noOdbcDialog);

			// 데이터 테이블 목록 불러오기
			CallDBTable();

			// TABLE에 포함된 FIELD값들을 가져와 표시하는 부분
			CODBCFieldInfo fieldInfo;
			CString strQuery;

			UpdateData(1);// 컨트롤 >> 변수 // 테이블 변경 시 필드 목록 재설정

			TRY
			{
				CRecordset a;
				m_pRecordset = new CRecordset(&m_DB);
							
				// 총 레코드 값 가져오기----------------------------------------
				m_nTotalRecordCount = CalcTotalRow(m_strTable, &m_DB);
				//------------------------------------------------------------

				strQuery.Format(_T("SELECT * FROM %s"), m_strTable);
				m_pRecordset->Open(CRecordset::dynaset, strQuery);

				int FCount = m_pRecordset->GetODBCFieldCount();

				for (int i = 0; i < FCount; i++)
				{
					m_pRecordset->GetODBCFieldInfo(i, fieldInfo);
					m_stFieldInfo.strFieldName = fieldInfo.m_strName;
					m_stFieldInfo.nFieldIdcnt = i;

					m_ctrlFieldList.InsertItem(i, m_stFieldInfo.strFieldName);
					m_VstTrueFieldValue.push_back(m_stFieldInfo);
				}
				GetDlgItem(btnCancel)->EnableWindow(FALSE);
				GetDlgItem(btnAllSelect)->EnableWindow(TRUE);
				GetDlgItem(IDC_PASSWORD)->EnableWindow(FALSE);
				GetDlgItem(btnInput)->EnableWindow(FALSE);
				
			}
			CATCH(CDBException, e)
			{
				GetDlgItem(btnAllSelect)->EnableWindow(FALSE);
				GetDlgItem(btnCancel)->EnableWindow(FALSE);
				GetDlgItem(btnInput)->EnableWindow(FALSE);
				GetDlgItem(btnUpdate)->EnableWindow(FALSE);

				AfxMessageBox(_T("해당 테이블이 존재하지 않습니다."));
			}
			END_CATCH;

			FirstInput(_T("SV_DVVAL"));
			m_bConn = TRUE;
		}
		CATCH(CDBException, e)
		{
			GetDlgItem(IDC_PASSWORD)->EnableWindow(TRUE);
			GetDlgItem(btnUpdate)->EnableWindow(FALSE);
			GetDlgItem(btnSave)->EnableWindow(FALSE);
			GetDlgItem(btnCancel)->EnableWindow(FALSE);
			GetDlgItem(btnAllSelect)->EnableWindow(FALSE);
			GetDlgItem(btnInput)->EnableWindow(TRUE);

			GetDlgItem(IDC_PASSWORD)->SetFocus();
		}
		END_CATCH;
	}
}


void CMdbToXlDlg::OnBnClickedbtninput()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	m_ctrlTable.ResetContent();

	CloseDBConn(m_bConn);

	m_nXlRowNum = 0;

	m_stFieldInfo.nFieldIdcnt = 0;
	m_stFieldInfo.nExcelIdcnt = 0;
	m_stFieldInfo.strExcelName = _T("");
	m_stFieldInfo.strFieldName = _T("");

	UpdateData(1);// 컨트롤 >> 변수 // 비밀번호 가져오기 // 테이블 변경 시 필드 목록 재설정

	TRY
	{
		//DB연결
		// //B1594C47 비밀번호
		CString strtemp;
		strtemp.Format(_T("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s; PWD=%s"),
					m_strPathName, m_strPW);

		m_DB.OpenEx(strtemp, CDatabase::noOdbcDialog);
		//테이블 이름 불러오기
		CallDBTable();

		// TABLE에 포함된 FIELD값들을 가져와 표시하는 부분
		CODBCFieldInfo fieldInfo;
		CString strQuery;

		UpdateData(1);// 컨트롤 >> 변수 // 테이블 변경 시 필드 목록 재설정

		TRY
		{
			m_pRecordset = new CRecordset(&m_DB);

			m_nTotalRecordCount = CalcTotalRow(m_strTable, &m_DB);

			strQuery.Format(_T("SELECT * FROM %s"), m_strTable);
			m_pRecordset->Open(CRecordset::dynaset, strQuery);
			int FCount = m_pRecordset->GetODBCFieldCount();

			m_bConn = TRUE;
			for (int i = 0; i < FCount; i++)
			{
				m_pRecordset->GetODBCFieldInfo(i, fieldInfo);
				m_stFieldInfo.strFieldName = fieldInfo.m_strName;
				m_stFieldInfo.nFieldIdcnt = i;

				m_ctrlFieldList.InsertItem(i, m_stFieldInfo.strFieldName);
				m_VstTrueFieldValue.push_back(m_stFieldInfo);
			}
			GetDlgItem(btnAllSelect)->EnableWindow(TRUE);
		}
		CATCH(CDBException, e)
		{
			GetDlgItem(btnAllSelect)->EnableWindow(FALSE);
			GetDlgItem(IDC_PASSWORD)->EnableWindow(FALSE);
			AfxMessageBox(_T("해당 테이블이 존재하지 않습니다."));
		}END_CATCH;

		//테이블이 SV_DVVAL일 때
		FirstInput(_T("SV_DVVAL"));
		//////////////////////////////////////////////////////////////////////////
		m_bConn = TRUE;
		GetDlgItem(btnInput)->EnableWindow(FALSE);
		GetDlgItem(btnAllSelect)->EnableWindow(TRUE);
		SetDlgItemText(IDC_PASSWORD, _T(""));
		GetDlgItem(IDC_PASSWORD)->EnableWindow(FALSE);
	}
	CATCH(CDBException, e)
		AfxMessageBox(_T("비밀번호가 유효하지 않습니다."));
	END_CATCH;
}


void CMdbToXlDlg::OnCbnSelchangeTable()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	CloseDBConn();

	UpdateData(1);// 컨트롤 >> 변수 //테이블 변경 시 필드 목록 재설정

	TRY
	{
		CODBCFieldInfo fieldInfo;
		CString strQuery;
		/////////////////////////////////
		
		//////////////////////////////////
		strQuery.Format(_T("SELECT * FROM %s"), m_strTable);
		m_pRecordset->Open(CRecordset::dynaset, strQuery);
		int FCount = m_pRecordset->GetODBCFieldCount();

		m_bConn = TRUE;
		m_nTotalRecordCount = CalcTotalRow(m_strTable, &m_DB);
		for (int i = 0; i < FCount; i++)
		{
			m_pRecordset->GetODBCFieldInfo(i, fieldInfo);
			m_stFieldInfo.strFieldName = fieldInfo.m_strName;
			m_stFieldInfo.nFieldIdcnt = i;

			m_ctrlFieldList.InsertItem(i, m_stFieldInfo.strFieldName);
			m_VstTrueFieldValue.push_back(m_stFieldInfo);

		}
		GetDlgItem(btnInput)->EnableWindow(FALSE);
		GetDlgItem(btnAllSelect)->EnableWindow(TRUE);
		m_pRecordset->Close();
	}
		CATCH(CDBException, e)
	{
		GetDlgItem(btnSave)->EnableWindow(FALSE);
		GetDlgItem(btnAllSelect)->EnableWindow(FALSE);
		GetDlgItem(btnCancel)->EnableWindow(FALSE);
		AfxMessageBox(_T("해당 테이블이 존재하지 않습니다."));
	}
	END_CATCH;

	FirstInput(_T("SV_DVVAL"));

	m_bConn = TRUE;
	SetDlgItemText(IDC_PASSWORD, _T(""));
	GetDlgItem(btnInput)->EnableWindow(FALSE);
	GetDlgItem(btnAllSelect)->EnableWindow(TRUE);
}

BOOL CMdbToXlDlg::KillProcess(CString sProcessName)
{
	HANDLE         hProcessSnap = NULL;
	DWORD          Return = FALSE;
	PROCESSENTRY32 pe32 = { 0 };

	CString ProcessName = sProcessName;
	ProcessName.MakeLower();

	hProcessSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

	if (hProcessSnap == INVALID_HANDLE_VALUE)
		return (DWORD)INVALID_HANDLE_VALUE;

	pe32.dwSize = sizeof(PROCESSENTRY32);

	if (Process32First(hProcessSnap, &pe32))
	{
		DWORD Code = 0;
		DWORD dwPriorityClass;

		do {
			HANDLE hProcess;
			hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, pe32.th32ProcessID);
			dwPriorityClass = GetPriorityClass(hProcess);

			CString Temp = pe32.szExeFile;
			Temp.MakeLower();

			if (Temp == ProcessName)
			{
				if (TerminateProcess(hProcess, 0))
					GetExitCodeProcess(hProcess, &Code);
				else
					return GetLastError();
			}
			CloseHandle(hProcess);
		} while (Process32Next(hProcessSnap, &pe32));
		Return = TRUE;
	}
	else
	{
		Return = FALSE;
	}

	CloseHandle(hProcessSnap);

	return Return;
}

int CMdbToXlDlg::CalcTotalRow(CString strTableName, CDatabase* db)
{
	CString strQuery2;
	CDBVariant vtval;
	CRecordset* pRecordset;

	pRecordset = new CRecordset(db);
	int TotalRecord = 0;

	strQuery2.Format(_T("SELECT COUNT(*) FROM %s"), strTableName);

	pRecordset->Open(CRecordset::forwardOnly, strQuery2);

	pRecordset->GetFieldValue((short)0, vtval);
	TotalRecord = vtval.m_lVal;

	pRecordset->Close();
	delete pRecordset;

	return TotalRecord;
}

DWORD WINAPI MDBtoExcelWorkThread(LPVOID p)
{
	CMdbToXlDlg* pMainWnd = (CMdbToXlDlg*)p;
	vector<FieldINFO> VstExcelValClone = pMainWnd->m_VstExcelValue;

	int nRow = 0;
	int nColumn = 0;
	CString strData; // CELL에 입력될 값을 받을 변수
	
	pMainWnd->GetDlgItem(btnFileLoad)->EnableWindow(FALSE);
	CXLEzAutomation ExcelServer(FALSE);

	//첫번째 ROW 필드명들만 삽입
	int nFieldCount = (int)VstExcelValClone.size();
	for (int i = 0; i < nFieldCount; i++)
	{
		//ExcelServer.SetCellValue(i + 1, 1, pMainWnd->m_VstExcelValue.at(i).strExcelName); //SetCellValue : 셀의 내용 설정 
		ExcelServer.SetCellValue(i + 1, 1, VstExcelValClone.at(i).strExcelName); //SetCellValue : 셀의 내용 설정
	}

	pMainWnd->m_ctrlProgress.SetRange(0, pMainWnd->m_nTotalRecordCount); //프로그래스 바 범위 
	//-------------------------------------------------------------------------------
	//pMainWnd->m_pRecordset->MoveFirst();
	int nSetCnt = 0; // 프로그래스 바
	nRow = 2; // 2행부터 출력하기 위해 Row를 2로 잡음(함수 -> 1부터 시작)
	while (!pMainWnd->m_pRecordset->IsEOF())
	{
		for (int i = 0; i < nFieldCount; i++)
		{
			pMainWnd->m_pRecordset->GetFieldValue(short(VstExcelValClone.at(i).nFieldIdcnt), strData);
			ExcelServer.SetCellValue(i + 1, nRow, strData);
		}
		nSetCnt = nSetCnt + 1;
		pMainWnd->m_ctrlProgress.SetPos(nSetCnt); // 프로그래스 바 진행상황
		pMainWnd->m_pRecordset->MoveNext();
		nRow++;
		pMainWnd->SetDlgItemText(IDC_SAVE_PROGRESS, _T("데이터 변환 중..."));
	}

	bool bSaveXL = ExcelServer.SaveFileAs(pMainWnd->m_strExcelPathName); //SaveFileAs : 엑셀 파일 저장
	if (bSaveXL == true)
	{
		pMainWnd->SetDlgItemText(IDC_SAVE_PROGRESS, _T("파일 저장 완료"));
		pMainWnd->m_ctrlProgress.SetPos(0); // 프로그래스 바 종료
	}
	pMainWnd->GetDlgItem(btnFileLoad)->EnableWindow(TRUE);

	pMainWnd->SetDlgItemText(btnSave, _T("저장하기"));
	pMainWnd->m_pRecordset->MoveFirst();

	pMainWnd->m_pRecordset->Close();
	ExcelServer.ReleaseExcel();	// 엑셀파일 종료
	CloseHandle(pMainWnd->m_hExcelThread);		// 핸들 제거
	pMainWnd->m_hExcelThread = NULL;			// 핸들 초기화

	return 0;
}

void CMdbToXlDlg::OnBnClickedbtnsave()
{
	if (m_hExcelThread != NULL) //스레드가 동작하지 않을 경우
	{		
		AfxMessageBox(_T("파일 저장 중입니다."));
	}
	else
	{
		CFileDialog dlg(false, _T("xlsx"), m_strFileName,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_NOCHANGEDIR,
			_T("xlsx 파일 (*.xlsx)|*.xlsx"), NULL);
		if (dlg.DoModal() != IDOK)
		{
			return; //다이얼로그 취소 시 다시 모달로 띄울수 있도록 에러 처리
		}
		m_strExcelPathName = dlg.GetPathName();
		//CloseHandle(CreateThread(NULL, 0, MDBtoExcelWorkThread, this, 0, 0));
		m_hExcelThread = CreateThread(NULL, 0, MDBtoExcelWorkThread, this, 0, 0);
	}
}


DWORD WINAPI XJemtoExcelWorkThread(LPVOID p)
{
	CMdbToXlDlg* pMainWnd = (CMdbToXlDlg*)p;
	
	vector<FieldINFO> VstExcelValClone = pMainWnd->m_VstExcelValue;

	CString tempPame = pMainWnd->m_strExcelPathName;

	wchar_t strUnicode[256] = { 0, };
	char   strMultibyte[256] = { 0, };
	wcscpy_s(strUnicode, 256, tempPame);
	int len = WideCharToMultiByte(CP_UTF8, 0, strUnicode, -1, NULL, 0, NULL, NULL);
	WideCharToMultiByte(CP_UTF8, 0, strUnicode, -1, strMultibyte, len, NULL, NULL);

	lxw_workbook* workbook = workbook_new(strMultibyte); // 안시
	lxw_worksheet* worksheet = workbook_add_worksheet(workbook, NULL);
	worksheet_set_vba_name(worksheet, "asd");

	int nFieldCount = VstExcelValClone.size();

	for (int i = 0; i < nFieldCount; i++)
	{
		CStringA temp = CStringA(VstExcelValClone.at(i).strExcelName);
		const char* cFullAddr = temp;
		char* cpFullAddr = const_cast<char*>(cFullAddr);

		if (pMainWnd->m_strTable == _T("SV_DVVAL"))
		{
			lxw_format* format = workbook_add_format(workbook);
			format_set_bold(format);
			format_set_font_color(format, LXW_COLOR_WHITE); // 폰트 흰색

			format_set_pattern(format, LXW_PATTERN_SOLID);
			format_set_bg_color(format, LXW_COLOR_BLACK); // 셀 색상 검은색

			format_set_align(format, LXW_ALIGN_CENTER);
			format_set_align(format, LXW_ALIGN_VERTICAL_CENTER);
			worksheet_write_string(worksheet, 0, i, cpFullAddr, format);
		}
		else
		{
			return 0;
		}
	}
	int nValueCount = 0; // 전체 count가 필요함

	while (!pMainWnd->m_pRecordset->IsEOF())
	{
		for (int i = 0; i < nFieldCount; i++) //필드 갯수 만큼만 반복
		{
			nValueCount++; //필드 갯수 * Row 만큼 반복
		}
		pMainWnd->m_pRecordset->MoveNext();
	}
	pMainWnd->m_ctrlProgress.SetRange(0, nValueCount); //프로그래스 바 범위 

	CString strData = _T("");
	pMainWnd->m_pRecordset->MoveFirst();
	int nSetCnt = 0; // 프로그래스 바
	int nRow = 1; // 2행부터 출력하기 위해 Row를 2로 잡음(함수 -> 1부터 시작)

	while (!pMainWnd->m_pRecordset->IsEOF())
	{
		for (int i = 0; i < nFieldCount; i++)
		{
			pMainWnd->m_pRecordset->GetFieldValue(short(VstExcelValClone.at(i).nFieldIdcnt), strData);

			wcscpy_s(strUnicode, 256, strData);
			int len = WideCharToMultiByte(CP_UTF8, 0, strUnicode, -1, NULL, 0, NULL, NULL);
			WideCharToMultiByte(CP_UTF8, 0, strUnicode, -1, strMultibyte, len, NULL, NULL);

			worksheet_write_string(worksheet, nRow, i, strMultibyte, NULL); // 안시
			nSetCnt = nSetCnt + 1;
			pMainWnd->m_ctrlProgress.SetPos(nSetCnt); // 프로그래스 바 진행상황
		}
		pMainWnd->m_pRecordset->MoveNext();
		nRow++;
		pMainWnd->SetDlgItemText(IDC_SAVE_PROGRESS, _T("데이터 변환 중..."));
	}
		//테두리	
	lxw_format* format = workbook_add_format(workbook);
	format_set_border(format, LXW_BORDER_THICK);
	
	workbook_close(workbook);
	pMainWnd->SetDlgItemText(IDC_SAVE_PROGRESS, _T("저장 완료"));

	return 0;
}

void CMdbToXlDlg::OnBnClickedDatasaveXjem()
{
	if (m_hExcelThread != NULL) //스레드가 동작하지 않을 경우
	{
		AfxMessageBox(_T("파일 저장 중입니다."));
	}
	else
	{
		CFileDialog dlg(false, _T("xlsx"), m_strFileName,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_NOCHANGEDIR,
			_T("xlsx 파일 (*.xlsx)|*.xlsx"), NULL);
		if (dlg.DoModal() != IDOK)
		{
			return; //다이얼로그 취소 시 다시 모달로 띄울수 있도록 에러 처리
		}
		m_strExcelPathName = dlg.GetPathName();
		//CloseHandle(CreateThread(NULL, 0, MDBtoExcelWorkThread, this, 0, 0));
		m_hExcelThread = CreateThread(NULL, 0, XJemtoExcelWorkThread, this, 0, 0);
	}
}

void CMdbToXlDlg::OnLvnItemchangedFieldlist(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);

	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	int nCount = m_ctrlFieldList.GetItemCount();

	for (int i = nCount - 1; i >= 0; i--)
	{
		if (m_ctrlFieldList.GetCheck(i) == TRUE) // 체크가 되어있을 때 + 체크할 때
		{
			CString strtemp;
			strtemp = m_ctrlFieldList.GetItemText(i, 0);

			FieldINFO sttmp;
			sttmp.nExcelIdcnt = m_nXlRowNum;
			sttmp.nFieldIdcnt = m_VstTrueFieldValue.at(i).nFieldIdcnt;
			sttmp.strFieldName = m_VstTrueFieldValue.at(i).strFieldName;
			sttmp.strExcelName = m_VstTrueFieldValue.at(i).strFieldName;

			BOOL bDistinct = FALSE; //중복 여부

			for (int j = m_ctrlExcelList.GetItemCount() - 1; j >= 0; j--)
			{
				if (strtemp == m_VstExcelValue.at(j).strFieldName)
				{
					bDistinct = TRUE;
					break;
				}
			}
			if (bDistinct != TRUE)
			{
				//excellist에 생성하는 부분
				m_ctrlExcelList.InsertItem(m_ctrlExcelList.GetItemCount(), strtemp, NULL);
				m_VstExcelValue.push_back(sttmp);
				m_nXlRowNum++;
			}
		}
		else//체크가 안되어 있을 때 + 체크가 풀릴 때
		{
			for (int j = m_ctrlExcelList.GetItemCount() - 1; j >= 0; j--)
			{
				if (m_VstTrueFieldValue.at(i).nFieldIdcnt == m_VstExcelValue.at(j).nFieldIdcnt)
				{
					m_ctrlExcelList.DeleteItem(m_VstExcelValue.at(j).nExcelIdcnt);
					m_VstExcelValue.erase(m_VstExcelValue.begin() + j);
					m_nXlRowNum -= 1;

					SetDlgItemText(IDC_UPDATE_NAME, _T(""));
					GetDlgItem(btnUpdate)->EnableWindow(FALSE);
					GetDlgItem(IDC_UPDATE_NAME)->EnableWindow(FALSE);
					for (int a = j; a < m_VstExcelValue.size(); a++)
					{
						m_VstExcelValue.at(a).nExcelIdcnt -= 1;
					}
				}
			}
		}
	}
	if (nCount > 0)
	{
		GetDlgItem(btnAllSelect)->EnableWindow(TRUE);
		GetDlgItem(btnCancel)->EnableWindow(TRUE);
		GetDlgItem(btnSave)->EnableWindow(TRUE);
	}
	else
	{
		GetDlgItem(btnCancel)->EnableWindow(FALSE);
		GetDlgItem(btnSave)->EnableWindow(FALSE);
	}

	if (m_ctrlExcelList.GetItemCount() == nCount)
		GetDlgItem(btnAllSelect)->EnableWindow(FALSE);

	if (m_VstExcelValue.size() == 0)
	{
		GetDlgItem(btnSave)->EnableWindow(FALSE);
		GetDlgItem(btnCancel)->EnableWindow(FALSE);
	}

	GetDlgItem(IDC_ExcelList)->SendMessage(WM_KILLFOCUS, NULL);
	GetDlgItem(IDC_FieldList)->SendMessage(WM_KILLFOCUS, NULL);
	*pResult = 0;
}

void CMdbToXlDlg::OnBnClickedbtnupdate()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	UpdateData(1);// 컨트롤 >> 변수 // 버튼 클릭 시 변경하기 위한 데이터 가져오기
	m_ctrlExcelList.SetItemText(m_nidx, 0, m_strEdit);
	m_VstExcelValue.at(m_nidx).strExcelName = m_strEdit;
	LVCOLUMNW lvcol;

	lvcol.pszText = (LPWSTR)(LPCTSTR)m_strEdit;
	UpdateData(0);// 컨트롤 << 변수 //edit 창에 이름 띄움
}

void CMdbToXlDlg::OnNMClickExcellist(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	*pResult = 0;
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
	m_nidx = pNMListView->iItem;

	m_strEdit.SetString(m_ctrlExcelList.GetItemText(m_nidx, 0));

	//excellist 행 클릭시 같은 정보의 Fieldlist 행 선택
	if (m_nidx != -1) // 빈 공간 클릭시 m_nidx에 -1이 들어옴
	{
		if (m_nidx < m_ctrlExcelList.GetItemCount())
		{
			m_ctrlFieldList.SetSelectionMark(m_VstExcelValue.at(m_nidx).nFieldIdcnt);
			m_ctrlFieldList.SetItemState(m_VstExcelValue.at(m_nidx).nFieldIdcnt, LVIS_SELECTED | LVIS_FOCUSED,
				LVIS_SELECTED | LVIS_FOCUSED);

			m_ctrlFieldList.SetFocus();
		}
	}

	if (m_strEdit == _T(""))
	{
		GetDlgItem(btnUpdate)->EnableWindow(FALSE);
		GetDlgItem(IDC_UPDATE_NAME)->EnableWindow(FALSE);
	}
	else
	{
		GetDlgItem(btnUpdate)->EnableWindow(TRUE);
		GetDlgItem(IDC_UPDATE_NAME)->EnableWindow(TRUE);
	}
	UpdateData(0);// 컨트롤 << 변수 // edit창에 이름 박기
}

void CMdbToXlDlg::OnBnClickedbtnallselect()
{
	for (int i = 0; i < m_ctrlFieldList.GetItemCount(); i++)
	{
		m_ctrlFieldList.SetCheck(i);
	}
}

void CMdbToXlDlg::OnBnClickedbtncancel()
{
	for (int i = 0; i < m_ctrlFieldList.GetItemCount(); i++)
	{
		m_ctrlFieldList.SetCheck(i, FALSE);
	}
}

void CMdbToXlDlg::OnNMClickFieldlist(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	*pResult = 0;
	int idx = pNMItemActivate->iItem;

	if (idx != -1)
	{
		for (int a = 0; a < m_ctrlExcelList.GetItemCount(); a++)
		{
			if (m_VstTrueFieldValue.at(idx).nFieldIdcnt == m_VstExcelValue.at(a).nFieldIdcnt)
			{
				m_ctrlExcelList.SetSelectionMark(m_VstExcelValue.at(a).nExcelIdcnt);
				m_ctrlExcelList.SetItemState(m_VstExcelValue.at(a).nExcelIdcnt, LVIS_SELECTED | LVIS_FOCUSED,
					LVIS_SELECTED | LVIS_FOCUSED);

				m_ctrlExcelList.SetFocus();
			}
		}
	}
}

BOOL CMdbToXlDlg::PreTranslateMessage(MSG* pMsg)
{
	// TODO: 여기에 특수화된 코드를 추가 및/또는 기본 클래스를 호출합니다.

	return CDialogEx::PreTranslateMessage(pMsg);
	if (WM_KEYDOWN == pMsg->message)
	{
		if (VK_RETURN == pMsg->wParam || VK_ESCAPE == pMsg->wParam)
		{
			return TRUE;
		}
	}
	return CDialogEx::PreTranslateMessage(pMsg);
}

void CMdbToXlDlg::OnOK()
{
	//Enter 클릭시 종료되는 상황을 막기 위해, 아래 소스 막음
	//CDialogEx::OnOK();
}


void CMdbToXlDlg::OnCancel()
{
	// ESC 클릭시 종료되는 상황을 막기 위해, 아래 소스 막음
	// :  X버튼에 대해서는 열어줄 필요가 있음
	//CDialogEx::OnCancel();
}
