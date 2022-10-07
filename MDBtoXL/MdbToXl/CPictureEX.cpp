#include "pch.h"
#include "CPictureEX.h"

IMPLEMENT_DYNAMIC(CPictureEx, CStatic)

CPictureEx::CPictureEx()
{
	m_colorTransparent = RGB(255, 255, 255);
	//m_colorTransparent = RGB(0, 0, 0);	
}

CPictureEx::~CPictureEx()
{
}

BEGIN_MESSAGE_MAP(CPictureEx, CStatic)
	ON_WM_PAINT()
END_MESSAGE_MAP()

// CPictureEx 메시지 처리기입니다.

void CPictureEx::OnPaint()
{
	CPaintDC dc(this); // device context for painting
	// TODO: 여기에 메시지 처리기 코드를 추가합니다.
	HBITMAP old, bmp = GetBitmap();
	BITMAP bminfo;
	::GetObject(bmp, sizeof(BITMAP), &bminfo);

	CDC memDC;
	memDC.CreateCompatibleDC(&dc);
	old = (HBITMAP)::SelectObject(memDC.m_hDC, bmp);
	::TransparentBlt(dc.m_hDC, 0, 0, bminfo.bmWidth, bminfo.bmHeight, memDC.m_hDC, 0, 0, bminfo.bmWidth, bminfo.bmHeight, m_colorTransparent);
	::SelectObject(memDC.m_hDC, old);
	memDC.DeleteDC();
	// 그리기 메시지에 대해서는 CStatic::OnPaint()을(를) 호출하지 마십시오.
}