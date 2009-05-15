////////////////////////////////////////////////////////////////////////////////
//
// Copyright (c) 2004 Sergey Solozhentsev
// Author: 	Sergey Solozhentsev e-mail: salos@mail.ru
// Product:	WTL Helper
// File:      	SizeListCtrl.cpp
// Created:	13.01.2005 14:32
// 
//   Using this software in commercial applications requires an author
// permission. The permission will be granted to everyone excluding the cases
// when someone simply tries to resell the code.
//   This file may be redistributed by any means PROVIDING it is not sold for
// profit without the authors written consent, and providing that this notice
// and the authors name is included.
//   This file is provided "as is" with no expressed or implied warranty. The
// author accepts no liability if it causes any damage to you or your computer
// whatsoever.
//
////////////////////////////////////////////////////////////////////////////////

// This file was generated by WTL subclass control wizard 
// SizeListCtrl.cpp : Implementation of CSizeListCtrl

#include "stdafx.h"
#include "SizeListCtrl.h"

// CSizeListCtrl
CSizeListCtrl::CSizeListCtrl() : m_PrevWidth(0)
{
}

CSizeListCtrl::~CSizeListCtrl()
{
}


DWORD CSizeListCtrl::OnPrePaint(int idCtrl, LPNMCUSTOMDRAW lpNMCustomDraw)
{
	return CDRF_NOTIFYITEMDRAW;
}

DWORD CSizeListCtrl::OnItemPrePaint(int idCtrl, LPNMCUSTOMDRAW lpNMCustomDraw)
{
	LPNMLVCUSTOMDRAW lpLvCustomDraw = (LPNMLVCUSTOMDRAW)lpNMCustomDraw;
	LPARAMSTRUCT* pParam = (LPARAMSTRUCT*)lpLvCustomDraw->nmcd.lItemlParam;
	if (pParam->DrawStruct.BackColor == CLR_NONE && 
		pParam->DrawStruct.TextColor == CLR_NONE &&
		!pParam->DrawStruct.dwStyle)
	{
		return CDRF_DODEFAULT;
	}
	if (pParam->DrawStruct.BackColor != CLR_NONE)
	{
		lpLvCustomDraw->clrTextBk = pParam->DrawStruct.BackColor;
	}
	if (pParam->DrawStruct.TextColor != CLR_NONE)
	{
		lpLvCustomDraw->clrText = pParam->DrawStruct.TextColor;
	}
	DWORD Res = CDRF_NEWFONT;
	if ((pParam->DrawStruct.dwMask & MCDS_STYLE) && pParam->DrawStruct.dwStyle)
	{
		Res |= CDRF_NOTIFYSUBITEMDRAW;
	}

	return Res;
}

DWORD CSizeListCtrl::OnSubItemPrePaint(int idCtrl, LPNMCUSTOMDRAW lpNMCustomDraw)
{
	LPNMLVCUSTOMDRAW lpLvCustomDraw = (LPNMLVCUSTOMDRAW)lpNMCustomDraw;
	LPARAMSTRUCT* pParam = (LPARAMSTRUCT*)lpLvCustomDraw->nmcd.lItemlParam;
	CDCHandle dc = lpLvCustomDraw->nmcd.hdc;
	if (lpLvCustomDraw->iSubItem == 0)
	{
		LOGFONT lf;
		DWORD dwStyle = pParam->DrawStruct.dwStyle;
		::GetObject(GetFont(), sizeof(LOGFONT), &lf);
		if (dwStyle & SLC_STYLE_ITALIC)
		{
			lf.lfItalic = TRUE;
		}
		if (dwStyle & SLC_STYLE_UNDERLINE)
		{
			lf.lfUnderline = TRUE;
		}
		if (dwStyle & SLC_STYLE_STRIKEOUT)
		{
			lf.lfStrikeOut = TRUE;
		}
		if (dwStyle & SLC_STYLE_BOLD)
		{
			lf.lfWeight = FW_BOLD;
		}
		
		if (m_NewFont.IsNull() == FALSE) {
			m_NewFont.Detach(); 
		}
		m_NewFont.CreateFontIndirect(&lf);
		
		HFONT hTmp = dc.SelectFont(m_NewFont);

		if (m_OldFont.IsNull()) {
			m_OldFont = hTmp;
		} else { 
			::DeleteObject((HGDIOBJ)hTmp);
		}
	}
	else
	{
		if (m_NewFont.m_hFont)
		{
			dc.SelectFont(m_OldFont.Detach());
			m_NewFont.DeleteObject();
		}
	}
	return CDRF_DODEFAULT;
}

void CSizeListCtrl::DeleteHelperItemData(LPARAM lParam)
{
	LPARAMSTRUCT* pParam = (LPARAMSTRUCT*)lParam;
	delete pParam;
}

LPARAM CSizeListCtrl::_GetItemData(int nItem)
{
	ATLASSERT(::IsWindow(m_hWnd));
	LVITEM lvi = { 0 };
	lvi.iItem = nItem;
	lvi.mask = LVIF_PARAM;
	if (CallWindowProc(m_pfnSuperWindowProc, m_hWnd, LVM_GETITEM, 0, (LPARAM)&lvi))
	{
		return lvi.lParam;
	}
	return 0;
}

BOOL CSizeListCtrl::SetDrawStruct(int nItem, const MyCustomDrawStruct* pDrawStruct)
{
	LPARAMSTRUCT* pParam = (LPARAMSTRUCT*)_GetItemData(nItem);
	if (!pParam)
		return FALSE;
	if (pDrawStruct->dwMask & MCDS_BACKCOLOR)
	{
		pParam->DrawStruct.BackColor = pDrawStruct->BackColor;
		pParam->DrawStruct.dwMask |= MCDS_BACKCOLOR;
	}
	if (pDrawStruct->dwMask & MCDS_TEXTCOLOR)
	{
		pParam->DrawStruct.TextColor = pDrawStruct->TextColor;
		pParam->DrawStruct.dwMask |= MCDS_TEXTCOLOR;
	}
	if (pDrawStruct->dwMask & MCDS_STYLE)
	{
		pParam->DrawStruct.dwStyle = pDrawStruct->dwStyle;
		pParam->DrawStruct.dwMask |= MCDS_STYLE;
	}
	return TRUE;
}

BOOL CSizeListCtrl::GetDrawStruct(int nItem, MyCustomDrawStruct* pDrawStruct)
{
	LPARAMSTRUCT* pParam = (LPARAMSTRUCT*)_GetItemData(nItem);
	if (!pParam)
		return FALSE;
	*pDrawStruct = pParam->DrawStruct;
	return TRUE;
}
//////////////////////////////////////////////////////////////////////////

LRESULT CSizeListCtrl::OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	CRect rc;
	GetWindowRect(rc);
	int Width = rc.Width();
	if (Width != m_PrevWidth)
	{
		m_PrevWidth = Width;
	PostMessage(SLC_POSTSIZE, wParam, lParam);
	}
	return DefWindowProc();
}

LRESULT CSizeListCtrl::OnSetItem(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	LPLVITEM pItem = (LPLVITEM)lParam;
	LVITEM Item;
	if (pItem->mask & LVIF_PARAM)
	{
		Item = *pItem;
		lParam = (LPARAM)&Item;
		LPARAMSTRUCT* pParam = (LPARAMSTRUCT*)_GetItemData(pItem->iItem);
		if (!pParam)
			return 0;
		pParam->lOldParam = pItem->lParam;
		Item.lParam = (LPARAM)pParam;
	}

	bHandled = TRUE;
	return DefWindowProc(uMsg, wParam, lParam);
}

LRESULT CSizeListCtrl::OnInsertItem(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	LPLVITEM pItem = (LPLVITEM)lParam;
	LVITEM Item = *pItem;
	Item.mask |= LVIF_PARAM;
	LPARAMSTRUCT* pParam = new LPARAMSTRUCT;
	pParam->lOldParam = 0;
	pParam->DrawStruct.BackColor = CLR_NONE;
	pParam->DrawStruct.dwStyle = 0;
	pParam->DrawStruct.TextColor = CLR_NONE;
	if (pItem->mask & LVIF_PARAM)
	{
		pParam->lOldParam = pItem->lParam;
	}
	
	Item.lParam = (LPARAM)pParam;
	lParam = (LPARAM)&Item;
	return DefWindowProc(uMsg, wParam, lParam);
}

LRESULT CSizeListCtrl::OnGetItem(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	if (!DefWindowProc(uMsg, wParam, lParam))
		return FALSE;
	LPLVITEM pItem = (LPLVITEM)lParam;
	if (pItem->mask & LVIF_PARAM)
	{
		LPARAMSTRUCT* pParam = (LPARAMSTRUCT*)pItem->lParam;
		pItem->lParam = pParam->lOldParam;
	}
	return TRUE;
}

LRESULT CSizeListCtrl::OnDeleteItem(int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
	bHandled = FALSE;
	LPNMLISTVIEW plvn = (LPNMLISTVIEW)pnmh;
	DeleteHelperItemData(_GetItemData(plvn->iItem));
	return 0;
}

LRESULT CSizeListCtrl::OnPostSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	bHandled = TRUE;

	CRect rc;
	GetClientRect(rc);
	int CtrlWidth = rc.Width();

	int Width = GetColumnWidth(0);
	SetColumnWidth(1, CtrlWidth - Width - 5);
	CString str;
	str.Format(_T("CSizeListCtrl::OnPostSize CtrlWidth = %d CtrlrHeight = %d\r\n"), CtrlWidth, rc.Height()); 
	ATLTRACE(str);
	return 0;
}
