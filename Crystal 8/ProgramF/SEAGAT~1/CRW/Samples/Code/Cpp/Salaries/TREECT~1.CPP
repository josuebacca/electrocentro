/////////////////////////////////////////////////////////////////////////////
// TreeCtrlCheckBox.cpp : implementation file for the Custom Tree Control Class
// used in the Selection form
//

#include "stdafx.h"
#include "ProSalaries.h"
#include "TreeCtrlCheckBox.h"

/////////////////////////////////////////////////////////////////////////////
// CTreeCtrlCheckBox

CTreeCtrlCheckBox::CTreeCtrlCheckBox()
{
}

CTreeCtrlCheckBox::~CTreeCtrlCheckBox()
{
}


BEGIN_MESSAGE_MAP(CTreeCtrlCheckBox, CTreeCtrl)
	//{{AFX_MSG_MAP(CTreeCtrlCheckBox)
	ON_WM_LBUTTONDOWN()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTreeCtrlCheckBox message handlers

/////////////////////////////////////////////////////////////////////////////
// Check and uncheck the item's corresponding checkbox if the item text is clicked
// on (not just the check box itself)

void CTreeCtrlCheckBox::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CTreeCtrl::OnLButtonDown(nFlags, point);

    UINT uFlags=0;	
    HTREEITEM hti = HitTest(point, &uFlags);

    if( uFlags & TVHT_ONITEMSTATEICON)	
    {
        UINT checkState = GetItemState(hti, TVIS_STATEIMAGEMASK)>>12;
        CheckNode(hti, checkState);
        return;	
    }
}

/////////////////////////////////////////////////////////////////////////////
// When a parent node is checked/unchecked, recursively check/uncheck
// all child nodes to be equivalent to the parent

void CTreeCtrlCheckBox::CheckNode(HTREEITEM nodh, UINT checkState)
{
    // Recursively make all child nodes the same state as the parent node.
    if (nodh == 0) return;
    
    if (ItemHasChildren(nodh))
    {
        HTREEITEM hti = GetChildItem(nodh);

        while (hti != 0)
        {
            SetItemState(hti, INDEXTOSTATEIMAGEMASK(checkState), TVIS_STATEIMAGEMASK);
            CheckNode(hti, checkState);
            hti = GetNextSiblingItem(hti);
        }
    }
}
