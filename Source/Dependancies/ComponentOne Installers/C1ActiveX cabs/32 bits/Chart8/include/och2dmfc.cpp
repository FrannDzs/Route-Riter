/******************************************************************************
*
* Copyright (c) ComponentOne, LLC.  All Rights Reserved.
* Portions copyright (c) 1999, KL GROUP INC.
* http://www.componentone.com
*
* This software is the confidential and proprietary information of ComponentOne
* LLC. ("Confidential Information").  You shall not disclose such
* Confidential Information and shall use it only in accordance with the
* terms of the license agreement you entered into with ComponentOne.
*
* COMPONENTONE MAKES NO REPRESENTATIONS OR WARRANTIES ABOUT THE SUITABILITY
* OF THE SOFTWARE, EITHER EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
* TO THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
* PURPOSE, OR NON-INFRINGEMENT. COMPONENTONE SHALL NOT BE LIABLE FOR ANY
* DAMAGES SUFFERED BY LICENSEE AS A RESULT OF USING, MODIFYING OR
* DISTRIBUTING THIS SOFTWARE OR ITS DERIVATIVES.
*
******************************************************************************/

/*
 *  Implement all the necessary functions for the CChart2D MFC class
 */

#include "och2dmfc.h"

BOOL CChart2D::Create(LPCTSTR lpszCaption, DWORD dwStyle,	const RECT& rect, 
		CWnd* pParentWnd, UINT nID)
{
    BOOL bResult(FALSE);

    if(CWnd::Create(XRT2D, lpszCaption, dwStyle, rect,pParentWnd, nID))
    {
        m_hChart = XrtCreate();
        if(m_hChart)
        {
            XrtAttachWindow(m_hChart, m_hWnd);
            bResult = TRUE;
        } 
    } 

    return(bResult);
}


#if (_MFC_VER < 0x0300)
//////////////////
// Return place to hold original window proc
//
WNDPROC* CChart2D::GetSuperWndProcAddr()
{
    static WNDPROC NEAR pfnSuper;   // place to store
                                    // window proc
    return &pfnSuper;               // always return the
                                    // same address
}
#endif
  
IMPLEMENT_DYNAMIC(CChart2D, CWnd)
