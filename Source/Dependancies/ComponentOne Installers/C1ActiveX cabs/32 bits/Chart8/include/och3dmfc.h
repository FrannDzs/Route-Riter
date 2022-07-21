#if !defined(__OCH3DMFC_H__)
#define __OCH3DMFC_H__

/******************************************************************************
*
* Copyright (c) ComponentOne, LLC.  All Rights Reserved.
* Portions copyright (c) 1999, KL GROUP INC.
* http://www.componentone.com
*
* This software is the confidential and proprietary information of
* ComponentOne, LLC. ("Confidential Information").  You shall not disclose
* such Confidential Information and shall use it only in accordance with the
* terms of the license agreement you entered into with ComponentOne, LLC.
*
* COMPONENTONE, LLC MAKES NO REPRESENTATIONS OR WARRANTIES ABOUT THE
* SUITABILITY OF THE SOFTWARE, EITHER EXPRESS OR IMPLIED, INCLUDING BUT NOT
* LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
* PARTICULAR PURPOSE, OR NON-INFRINGEMENT. COMPONENTONE, LLC SHALL NOT BE LIABLE
* FOR ANY DAMAGES SUFFERED BY LICENSEE AS A RESULT OF USING, MODIFYING OR
* DISTRIBUTING THIS SOFTWARE OR ITS DERIVATIVES.
*
******************************************************************************/

/* ========================================================================
 *                 How to use the MFC wrapper classes:
 * ========================================================================
 * The MFC wrapper classes below provide a fairly flat wrapper around
 * the native ComponentOne Chart DLL function calls.
 *
 * If you wish to develop using the OCX, please *DO NOT* use these header
 * files. They are for use in C++ based, straight DLL programming only.
 * Wrapper files for the OCX will be generated automatically when you
 * insert the component into your development environment.
 *
 * To find more information about the "get" and "set" methods given below,
 * look up the associated property in the DLL manual. Properties are always
 * the second parameter in methods that wrap "GetValues" or "SetValues" DLL
 * calls. For example, to find out about GetFooterWidth(), look up the
 * XRT_FOOTER_WIDTH property.
 *
 * To find more information about the "operations" methods given below,
 * look up the wrapped DLL call directly. For example, to find out about
 * Pick(), look up the "XrtPick" function in the DLL manual.
 *
 * ------------------------------------
 * The main class, CChart3D, derives from CWnd. The majority of native
 * ComponentOne Chart DLL calls require a handle to the chart (usually shown in
 * the DLL manual as HXRT2D hChart) as the first parameter. CChart3D wraps
 * these calls and provides straightforward names for the access functions.
 *
 * As with a CWnd, you construct a CChart3D object in two steps. First, call
 * the constructor, which in turn constructs the associated CWnd object.
 * Then call the Create method, which creates and sets up the CWnd, then
 * creates the Chart and attaches the Windows child window to it.
 * See the MFC help for CWnd for the parameters to be passed to Create().
 * A caption is unnecessary. A typical example looks like this:
 *     m_chart.Create(, WS_CHILD|WS_VISIBLE, rect, this, 0);
 *
 * ------------------------------------
 * The CChart3DTextArea class can be created via a call to "new". Always
 * remember to free these objects via "delete" at the end of execution!
 *
 * CChart3DTextArea wraps ComponentOne Chart DLL calls that require a data handle
 * (usually shown in the DLL manual as XrtTextHandle text) as the first
 * parameter. CChart3DTextArea wraps these calls and provides more
 * straightforward names for the access functions.
 *
 * ------------------------------------
 * The CChart3DData class is normally created via a call to "new". Always
 * remember to free these objects via "delete" at the end of execution!
 *
 * CChart3DData objects can be constructed in one of four ways:
 *   (i)  You can allocate an empty data object of either Array or General
 *        data. In addition to the type, you must specify the number of
 *        sets (called Series by the OCX) and the number of points in each
 *        set. This constructor calls the DLL's XrtDataCreate() function.
 *  (ii)  You can allocate a data object and load it with data
 *        specified in a text file. This constructor calls the DLL's
 *        XrtDataCreateFromFile() function.
 *  (iii) You can allocate a data object and copy the data from another
 *        provided data object. This constructor calls the DLL's
 *        XrtDataCopy() function.
 *  (iv)  You can directly assign a data handle acquired from another
 *        source. This would allow multiple charts to share a common
 *        data object.
 *
 * CChart3DData wraps ComponentOne Chart DLL calls that require a data handle
 * (usually shown in the DLL manual as XrtDataHandle hData) as the first
 * parameter. CChart3DData wraps these calls and provides more
 * straightforward names for the access functions.
 *
 * ==========================================================================
 */

#include <olch3d.h>

/*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 *
 *  Class CChart3D
 *
 *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
*/

class CChart3D : public CWnd
{
    DECLARE_DYNAMIC(CChart3D)

public:
    // Constructors
    CChart3D() { m_hChart = NULL; };
    BOOL Create(LPCTSTR lpszCaption, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID);

    HXRT3D m_hChart;

#if (_MFC_VER < 0x0300)
	WNDPROC* GetSuperWndProcAddr();
#endif

    // Destructor
    virtual ~CChart3D() { DestroyWindow(); if (m_hChart) { Xrt3dDestroy(m_hChart); } }

    // Get methods
    HFONT GetAxisFont() { HFONT val; Xrt3dGetValues(m_hChart, XRT3D_AXIS_FONT, &val, NULL); return(val); }
    Xrt3dStrokeFont GetAxisStrokeFont() { Xrt3dStrokeFont val; Xrt3dGetValues(m_hChart, XRT3D_AXIS_STROKE_FONT, &val, NULL); return(val); }
    int GetAxisStrokeSize() { int val; Xrt3dGetValues(m_hChart, XRT3D_AXIS_STROKE_SIZE, &val, NULL); return(val); }
    HFONT GetAxisTitleFont() { HFONT val; Xrt3dGetValues(m_hChart, XRT3D_AXIS_TITLE_FONT, &val, NULL); return(val); }
    Xrt3dStrokeFont GetAxisTitleStrokeFont() { Xrt3dStrokeFont val; Xrt3dGetValues(m_hChart, XRT3D_AXIS_TITLE_STROKE_FONT, &val, NULL); return(val); }
    int GetAxisTitleStrokeSize() { int val; Xrt3dGetValues(m_hChart, XRT3D_AXIS_TITLE_STROKE_SIZE, &val, NULL); return(val); }
    COLORREF GetBackgroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_BACKGROUND_COLOR, &val, NULL); return(val); }
    Xrt3dBorder GetBorder() { Xrt3dBorder val; Xrt3dGetValues(m_hChart, XRT3D_BORDER, &val, NULL); return(val); }
    int GetBorderWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_BORDER_WIDTH, &val, NULL); return(val); }
    Xrt3dContourStyle ** GetContourStyles() { Xrt3dContourStyle ** val; Xrt3dGetValues(m_hChart, XRT3D_CONTOUR_STYLES, &val, NULL); return(val); }
    COLORREF GetDataAreaBackgroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_DATA_AREA_BACKGROUND_COLOR, &val, NULL); return(val); }
    Xrt3dDataStyle ** GetDataStyles() { Xrt3dDataStyle ** val; Xrt3dGetValues(m_hChart, XRT3D_DATA_STYLES, &val, NULL); return(val); }
    BOOL GetDataStylesUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DATA_STYLES_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetDebug() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DEBUG, &val, NULL); return(val); }
    Xrt3dDistnMethod GetDistnMethod() { Xrt3dDistnMethod val; Xrt3dGetValues(m_hChart, XRT3D_DISTN_METHOD, &val, NULL); return(val); }
    Xrt3dDistnTable * GetDistnTable() { Xrt3dDistnTable * val; Xrt3dGetValues(m_hChart, XRT3D_DISTN_TABLE, &val, NULL); return(val); }
    BOOL GetDoubleBuffer() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DOUBLE_BUFFER, &val, NULL); return(val); }
    BOOL GetDrawContours() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DRAW_CONTOURS, &val, NULL); return(val); }
    BOOL GetDrawDropLines() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DRAW_DROP_LINES, &val, NULL); return(val); }
    BOOL GetDrawHiddenLines() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DRAW_HIDDEN_LINES, &val, NULL); return(val); }
    BOOL GetDrawMesh() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DRAW_MESH, &val, NULL); return(val); }
    BOOL GetDrawShaded() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DRAW_SHADED, &val, NULL); return(val); }
    BOOL GetDrawZones() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_DRAW_ZONES, &val, NULL); return(val); }
    Xrt3dAdjust GetFooterAdjust() { Xrt3dAdjust val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_ADJUST, &val, NULL); return(val); }
    COLORREF GetFooterBackgroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_BACKGROUND_COLOR, &val, NULL); return(val); }
    Xrt3dBorder GetFooterBorder() { Xrt3dBorder val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_BORDER, &val, NULL); return(val); }
    int GetFooterBorderWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_BORDER_WIDTH, &val, NULL); return(val); }
    HFONT GetFooterFont() { HFONT val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_FONT, &val, NULL); return(val); }
    COLORREF GetFooterForegroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetFooterHeight() { int val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_HEIGHT, &val, NULL); return(val); }
    Xrt3dImage GetFooterImage() { Xrt3dImage val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_IMAGE, &val, NULL); return(val); }
    Xrt3dImageLayout GetFooterImageLayout() { Xrt3dImageLayout val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetFooterImageTransparent() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    BOOL GetFooterImageMinimumSize() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_IMAGE_MINIMUM_SIZE, &val, NULL); return(val); }
    TCHAR * GetFooterImageObsolete() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_IMAGE_OBSOLETE, &val, NULL); return(val); }
    TCHAR ** GetFooterStrings() { TCHAR ** val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_STRINGS, &val, NULL); return(val); }
    int GetFooterWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_WIDTH, &val, NULL); return(val); }
    int GetFooterX() { int val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_X, &val, NULL); return(val); }
    BOOL GetFooterXUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_X_USE_DEFAULT, &val, NULL); return(val); }
    int GetFooterY() { int val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_Y, &val, NULL); return(val); }
    BOOL GetFooterYUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_FOOTER_Y_USE_DEFAULT, &val, NULL); return(val); }
    COLORREF GetForegroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_FOREGROUND_COLOR, &val, NULL); return(val); }
    COLORREF GetGraphBackgroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_BACKGROUND_COLOR, &val, NULL); return(val); }
    Xrt3dBorder GetGraphBorder() { Xrt3dBorder val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_BORDER, &val, NULL); return(val); }
    int GetGraphBorderWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_BORDER_WIDTH, &val, NULL); return(val); }
    COLORREF GetGraphForegroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetGraphHeight() { int val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_HEIGHT, &val, NULL); return(val); }
    BOOL GetGraphHeightUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_HEIGHT_USE_DEFAULT, &val, NULL); return(val); }
    Xrt3dImage GetGraphImage() { Xrt3dImage val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_IMAGE, &val, NULL); return(val); }
    Xrt3dImageLayout GetGraphImageLayout() { Xrt3dImageLayout val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetGraphImageTransparent() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    TCHAR * GetGraphImageObsolete() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_IMAGE_OBSOLETE, &val, NULL); return(val); }
    int GetGraphWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_WIDTH, &val, NULL); return(val); }
    BOOL GetGraphWidthUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_WIDTH_USE_DEFAULT, &val, NULL); return(val); }
    int GetGraphX() { int val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_X, &val, NULL); return(val); }
    BOOL GetGraphXUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_X_USE_DEFAULT, &val, NULL); return(val); }
    int GetGraphY() { int val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_Y, &val, NULL); return(val); }
    BOOL GetGraphYUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_GRAPH_Y_USE_DEFAULT, &val, NULL); return(val); }
    Xrt3dAdjust GetHeaderAdjust() { Xrt3dAdjust val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_ADJUST, &val, NULL); return(val); }
    COLORREF GetHeaderBackgroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_BACKGROUND_COLOR, &val, NULL); return(val); }
    Xrt3dBorder GetHeaderBorder() { Xrt3dBorder val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_BORDER, &val, NULL); return(val); }
    int GetHeaderBorderWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_BORDER_WIDTH, &val, NULL); return(val); }
    HFONT GetHeaderFont() { HFONT val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_FONT, &val, NULL); return(val); }
    COLORREF GetHeaderForegroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetHeaderHeight() { int val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_HEIGHT, &val, NULL); return(val); }
    Xrt3dImage GetHeaderImage() { Xrt3dImage val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_IMAGE, &val, NULL); return(val); }
    Xrt3dImageLayout GetHeaderImageLayout() { Xrt3dImageLayout val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetHeaderImageTransparent() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    BOOL GetHeaderImageMinimumSize() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_IMAGE_MINIMUM_SIZE, &val, NULL); return(val); }
    TCHAR * GetHeaderImageObsolete() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_IMAGE_OBSOLETE, &val, NULL); return(val); }
    TCHAR ** GetHeaderStrings() { TCHAR ** val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_STRINGS, &val, NULL); return(val); }
    int GetHeaderWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_WIDTH, &val, NULL); return(val); }
    int GetHeaderX() { int val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_X, &val, NULL); return(val); }
    BOOL GetHeaderXUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_X_USE_DEFAULT, &val, NULL); return(val); }
    int GetHeaderY() { int val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_Y, &val, NULL); return(val); }
    BOOL GetHeaderYUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_HEADER_Y_USE_DEFAULT, &val, NULL); return(val); }
    int GetHeight() { int val; Xrt3dGetValues(m_hChart, XRT3D_HEIGHT, &val, NULL); return(val); }
    Xrt3dImage GetImage() { Xrt3dImage val; Xrt3dGetValues(m_hChart, XRT3D_IMAGE, &val, NULL); return(val); }
    Xrt3dImageLayout GetImageLayout() { Xrt3dImageLayout val; Xrt3dGetValues(m_hChart, XRT3D_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetImageTransparent() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    TCHAR * GetImageObsolete() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_IMAGE_OBSOLETE, &val, NULL); return(val); }
    Xrt3dAnchor GetLegendAnchor() { Xrt3dAnchor val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_ANCHOR, &val, NULL); return(val); }
    COLORREF GetLegendBackgroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_BACKGROUND_COLOR, &val, NULL); return(val); }
    Xrt3dBorder GetLegendBorder() { Xrt3dBorder val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_BORDER, &val, NULL); return(val); }
    int GetLegendBorderWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_BORDER_WIDTH, &val, NULL); return(val); }
    Xrt3dDistnRange GetLegendDistnRange() { Xrt3dDistnRange val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_DISTN_RANGE, &val, NULL); return(val); }
    HFONT GetLegendFont() { HFONT val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_FONT, &val, NULL); return(val); }
    COLORREF GetLegendForegroundColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetLegendHeight() { int val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_HEIGHT, &val, NULL); return(val); }
    Xrt3dImage GetLegendImage() { Xrt3dImage val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_IMAGE, &val, NULL); return(val); }
    Xrt3dImageLayout GetLegendImageLayout() { Xrt3dImageLayout val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetLegendImageTransparent() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    TCHAR * GetLegendImageObsolete() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_IMAGE_OBSOLETE, &val, NULL); return(val); }
    int GetLegendLabelFilter() { int val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_LABEL_FILTER, &val, NULL); return(val); }
    TCHAR* GetLegendLabelFunc() { TCHAR* val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_LABEL_FUNC, &val, NULL); return(val); }
    Xrt3dAlign GetLegendOrientation() { Xrt3dAlign val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_ORIENTATION, &val, NULL); return(val); }
    BOOL GetLegendShow() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_SHOW, &val, NULL); return(val); }
    TCHAR ** GetLegendStrings() { TCHAR ** val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_STRINGS, &val, NULL); return(val); }
    Xrt3dLegendStyle GetLegendStyle() { Xrt3dLegendStyle val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_STYLE, &val, NULL); return(val); }
    TCHAR * GetLegendTitle() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_TITLE, &val, NULL); return(val); }
    int GetLegendWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_WIDTH, &val, NULL); return(val); }
    int GetLegendX() { int val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_X, &val, NULL); return(val); }
    BOOL GetLegendXUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_X_USE_DEFAULT, &val, NULL); return(val); }
    int GetLegendY() { int val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_Y, &val, NULL); return(val); }
    BOOL GetLegendYUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_LEGEND_Y_USE_DEFAULT, &val, NULL); return(val); }
    COLORREF GetMeshBottomColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_MESH_BOTTOM_COLOR, &val, NULL); return(val); }
    COLORREF GetMeshTopColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_MESH_TOP_COLOR, &val, NULL); return(val); }
    TCHAR * GetName() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_NAME, &val, NULL); return(val); }
    int GetNumDistnLevels() { int val; Xrt3dGetValues(m_hChart, XRT3D_NUM_DISTN_LEVELS, &val, NULL); return(val); }
    double GetPerspectiveDepth() { double val; Xrt3dGetValues(m_hChart, XRT3D_PERSPECTIVE_DEPTH, &val, NULL); return(val); }
    Xrt3dPreviewMethod GetPreviewMethod() { Xrt3dPreviewMethod val; Xrt3dGetValues(m_hChart, XRT3D_PREVIEW_METHOD, &val, NULL); return(val); }
    int GetProjectZMax() { int val; Xrt3dGetValues(m_hChart, XRT3D_PROJECT_ZMAX, &val, NULL); return(val); }
    int GetProjectZMin() { int val; Xrt3dGetValues(m_hChart, XRT3D_PROJECT_ZMIN, &val, NULL); return(val); }
    BOOL GetRepaint() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_REPAINT, &val, NULL); return(val); }
    BOOL GetSolidSurface() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_SOLID_SURFACE, &val, NULL); return(val); }
    COLORREF GetSurfaceBottomColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_SURFACE_BOTTOM_COLOR, &val, NULL); return(val); }
    Xrt3dData * GetSurfaceData() { Xrt3dData * val; Xrt3dGetValues(m_hChart, XRT3D_SURFACE_DATA, &val, NULL); return(val); }
    COLORREF GetSurfaceTopColor() { COLORREF val; Xrt3dGetValues(m_hChart, XRT3D_SURFACE_TOP_COLOR, &val, NULL); return(val); }
    Xrt3dType GetType() { Xrt3dType val; Xrt3dGetValues(m_hChart, XRT3D_TYPE, &val, NULL); return(val); }
    BOOL GetUseTruetype() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_USE_TRUETYPE, &val, NULL); return(val); }
    BOOL GetViewNormalized() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_VIEW_NORMALIZED, &val, NULL); return(val); }
    double GetViewScale() { double val; Xrt3dGetValues(m_hChart, XRT3D_VIEW_SCALE, &val, NULL); return(val); }
    double GetViewXTranslate() { double val; Xrt3dGetValues(m_hChart, XRT3D_VIEW_XTRANSLATE, &val, NULL); return(val); }
    double GetViewYTranslate() { double val; Xrt3dGetValues(m_hChart, XRT3D_VIEW_YTRANSLATE, &val, NULL); return(val); }
    int GetWidth() { int val; Xrt3dGetValues(m_hChart, XRT3D_WIDTH, &val, NULL); return(val); }
    Xrt3dAnnoMethod GetXAnnoMethod() { Xrt3dAnnoMethod val; Xrt3dGetValues(m_hChart, XRT3D_XANNO_METHOD, &val, NULL); return(val); }
    BOOL GetXAxisShow() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_XAXIS_SHOW, &val, NULL); return(val); }
    TCHAR * GetXAxisTitle() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_XAXIS_TITLE, &val, NULL); return(val); }
    Xrt3dBarFormat GetXBarFormat() { Xrt3dBarFormat val; Xrt3dGetValues(m_hChart, XRT3D_XBAR_FORMAT, &val, NULL); return(val); }
    double GetXBarSpacing() { double val; Xrt3dGetValues(m_hChart, XRT3D_XBAR_SPACING, &val, NULL); return(val); }
    TCHAR ** GetXDataLabels() { TCHAR ** val; Xrt3dGetValues(m_hChart, XRT3D_XDATA_LABELS, &val, NULL); return(val); }
    Xrt3dLineStyle * GetXGridLineStyle() { Xrt3dLineStyle * val; Xrt3dGetValues(m_hChart, XRT3D_XGRID_LINE_STYLE, &val, NULL); return(val); }
    BOOL GetXGridLineStyleUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_XGRID_LINE_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    int GetXGridLines() { int val; Xrt3dGetValues(m_hChart, XRT3D_XGRID_LINES, &val, NULL); return(val); }
    double GetXMax() { double val; Xrt3dGetValues(m_hChart, XRT3D_XMAX, &val, NULL); return(val); }
    BOOL GetXMaxUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_XMAX_USE_DEFAULT, &val, NULL); return(val); }
    int GetXMeshFilter() { int val; Xrt3dGetValues(m_hChart, XRT3D_XMESH_FILTER, &val, NULL); return(val); }
    BOOL GetXMeshShow() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_XMESH_SHOW, &val, NULL); return(val); }
    double GetXMin() { double val; Xrt3dGetValues(m_hChart, XRT3D_XMIN, &val, NULL); return(val); }
    BOOL GetXMinUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_XMIN_USE_DEFAULT, &val, NULL); return(val); }
    double GetXRotation() { double val; Xrt3dGetValues(m_hChart, XRT3D_XROTATION, &val, NULL); return(val); }
    double GetXScale() { double val; Xrt3dGetValues(m_hChart, XRT3D_XSCALE, &val, NULL); return(val); }
    Xrt3dValueLabel ** GetXValueLabels() { Xrt3dValueLabel ** val; Xrt3dGetValues(m_hChart, XRT3D_XVALUE_LABELS, &val, NULL); return(val); }
    Xrt3dXYColor ** GetXYColors() { Xrt3dXYColor ** val; Xrt3dGetValues(m_hChart, XRT3D_XY_COLORS, &val, NULL); return(val); }
    Xrt3dAnnoMethod GetYAnnoMethod() { Xrt3dAnnoMethod val; Xrt3dGetValues(m_hChart, XRT3D_YANNO_METHOD, &val, NULL); return(val); }
    BOOL GetYAxisShow() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_YAXIS_SHOW, &val, NULL); return(val); }
    TCHAR * GetYAxisTitle() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_YAXIS_TITLE, &val, NULL); return(val); }
    Xrt3dBarFormat GetYBarFormat() { Xrt3dBarFormat val; Xrt3dGetValues(m_hChart, XRT3D_YBAR_FORMAT, &val, NULL); return(val); }
    double GetYBarSpacing() { double val; Xrt3dGetValues(m_hChart, XRT3D_YBAR_SPACING, &val, NULL); return(val); }
    TCHAR ** GetYDataLabels() { TCHAR ** val; Xrt3dGetValues(m_hChart, XRT3D_YDATA_LABELS, &val, NULL); return(val); }
    Xrt3dLineStyle * GetYGridLineStyle() { Xrt3dLineStyle * val; Xrt3dGetValues(m_hChart, XRT3D_YGRID_LINE_STYLE, &val, NULL); return(val); }
    BOOL GetYGridLineStyleUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_YGRID_LINE_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    int GetYGridLines() { int val; Xrt3dGetValues(m_hChart, XRT3D_YGRID_LINES, &val, NULL); return(val); }
    double GetYMax() { double val; Xrt3dGetValues(m_hChart, XRT3D_YMAX, &val, NULL); return(val); }
    BOOL GetYMaxUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_YMAX_USE_DEFAULT, &val, NULL); return(val); }
    int GetYMeshFilter() { int val; Xrt3dGetValues(m_hChart, XRT3D_YMESH_FILTER, &val, NULL); return(val); }
    BOOL GetYMeshShow() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_YMESH_SHOW, &val, NULL); return(val); }
    double GetYMin() { double val; Xrt3dGetValues(m_hChart, XRT3D_YMIN, &val, NULL); return(val); }
    BOOL GetYMinUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_YMIN_USE_DEFAULT, &val, NULL); return(val); }
    double GetYRotation() { double val; Xrt3dGetValues(m_hChart, XRT3D_YROTATION, &val, NULL); return(val); }
    double GetYScale() { double val; Xrt3dGetValues(m_hChart, XRT3D_YSCALE, &val, NULL); return(val); }
    Xrt3dValueLabel ** GetYValueLabels() { Xrt3dValueLabel ** val; Xrt3dGetValues(m_hChart, XRT3D_YVALUE_LABELS, &val, NULL); return(val); }
    Xrt3dAnnoMethod GetZAnnoMethod() { Xrt3dAnnoMethod val; Xrt3dGetValues(m_hChart, XRT3D_ZANNO_METHOD, &val, NULL); return(val); }
    BOOL GetZAxisShow() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_ZAXIS_SHOW, &val, NULL); return(val); }
    TCHAR * GetZAxisTitle() { TCHAR * val; Xrt3dGetValues(m_hChart, XRT3D_ZAXIS_TITLE, &val, NULL); return(val); }
    Xrt3dLineStyle * GetZGridLineStyle() { Xrt3dLineStyle * val; Xrt3dGetValues(m_hChart, XRT3D_ZGRID_LINE_STYLE, &val, NULL); return(val); }
    BOOL GetZGridLineStyleUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_ZGRID_LINE_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    int GetZGridLines() { int val; Xrt3dGetValues(m_hChart, XRT3D_ZGRID_LINES, &val, NULL); return(val); }
    double GetZMax() { double val; Xrt3dGetValues(m_hChart, XRT3D_ZMAX, &val, NULL); return(val); }
    BOOL GetZMaxUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_ZMAX_USE_DEFAULT, &val, NULL); return(val); }
    double GetZMin() { double val; Xrt3dGetValues(m_hChart, XRT3D_ZMIN, &val, NULL); return(val); }
    BOOL GetZMinUseDefault() { BOOL val; Xrt3dGetValues(m_hChart, XRT3D_ZMIN_USE_DEFAULT, &val, NULL); return(val); }
    double GetZOrigin() { double val; Xrt3dGetValues(m_hChart, XRT3D_ZORIGIN, &val, NULL); return(val); }
    double GetZRotation() { double val; Xrt3dGetValues(m_hChart, XRT3D_ZROTATION, &val, NULL); return(val); }
    double GetZScale() { double val; Xrt3dGetValues(m_hChart, XRT3D_ZSCALE, &val, NULL); return(val); }
    Xrt3dValueLabel ** GetZValueLabels() { Xrt3dValueLabel ** val; Xrt3dGetValues(m_hChart, XRT3D_ZVALUE_LABELS, &val, NULL); return(val); }
    Xrt3dData * GetZoneData() { Xrt3dData * val; Xrt3dGetValues(m_hChart, XRT3D_ZONE_DATA, &val, NULL); return(val); }
    Xrt3dZoneMethod GetZoneMethod() { Xrt3dZoneMethod val; Xrt3dGetValues(m_hChart, XRT3D_ZONE_METHOD, &val, NULL); return(val); }
	Xrt3dFontRotation GetFontRotation() { Xrt3dFontRotation val; Xrt3dGetValues(m_hChart, XRT3DT_FONT_ROTATION, &val, NULL); return(val); }

    // Set methods
    void SetAxisFont(HFONT val) { Xrt3dSetValues(m_hChart, XRT3D_AXIS_FONT, (int)val, NULL); }
    void SetAxisStrokeFont(Xrt3dStrokeFont val) { Xrt3dSetValues(m_hChart, XRT3D_AXIS_STROKE_FONT, val, NULL); }
    void SetAxisStrokeSize(int val) { Xrt3dSetValues(m_hChart, XRT3D_AXIS_STROKE_SIZE, val, NULL); }
    void SetAxisTitleFont(HFONT val) { Xrt3dSetValues(m_hChart, XRT3D_AXIS_TITLE_FONT, (int)val, NULL); }
    void SetAxisTitleStrokeFont(Xrt3dStrokeFont val) { Xrt3dSetValues(m_hChart, XRT3D_AXIS_TITLE_STROKE_FONT, val, NULL); }
    void SetAxisTitleStrokeSize(int val) { Xrt3dSetValues(m_hChart, XRT3D_AXIS_TITLE_STROKE_SIZE, val, NULL); }
    void SetBackgroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_BACKGROUND_COLOR, val, NULL); }
    void SetBorder(Xrt3dBorder val) { Xrt3dSetValues(m_hChart, XRT3D_BORDER, val, NULL); }
    void SetBorderWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_BORDER_WIDTH, val, NULL); }
    void SetContourStyles(Xrt3dContourStyle ** val) { Xrt3dSetValues(m_hChart, XRT3D_CONTOUR_STYLES, val, NULL); }
    void SetDataAreaBackgroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_DATA_AREA_BACKGROUND_COLOR, val, NULL); }
    void SetDataStyles(Xrt3dDataStyle ** val) { Xrt3dSetValues(m_hChart, XRT3D_DATA_STYLES, val, NULL); }
    void SetDataStylesUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DATA_STYLES_USE_DEFAULT, val, NULL); }
    void SetDebug(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DEBUG, val, NULL); }
    void SetDistnMethod(Xrt3dDistnMethod val) { Xrt3dSetValues(m_hChart, XRT3D_DISTN_METHOD, val, NULL); }
    void SetDistnTable(Xrt3dDistnTable * val) { Xrt3dSetValues(m_hChart, XRT3D_DISTN_TABLE, val, NULL); }
    void SetDoubleBuffer(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DOUBLE_BUFFER, val, NULL); }
    void SetDrawContours(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DRAW_CONTOURS, val, NULL); }
    void SetDrawDropLines(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DRAW_DROP_LINES, val, NULL); }
    void SetDrawHiddenLines(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DRAW_HIDDEN_LINES, val, NULL); }
    void SetDrawMesh(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DRAW_MESH, val, NULL); }
    void SetDrawShaded(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DRAW_SHADED, val, NULL); }
    void SetDrawZones(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_DRAW_ZONES, val, NULL); }
    void SetFooterAdjust(Xrt3dAdjust val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_ADJUST, val, NULL); }
    void SetFooterBackgroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_BACKGROUND_COLOR, val, NULL); }
    void SetFooterBorder(Xrt3dBorder val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_BORDER, val, NULL); }
    void SetFooterBorderWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_BORDER_WIDTH, val, NULL); }
    void SetFooterFont(HFONT val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_FONT, (int)val, NULL); }
    void SetFooterForegroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_FOREGROUND_COLOR, val, NULL); }
    void SetFooterHeight(int val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_HEIGHT, val, NULL); }
    void SetFooterImage(Xrt3dImage val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_IMAGE, val, NULL); }
    void SetFooterImageLayout(Xrt3dImageLayout val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_IMAGE_LAYOUT, val, NULL); }
    void SetFooterImageMinimumSize(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_IMAGE_MINIMUM_SIZE, val, NULL); }
    void SetFooterImageObsolete(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_IMAGE_OBSOLETE, val, NULL); }
    void SetFooterImageTransparent(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_IMAGE_TRANSPARENT, val, NULL); }
    void SetFooterStrings(TCHAR ** val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_STRINGS, val, NULL); }
    void SetFooterWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_WIDTH, val, NULL); }
    void SetFooterX(int val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_X, val, NULL); }
    void SetFooterXUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_X_USE_DEFAULT, val, NULL); }
    void SetFooterY(int val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_Y, val, NULL); }
    void SetFooterYUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_FOOTER_Y_USE_DEFAULT, val, NULL); }
    void SetForegroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_FOREGROUND_COLOR, val, NULL); }
    void SetGraphBackgroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_BACKGROUND_COLOR, val, NULL); }
    void SetGraphBorder(Xrt3dBorder val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_BORDER, val, NULL); }
    void SetGraphBorderWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_BORDER_WIDTH, val, NULL); }
    void SetGraphForegroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_FOREGROUND_COLOR, val, NULL); }
    void SetGraphHeight(int val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_HEIGHT, val, NULL); }
    void SetGraphHeightUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_HEIGHT_USE_DEFAULT, val, NULL); }
    void SetGraphImage(Xrt3dImage val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_IMAGE, val, NULL); }
    void SetGraphImageLayout(Xrt3dImageLayout val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_IMAGE_LAYOUT, val, NULL); }
    void SetGraphImageObsolete(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_IMAGE_OBSOLETE, val, NULL); }
    void SetGraphImageTransparent(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_IMAGE_TRANSPARENT, val, NULL); }
    void SetGraphWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_WIDTH, val, NULL); }
    void SetGraphWidthUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_WIDTH_USE_DEFAULT, val, NULL); }
    void SetGraphX(int val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_X, val, NULL); }
    void SetGraphXUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_X_USE_DEFAULT, val, NULL); }
    void SetGraphY(int val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_Y, val, NULL); }
    void SetGraphYUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_GRAPH_Y_USE_DEFAULT, val, NULL); }
    void SetHeaderAdjust(Xrt3dAdjust val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_ADJUST, val, NULL); }
    void SetHeaderBackgroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_BACKGROUND_COLOR, val, NULL); }
    void SetHeaderBorder(Xrt3dBorder val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_BORDER, val, NULL); }
    void SetHeaderBorderWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_BORDER_WIDTH, val, NULL); }
    void SetHeaderFont(HFONT val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_FONT, (int)val, NULL); }
    void SetHeaderForegroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_FOREGROUND_COLOR, val, NULL); }
    void SetHeaderHeight(int val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_HEIGHT, val, NULL); }
    void SetHeaderImage(Xrt3dImage val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_IMAGE, val, NULL); }
    void SetHeaderImageLayout(Xrt3dImageLayout val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_IMAGE_LAYOUT, val, NULL); }
    void SetHeaderImageMinimumSize(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_IMAGE_MINIMUM_SIZE, val, NULL); }
    void SetHeaderImageObsolete(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_IMAGE_OBSOLETE, val, NULL); }
    void SetHeaderImageTransparent(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_IMAGE_TRANSPARENT, val, NULL); }
    void SetHeaderStrings(TCHAR ** val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_STRINGS, val, NULL); }
    void SetHeaderWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_WIDTH, val, NULL); }
    void SetHeaderX(int val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_X, val, NULL); }
    void SetHeaderXUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_X_USE_DEFAULT, val, NULL); }
    void SetHeaderY(int val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_Y, val, NULL); }
    void SetHeaderYUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_HEADER_Y_USE_DEFAULT, val, NULL); }
    void SetHeight(int val) { Xrt3dSetValues(m_hChart, XRT3D_HEIGHT, val, NULL); }
    void SetImage(Xrt3dImage val) { Xrt3dSetValues(m_hChart, XRT3D_IMAGE, val, NULL); }
    void SetImageLayout(Xrt3dImageLayout val) { Xrt3dSetValues(m_hChart, XRT3D_IMAGE_LAYOUT, val, NULL); }
    void SetImageTransparent(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_IMAGE_TRANSPARENT, val, NULL); }
    void SetImageObsolete(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_IMAGE_OBSOLETE, val, NULL); }
    void SetLegendAnchor(Xrt3dAnchor val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_ANCHOR, val, NULL); }
    void SetLegendBackgroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_BACKGROUND_COLOR, val, NULL); }
    void SetLegendBorder(Xrt3dBorder val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_BORDER, val, NULL); }
    void SetLegendBorderWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_BORDER_WIDTH, val, NULL); }
    void SetLegendDistnRange(Xrt3dDistnRange val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_DISTN_RANGE, val, NULL); }
    void SetLegendFont(HFONT val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_FONT, (int)val, NULL); }
    void SetLegendForegroundColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_FOREGROUND_COLOR, val, NULL); }
    void SetLegendHeight(int val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_HEIGHT, val, NULL); }
    void SetLegendImage(Xrt3dImage val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_IMAGE, val, NULL); }
    void SetLegendImageLayout(Xrt3dImageLayout val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_IMAGE_LAYOUT, val, NULL); }
    void SetLegendImageObsolete(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_IMAGE_OBSOLETE, val, NULL); }
    void SetLegendImageTransparent(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_IMAGE_TRANSPARENT, val, NULL); }
    void SetLegendLabelFilter(int val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_LABEL_FILTER, val, NULL); }
    void SetLegendLabelFunc(TCHAR* val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_LABEL_FUNC, val, NULL); }
    void SetLegendOrientation(Xrt3dAlign val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_ORIENTATION, val, NULL); }
    void SetLegendShow(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_SHOW, val, NULL); }
    void SetLegendStrings(TCHAR ** val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_STRINGS, val, NULL); }
    void SetLegendStyle(Xrt3dLegendStyle val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_STYLE, val, NULL); }
    void SetLegendTitle(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_TITLE, val, NULL); }
    void SetLegendWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_WIDTH, val, NULL); }
    void SetLegendX(int val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_X, val, NULL); }
    void SetLegendXUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_X_USE_DEFAULT, val, NULL); }
    void SetLegendY(int val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_Y, val, NULL); }
    void SetLegendYUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_LEGEND_Y_USE_DEFAULT, val, NULL); }
    void SetMeshBottomColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_MESH_BOTTOM_COLOR, val, NULL); }
    void SetMeshTopColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_MESH_TOP_COLOR, val, NULL); }
    void SetName(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_NAME, val, NULL); }
    void SetNumDistnLevels(int val) { Xrt3dSetValues(m_hChart, XRT3D_NUM_DISTN_LEVELS, val, NULL); }
    void SetPerspectiveDepth(double val) { Xrt3dSetValues(m_hChart, XRT3D_PERSPECTIVE_DEPTH, val, NULL); }
    void SetPreviewMethod(Xrt3dPreviewMethod val) { Xrt3dSetValues(m_hChart, XRT3D_PREVIEW_METHOD, val, NULL); }
    void SetProjectZMax(int val) { Xrt3dSetValues(m_hChart, XRT3D_PROJECT_ZMAX, val, NULL); }
    void SetProjectZMin(int val) { Xrt3dSetValues(m_hChart, XRT3D_PROJECT_ZMIN, val, NULL); }
    void SetRepaint(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_REPAINT, val, NULL); }
    void SetSolidSurface(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_SOLID_SURFACE, val, NULL); }
    void SetSurfaceBottomColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_SURFACE_BOTTOM_COLOR, val, NULL); }
    void SetSurfaceData(Xrt3dData * val) { Xrt3dSetValues(m_hChart, XRT3D_SURFACE_DATA, val, NULL); }
    void SetSurfaceTopColor(COLORREF val) { Xrt3dSetValues(m_hChart, XRT3D_SURFACE_TOP_COLOR, val, NULL); }
    void SetType(Xrt3dType val) { Xrt3dSetValues(m_hChart, XRT3D_TYPE, val, NULL); }
    void SetUseTruetype(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_USE_TRUETYPE, val, NULL); }
    void SetViewNormalized(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_VIEW_NORMALIZED, val, NULL); }
    void SetViewScale(double val) { Xrt3dSetValues(m_hChart, XRT3D_VIEW_SCALE, val, NULL); }
    void SetViewXTranslate(double val) { Xrt3dSetValues(m_hChart, XRT3D_VIEW_XTRANSLATE, val, NULL); }
    void SetViewYTranslate(double val) { Xrt3dSetValues(m_hChart, XRT3D_VIEW_YTRANSLATE, val, NULL); }
    void SetWidth(int val) { Xrt3dSetValues(m_hChart, XRT3D_WIDTH, val, NULL); }
    void SetXAnnoMethod(Xrt3dAnnoMethod val) { Xrt3dSetValues(m_hChart, XRT3D_XANNO_METHOD, val, NULL); }
    void SetXAxisShow(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_XAXIS_SHOW, val, NULL); }
    void SetXAxisTitle(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_XAXIS_TITLE, val, NULL); }
    void SetXBarFormat(Xrt3dBarFormat val) { Xrt3dSetValues(m_hChart, XRT3D_XBAR_FORMAT, val, NULL); }
    void SetXBarSpacing(double val) { Xrt3dSetValues(m_hChart, XRT3D_XBAR_SPACING, val, NULL); }
    void SetXDataLabels(TCHAR ** val) { Xrt3dSetValues(m_hChart, XRT3D_XDATA_LABELS, val, NULL); }
    void SetXGridLineStyle(Xrt3dLineStyle * val) { Xrt3dSetValues(m_hChart, XRT3D_XGRID_LINE_STYLE, val, NULL); }
    void SetXGridLineStyleUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_XGRID_LINE_STYLE_USE_DEFAULT, val, NULL); }
    void SetXGridLines(int val) { Xrt3dSetValues(m_hChart, XRT3D_XGRID_LINES, val, NULL); }
    void SetXMax(double val) { Xrt3dSetValues(m_hChart, XRT3D_XMAX, val, NULL); }
    void SetXMaxUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_XMAX_USE_DEFAULT, val, NULL); }
    void SetXMeshFilter(int val) { Xrt3dSetValues(m_hChart, XRT3D_XMESH_FILTER, val, NULL); }
    void SetXMeshShow(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_XMESH_SHOW, val, NULL); }
    void SetXMin(double val) { Xrt3dSetValues(m_hChart, XRT3D_XMIN, val, NULL); }
    void SetXMinUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_XMIN_USE_DEFAULT, val, NULL); }
    void SetXRotation(double val) { Xrt3dSetValues(m_hChart, XRT3D_XROTATION, val, NULL); }
    void SetXScale(double val) { Xrt3dSetValues(m_hChart, XRT3D_XSCALE, val, NULL); }
    void SetXValueLabels(Xrt3dValueLabel ** val) { Xrt3dSetValues(m_hChart, XRT3D_XVALUE_LABELS, val, NULL); }
    void SetXYColors(Xrt3dXYColor ** val) { Xrt3dSetValues(m_hChart, XRT3D_XY_COLORS, val, NULL); }
    void SetYAnnoMethod(Xrt3dAnnoMethod val) { Xrt3dSetValues(m_hChart, XRT3D_YANNO_METHOD, val, NULL); }
    void SetYAxisShow(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_YAXIS_SHOW, val, NULL); }
    void SetYAxisTitle(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_YAXIS_TITLE, val, NULL); }
    void SetYBarFormat(Xrt3dBarFormat val) { Xrt3dSetValues(m_hChart, XRT3D_YBAR_FORMAT, val, NULL); }
    void SetYBarSpacing(double val) { Xrt3dSetValues(m_hChart, XRT3D_YBAR_SPACING, val, NULL); }
    void SetYDataLabels(TCHAR ** val) { Xrt3dSetValues(m_hChart, XRT3D_YDATA_LABELS, val, NULL); }
    void SetYGridLineStyle(Xrt3dLineStyle * val) { Xrt3dSetValues(m_hChart, XRT3D_YGRID_LINE_STYLE, val, NULL); }
    void SetYGridLineStyleUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_YGRID_LINE_STYLE_USE_DEFAULT, val, NULL); }
    void SetYGridLines(int val) { Xrt3dSetValues(m_hChart, XRT3D_YGRID_LINES, val, NULL); }
    void SetYMax(double val) { Xrt3dSetValues(m_hChart, XRT3D_YMAX, val, NULL); }
    void SetYMaxUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_YMAX_USE_DEFAULT, val, NULL); }
    void SetYMeshFilter(int val) { Xrt3dSetValues(m_hChart, XRT3D_YMESH_FILTER, val, NULL); }
    void SetYMeshShow(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_YMESH_SHOW, val, NULL); }
    void SetYMin(double val) { Xrt3dSetValues(m_hChart, XRT3D_YMIN, val, NULL); }
    void SetYMinUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_YMIN_USE_DEFAULT, val, NULL); }
    void SetYRotation(double val) { Xrt3dSetValues(m_hChart, XRT3D_YROTATION, val, NULL); }
    void SetYScale(double val) { Xrt3dSetValues(m_hChart, XRT3D_YSCALE, val, NULL); }
    void SetYValueLabels(Xrt3dValueLabel ** val) { Xrt3dSetValues(m_hChart, XRT3D_YVALUE_LABELS, val, NULL); }
    void SetZAnnoMethod(Xrt3dAnnoMethod val) { Xrt3dSetValues(m_hChart, XRT3D_ZANNO_METHOD, val, NULL); }
    void SetZAxisShow(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_ZAXIS_SHOW, val, NULL); }
    void SetZAxisTitle(TCHAR * val) { Xrt3dSetValues(m_hChart, XRT3D_ZAXIS_TITLE, val, NULL); }
    void SetZGridLineStyle(Xrt3dLineStyle * val) { Xrt3dSetValues(m_hChart, XRT3D_ZGRID_LINE_STYLE, val, NULL); }
    void SetZGridLineStyleUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_ZGRID_LINE_STYLE_USE_DEFAULT, val, NULL); }
    void SetZGridLines(int val) { Xrt3dSetValues(m_hChart, XRT3D_ZGRID_LINES, val, NULL); }
    void SetZMax(double val) { Xrt3dSetValues(m_hChart, XRT3D_ZMAX, val, NULL); }
    void SetZMaxUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_ZMAX_USE_DEFAULT, val, NULL); }
    void SetZMin(double val) { Xrt3dSetValues(m_hChart, XRT3D_ZMIN, val, NULL); }
    void SetZMinUseDefault(BOOL val) { Xrt3dSetValues(m_hChart, XRT3D_ZMIN_USE_DEFAULT, val, NULL); }
    void SetZOrigin(double val) { Xrt3dSetValues(m_hChart, XRT3D_ZORIGIN, val, NULL); }
    void SetZRotation(double val) { Xrt3dSetValues(m_hChart, XRT3D_ZROTATION, val, NULL); }
    void SetZScale(double val) { Xrt3dSetValues(m_hChart, XRT3D_ZSCALE, val, NULL); }
    void SetZValueLabels(Xrt3dValueLabel ** val) { Xrt3dSetValues(m_hChart, XRT3D_ZVALUE_LABELS, val, NULL); }
    void SetZoneData(Xrt3dData * val) { Xrt3dSetValues(m_hChart, XRT3D_ZONE_DATA, val, NULL); }
    void SetZoneMethod(Xrt3dZoneMethod val) { Xrt3dSetValues(m_hChart, XRT3D_ZONE_METHOD, val, NULL); }
	void SetFontRotation(Xrt3dFontRotation val) { Xrt3dSetValues(m_hChart, XRT3DT_FONT_ROTATION, val, NULL); }


    // Operations
    double ComputeZValue(int x, int y, int px, int py){ double value; Xrt3dComputeZValueIndirect(m_hChart, x, y, px, py, &value); return value; }
    BOOL DrawToClipboard(Xrt3dDrawFormat format = XRT3D_DRAW_BITMAP) { return Xrt3dDrawToClipboard(m_hChart, format); }
    BOOL DrawToDC(HDC hdc, Xrt3dDrawFormat format = XRT3D_DRAW_ENHMETAFILE, Xrt3dDrawScale scale = XRT3D_DRAWSCALE_TOFIT, int left = 0, int top = 0, int width = 0, int height = 0) { return Xrt3dDrawToDC(m_hChart, hdc, format, scale, left, top, width, height); }
    BOOL DrawToFile(TCHAR *file, Xrt3dDrawFormat format = XRT3D_DRAW_BITMAP) { return Xrt3dDrawToFile(m_hChart, file, format); }
    Xrt3dContourStyle *GetNthContourStyle(int index, int used){ return Xrt3dGetNthContourStyle(m_hChart, index, used); }
    TCHAR *GetNthDataLabel(Xrt3dAxis axis, int index){ return Xrt3dGetNthDataLabel(m_hChart, axis, index); }
    BOOL GetPropString(int res, TCHAR **str) { return Xrt3dGetPropString(m_hChart, res, str); }
    Xrt3dValueLabel *GetValueLabel(Xrt3dAxis axis, Xrt3dValueLabel *vlabel){ return Xrt3dGetValueLabel(m_hChart, axis, vlabel); }
    Xrt3dRegion Map(int x, int y, Xrt3dMapResult *map){ return Xrt3dMap(m_hChart, x, y, map); }
    Xrt3dRegion Pick(int x, int y, Xrt3dPickResult *pick){ return Xrt3dPick(m_hChart, x, y, pick); }
    BOOL Print(Xrt3dDrawFormat format = XRT3D_DRAW_ENHMETAFILE, Xrt3dDrawScale scale = XRT3D_DRAWSCALE_TOFIT, int left = 0, int top = 0, int width = 0, int height = 0) { return Xrt3dPrint(m_hChart, format, scale, left, top, width, height); }
    BOOL SaveImageAsJpeg(TCHAR *filename, int quality = 80, BOOL grayscale = FALSE, BOOL optimize = FALSE, BOOL progressive = FALSE) { return Xrt3dSaveImageAsJpeg(m_hChart, filename, quality, grayscale, optimize, progressive); }
    BOOL SaveImageAsPng(TCHAR *filename, BOOL interlace = FALSE) { return Xrt3dSaveImageAsPng(m_hChart, filename, interlace); }
    void SetNthContourStyle(int index, Xrt3dContourStyle *cs, int used){ Xrt3dSetNthContourStyle(m_hChart, index, cs, used); }
    void SetNthDataLabel(Xrt3dAxis axis, int index, TCHAR *label){ Xrt3dSetNthDataLabel(m_hChart, axis, index, label); }
    BOOL SetPropString(int res, TCHAR *str) { return Xrt3dSetPropString(m_hChart, res, str); }
    void SetValueLabel(Xrt3dAxis axis, Xrt3dValueLabel *vlabel){ Xrt3dSetValueLabel(m_hChart, axis, vlabel); }
    void SetXYColor(int x, int y, COLORREF color){ Xrt3dSetXYColor(m_hChart, x, y, color); }
    void Unmap(double x, double y, double z, Xrt3dMapResult *map){ Xrt3dUnmap(m_hChart, x, y, z, map); }
    void Unpick(int x, int y, Xrt3dPickResult *pick){ Xrt3dUnpick(m_hChart, x, y, pick); }
    double ZInterpolate(double x, double y){ double value; Xrt3dZInterpolateIndirect(m_hChart, x, y, &value); return value; }

protected:
    // Implementation
    virtual BOOL OnChildNotify(UINT, WPARAM, LPARAM, LRESULT*) { return(FALSE); }
};


/*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 *
 *  Class CChart3DTextArea
 *
 *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
*/

class CChart3DTextArea {
public:
    // Constructor
    CChart3DTextArea(CChart3D *chart) { m_hText = Xrt3dTextCreate(chart->m_hChart); }

    // Destructor
    ~CChart3DTextArea() { Xrt3dTextDestroy(m_hText); }

    // Get methods
    Xrt3dAdjust GetAdjust() { Xrt3dAdjust val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ADJUST, &val, NULL); return(val); }
    int GetAttachIndexPoint() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_INDEX_POINT, &val, NULL); return(val); }
    int GetAttachIndexSeries() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_INDEX_SERIES, &val, NULL); return(val); }
    int GetAttachIndexX() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_INDEX_X, &val, NULL); return(val); }
    int GetAttachIndexY() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_INDEX_Y, &val, NULL); return(val); }
    Xrt3dTextAttachMethod GetAttachMethod() { Xrt3dTextAttachMethod val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_METHOD, &val, NULL); return(val); }
    int GetAttachPixelX() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_PIXEL_X, &val, NULL); return(val); }
    int GetAttachPixelY() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_PIXEL_Y, &val, NULL); return(val); }
    double GetAttachPointX() { double val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_POINT_X, &val, NULL); return(val); }
    double GetAttachPointY() { double val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_POINT_Y, &val, NULL); return(val); }
    double GetAttachPointZ() { double val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_ATTACH_POINT_Z, &val, NULL); return(val); }
    COLORREF GetBackgroundColor() { COLORREF val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_BACKGROUND_COLOR, &val, NULL); return(val); }
    Xrt3dBorder GetBorder() { Xrt3dBorder val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_BORDER, &val, NULL); return(val); }
    int GetBorderWidth() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_BORDER_WIDTH, &val, NULL); return(val); }
    HFONT GetFont() { HFONT val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_FONT, &val, NULL); return(val); }
    COLORREF GetForegroundColor() { COLORREF val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_FOREGROUND_COLOR, &val, NULL); return(val); }
    Xrt3dImage GetImage() { Xrt3dImage val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_IMAGE, &val, NULL); return(val); }
    Xrt3dImageLayout GetImageLayout() { Xrt3dImageLayout val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetImageTransparent() { BOOL val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    BOOL GetImageMinimumSize() { BOOL val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_IMAGE_MINIMUM_SIZE, &val, NULL); return(val); }
    TCHAR * GetImageObsolete() { TCHAR * val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_IMAGE_OBSOLETE, &val, NULL); return(val); }
    BOOL GetLineShow() { BOOL val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_LINE_SHOW, &val, NULL); return(val); }
    int GetOffsetX() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_OFFSET_X, &val, NULL); return(val); }
    int GetOffsetY() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_OFFSET_Y, &val, NULL); return(val); }
    int GetPlane() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_PLANE, &val, NULL); return(val); }
    TCHAR * GetPrintFont() { TCHAR * val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_PRINT_FONT, &val, NULL); return(val); }
    BOOL GetShow() { BOOL val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_SHOW, &val, NULL); return(val); }
    TCHAR ** GetStrings() { TCHAR ** val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_STRINGS, &val, NULL); return(val); }
    Xrt3dStrokeFont GetStrokeFont() { Xrt3dStrokeFont val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_STROKE_FONT, &val, NULL); return(val); }
    int GetStrokeSize() { int val; Xrt3dTextGetValues(m_hText, XRT3D_TEXT_STROKE_SIZE, &val, NULL); return(val); }

    // Set methods
    void SetAdjust(Xrt3dAdjust val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ADJUST, val, NULL); }
    void SetAttachIndexPoint(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_INDEX_POINT, val, NULL); }
    void SetAttachIndexSeries(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_INDEX_SERIES, val, NULL); }
    void SetAttachIndexX(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_INDEX_X, val, NULL); }
    void SetAttachIndexY(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_INDEX_Y, val, NULL); }
    void SetAttachMethod(Xrt3dTextAttachMethod val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_METHOD, val, NULL); }
    void SetAttachPixelX(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_PIXEL_X, val, NULL); }
    void SetAttachPixelY(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_PIXEL_Y, val, NULL); }
    void SetAttachPointX(double val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_POINT_X, val, NULL); }
    void SetAttachPointY(double val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_POINT_Y, val, NULL); }
    void SetAttachPointZ(double val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_ATTACH_POINT_Z, val, NULL); }
    void SetBackgroundColor(COLORREF val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_BACKGROUND_COLOR, val, NULL); }
    void SetBorder(Xrt3dBorder val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_BORDER, val, NULL); }
    void SetBorderWidth(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_BORDER_WIDTH, val, NULL); }
    void SetFont(HFONT val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_FONT, (int)val, NULL); }
    void SetForegroundColor(COLORREF val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_FOREGROUND_COLOR, val, NULL); }
    void SetImage(Xrt3dImage val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_IMAGE, val, NULL); }
    void SetImageLayout(Xrt3dImageLayout val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_IMAGE_LAYOUT, val, NULL); }
    void SetImageTransparent(BOOL val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_IMAGE_TRANSPARENT, val, NULL); }
    void SetImageMinimumSize(BOOL val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_IMAGE_MINIMUM_SIZE, val, NULL); }
    void SetImageObsolete(TCHAR * val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_IMAGE_OBSOLETE, val, NULL); }
    void SetLineShow(BOOL val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_LINE_SHOW, val, NULL); }
    void SetOffsetX(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_OFFSET_X, val, NULL); }
    void SetOffsetY(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_OFFSET_Y, val, NULL); }
    void SetPlane(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_PLANE, val, NULL); }
    void SetPrintFont(TCHAR * val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_PRINT_FONT, val, NULL); }
    void SetShow(BOOL val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_SHOW, val, NULL); }
    void SetStrings(TCHAR ** val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_STRINGS, val, NULL); }
    void SetStrokeFont(Xrt3dStrokeFont val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_STROKE_FONT, val, NULL); }
    void SetStrokeSize(int val) { Xrt3dTextSetValues(m_hText, XRT3D_TEXT_STROKE_SIZE, val, NULL); }

private:
    HXRT3DTEXT m_hText;
};


/*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 *
 *  Class CChart3DData
 *
 *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
*/

typedef Xrt3dData * Xrt3dDataP;

class CChart3DData {
public:
    // Constructors
    CChart3DData(int numx, int numy, double noval, double xstep, double ystep, double xorig, double yorig) { data = Xrt3dMakeGridData(numx, numy, noval, xstep, ystep, xorig, yorig, TRUE); }
    CChart3DData(int numx, int numy, double noval) { data = Xrt3dMakeIrGridData(numx, numy, noval, TRUE); }
    CChart3DData(TCHAR *fname) { data = Xrt3dMakeDataFromFile(fname, NULL); }
    CChart3DData(CChart3DData& d) { data = Xrt3dDataCopy(d.data); }

    // Destructor
    ~CChart3DData() { Xrt3dDestroyData(data, TRUE); }

    // conversion to Xrt3dData *
    operator Xrt3dDataP() { return(data); }
    void operator=(CChart3DData& d) { if (this != &d)  { Xrt3dDestroyData(data, TRUE); data = Xrt3dDataCopy(d.data); } }

    // Xrt3dData access macros
    Xrt3dData* GetData() { return (data); }
    Xrt3dDataType GetType() { return(data->g.type); }
    int GetNumX() { return(data->g.numx); }
    int GetNumY() { return(data->g.numy); }
    double GetNoVal() { return(data->g.noval); }
    double GetXStep() { return( data->g.type == XRT3D_DATA_GRID ? data->g.xstep : 0.0 ); }
    double GetYStep() { return( data->g.type == XRT3D_DATA_GRID ? data->g.ystep : 0.0 ); }
    double GetXOrig() { return( data->g.type == XRT3D_DATA_GRID ? data->g.xorig : data->ig.xgrid[0] ); }
    double GetYOrig() { return( data->g.type == XRT3D_DATA_GRID ? data->g.yorig : data->ig.ygrid[0] ); }
    double GetValue(int i, int j) { return((i >= 0 && i < data->g.numx && j >=0 && j < data->g.numy) ? (data->g.type == XRT3D_DATA_GRID ? data->g.values[i][j] : data->ig.values[i][j]) : data->g.noval); }
    void SetValue(int i, int j, double value)
        {
            if (i >= 0 && i < data->g.numx && j >=0 && j < data->g.numy)
            {
                if (data->g.type == XRT3D_DATA_GRID)    {
                    data->g.values[i][j] = value;
                }
                else    {
                    data->ig.values[i][j] = value;
                }
            }
        }

    // Operations
    Xrt3dData* DataShaded(double sweep, double rise, double scale, double ambient, double intensity) { return Xrt3dDataShaded(data, sweep, rise, scale, ambient, intensity); }
    void DataSmooth(double center_weight) { Xrt3dDataSmooth(data, center_weight); }
    Xrt3dData* DataWindow(double xstart, double ystart, double xend, double yend, int numx, int numy, Xrt3dInterpMethod interp) { return Xrt3dDataWindow(data, xstart, ystart, xend, yend, numx, numy, interp); }
    int SaveToFile(TCHAR *fname) { return Xrt3dSaveDataToFile(data, fname, NULL); }

    private:
        Xrt3dData *data;
};


#endif
