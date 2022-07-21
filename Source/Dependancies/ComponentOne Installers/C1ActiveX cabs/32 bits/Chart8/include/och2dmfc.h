#if !defined(__OCH2DMFC_H__)
#define __OCH2DMFC_H__

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
 * The main class, CChart2D, derives from CWnd. The majority of native
 * ComponentOne Chart DLL calls require a handle to the chart (usually shown in
 * the DLL manual as HXRT2D hChart) as the first parameter. CChart2D wraps
 * these calls and provides straightforward names for the access functions.
 *
 * As with a CWnd, you construct a CChart2D object in two steps. First, call
 * the constructor, which in turn constructs the associated CWnd object.
 * Then call the Create method, which creates and sets up the CWnd, then
 * creates the Chart and attaches the Windows child window to it.
 * See the MFC help for CWnd for the parameters to be passed to Create().
 * A caption is unnecessary. A typical example looks like this:
 *     m_chart.Create(, WS_CHILD|WS_VISIBLE, rect, this, 0);
 *
 * ------------------------------------
 * The CChart2DText class is normally created via a call to "new". Always
 * remember to free these objects via "delete" at the end of execution!
 *
 * CChart2DText wraps ComponentOne Chart DLL calls that require a data handle
 * (usually shown in the DLL manual as XrtTextHandle text) as the first
 * parameter. CChart2DText wraps these calls and provides more
 * straightforward names for the access functions.
 *
 * ------------------------------------
 * The CChart2DPointStyle class is normally created via a call to "new".
 * Always remember to free these objects via "delete" at the end of
 * execution!
 *
 * CChart2DPointStyle wraps ComponentOne Chart DLL calls that require a point
 * style handle (usually shown in the DLL manual as XrtPointStyleHandle)
 * as the first parameter. CChart2DPointStyle wraps these calls and
 * provides more straightforward names for the access functions.
 *
 * ------------------------------------
 * The CChart2DData class is normally created via a call to "new". Always
 * remember to free these objects via "delete" at the end of execution!
 *
 * CChart2DData objects can be constructed in one of four ways:
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
 * CChart2DData wraps ComponentOne Chart DLL calls that require a data handle
 * (usually shown in the DLL manual as XrtDataHandle hData) as the first
 * parameter. CChart2DData wraps these calls and provides more
 * straightforward names for the access functions.
 *
 * ==========================================================================
 */

#include <olch2d.h>

/*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 *
 *  Class CChart2D
 *
 *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
*/

class CChart2D : public CWnd
{
    DECLARE_DYNAMIC(CChart2D)

public:
    // Constructors
    CChart2D() { m_hChart = NULL; };
    BOOL Create(LPCTSTR lpszCaption, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID);

    HXRT2D m_hChart;

#if (_MFC_VER < 0x0300)
    WNDPROC* GetSuperWndProcAddr();
#endif

    // Destructor
    virtual ~CChart2D() { DestroyWindow(); if (m_hChart) { XrtDestroy(m_hChart); } }

    // Get methods
    XrtAlarmZone * GetAlarmZone(const TCHAR *name) {return XrtGetAlarmZone(m_hChart, name);}
    XrtAlarmZone * GetAlarmZoneList() {return XrtGetAlarmZoneList(m_hChart);}
    XrtAngleUnit GetAngleUnit() { XrtAngleUnit val; XrtGetValues(m_hChart, XRT_ANGLE_UNIT, &val, NULL); return(val); }
    BOOL GetAxisBoundingBox() { BOOL val; XrtGetValues(m_hChart, XRT_AXIS_BOUNDING_BOX, &val, NULL); return(val); }
    HFONT GetAxisFont() { HFONT val; XrtGetValues(m_hChart, XRT_AXIS_FONT, &val, NULL); return(val); }
    COLORREF GetBackgroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_BACKGROUND_COLOR, &val, NULL); return(val); }
    int GetBarClusterOverlap() { int val; XrtGetValues(m_hChart, XRT_BAR_CLUSTER_OVERLAP, &val, NULL); return(val); }
    int GetBarClusterWidth() { int val; XrtGetValues(m_hChart, XRT_BAR_CLUSTER_WIDTH, &val, NULL); return(val); }
    XrtBorder GetBorder() { XrtBorder val; XrtGetValues(m_hChart, XRT_BORDER, &val, NULL); return(val); }
    int GetBorderWidth() { int val; XrtGetValues(m_hChart, XRT_BORDER_WIDTH, &val, NULL); return(val); }
    double GetBubbleMax() { double val; XrtGetValues(m_hChart, XRT_BUBBLE_MAX, &val, NULL); return(val); }
    XrtBubbleMethod GetBubbleMethod() { XrtBubbleMethod val; XrtGetValues(m_hChart, XRT_BUBBLE_METHOD, &val, NULL); return(val); }
    double GetBubbleMin() { double val; XrtGetValues(m_hChart, XRT_BUBBLE_MIN, &val, NULL); return(val); }
    BOOL GetCandleComplex() { BOOL val; XrtGetValues(m_hChart, XRT_CANDLE_COMPLEX, &val, NULL); return(val); }
	BOOL GetCandleFillFalling() { BOOL val; XrtGetValues(m_hChart, XRT_CANDLE_FILLFALLING, &val, NULL); return(val); }
    XrtDataHandle GetData() { XrtDataHandle val; XrtGetValues(m_hChart, XRT_DATA, &val, NULL); return(val); }
    XrtDataHandle GetData2() { XrtDataHandle val; XrtGetValues(m_hChart, XRT_DATA2, &val, NULL); return(val); }
    COLORREF GetDataAreaBackgroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_DATA_AREA_BACKGROUND_COLOR, &val, NULL); return(val); }
    COLORREF GetDataAreaForegroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_DATA_AREA_FOREGROUND_COLOR, &val, NULL); return(val); }
    XrtImage GetDataAreaImage() { XrtImage val; XrtGetValues(m_hChart, XRT_DATA_AREA_IMAGE, &val, NULL); return(val); }
    XrtImageLayout GetDataAreaImageLayout() { XrtImageLayout val; XrtGetValues(m_hChart, XRT_DATA_AREA_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetDataAreaImageTransparent() { BOOL val; XrtGetValues(m_hChart, XRT_DATA_AREA_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    TCHAR * GetDataAreaImageObsolete() { TCHAR * val; XrtGetValues(m_hChart, XRT_DATA_AREA_IMAGE_OBSOLETE, &val, NULL); return(val); }
    XrtDataStyle ** GetDataStyles() { XrtDataStyle ** val; XrtGetValues(m_hChart, XRT_DATA_STYLES, &val, NULL); return(val); }
    XrtDataStyle ** GetDataStyles2() { XrtDataStyle ** val; XrtGetValues(m_hChart, XRT_DATA_STYLES2, &val, NULL); return(val); }
    BOOL GetDataStyles2UseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_DATA_STYLES2_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetDataStylesUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_DATA_STYLES_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetDebug() { BOOL val; XrtGetValues(m_hChart, XRT_DEBUG, &val, NULL); return(val); }
    BOOL GetDoubleBuffer() { BOOL val; XrtGetValues(m_hChart, XRT_DOUBLE_BUFFER, &val, NULL); return(val); }
    BOOL GetExtraDefaultDataStyles() { BOOL val; XrtGetValues(m_hChart, XRT_EXTRA_DEFAULT_DATA_STYLES, &val, NULL); return(val); }
    XrtAdjust GetFooterAdjust() { XrtAdjust val; XrtGetValues(m_hChart, XRT_FOOTER_ADJUST, &val, NULL); return(val); }
    COLORREF GetFooterBackgroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_FOOTER_BACKGROUND_COLOR, &val, NULL); return(val); }
    XrtBorder GetFooterBorder() { XrtBorder val; XrtGetValues(m_hChart, XRT_FOOTER_BORDER, &val, NULL); return(val); }
    int GetFooterBorderWidth() { int val; XrtGetValues(m_hChart, XRT_FOOTER_BORDER_WIDTH, &val, NULL); return(val); }
    HFONT GetFooterFont() { HFONT val; XrtGetValues(m_hChart, XRT_FOOTER_FONT, &val, NULL); return(val); }
    COLORREF GetFooterForegroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_FOOTER_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetFooterHeight() { int val; XrtGetValues(m_hChart, XRT_FOOTER_HEIGHT, &val, NULL); return(val); }
    XrtImage GetFooterImage() { XrtImage val; XrtGetValues(m_hChart, XRT_FOOTER_IMAGE, &val, NULL); return(val); }
    XrtImageLayout GetFooterImageLayout() { XrtImageLayout val; XrtGetValues(m_hChart, XRT_FOOTER_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetFooterImageTransparent() { BOOL val; XrtGetValues(m_hChart, XRT_FOOTER_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    BOOL GetFooterImageMinimumSize() { BOOL val; XrtGetValues(m_hChart, XRT_FOOTER_IMAGE_MINIMUM_SIZE, &val, NULL); return(val); }
    TCHAR * GetFooterImageObsolete() { TCHAR * val; XrtGetValues(m_hChart, XRT_FOOTER_IMAGE_OBSOLETE, &val, NULL); return(val); }
    TCHAR ** GetFooterStrings() { TCHAR ** val; XrtGetValues(m_hChart, XRT_FOOTER_STRINGS, &val, NULL); return(val); }
    int GetFooterWidth() { int val; XrtGetValues(m_hChart, XRT_FOOTER_WIDTH, &val, NULL); return(val); }
    int GetFooterX() { int val; XrtGetValues(m_hChart, XRT_FOOTER_X, &val, NULL); return(val); }
    BOOL GetFooterXUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_FOOTER_X_USE_DEFAULT, &val, NULL); return(val); }
    int GetFooterY() { int val; XrtGetValues(m_hChart, XRT_FOOTER_Y, &val, NULL); return(val); }
    BOOL GetFooterYUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_FOOTER_Y_USE_DEFAULT, &val, NULL); return(val); }
    COLORREF GetForegroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetFrontDataset() { int val; XrtGetValues(m_hChart, XRT_FRONT_DATASET, &val, NULL); return(val); }
    XrtShading GetGraph3dShading() { XrtShading val; XrtGetValues(m_hChart, XRT_GRAPH_3D_SHADING, &val, NULL); return(val); }
    COLORREF GetGraphBackgroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_GRAPH_BACKGROUND_COLOR, &val, NULL); return(val); }
    XrtBorder GetGraphBorder() { XrtBorder val; XrtGetValues(m_hChart, XRT_GRAPH_BORDER, &val, NULL); return(val); }
    int GetGraphBorderWidth() { int val; XrtGetValues(m_hChart, XRT_GRAPH_BORDER_WIDTH, &val, NULL); return(val); }
    int GetGraphDepth() { int val; XrtGetValues(m_hChart, XRT_GRAPH_DEPTH, &val, NULL); return(val); }
    COLORREF GetGraphForegroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_GRAPH_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetGraphHeight() { int val; XrtGetValues(m_hChart, XRT_GRAPH_HEIGHT, &val, NULL); return(val); }
    BOOL GetGraphHeightUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_HEIGHT_USE_DEFAULT, &val, NULL); return(val); }
    XrtImage GetGraphImage() { XrtImage val; XrtGetValues(m_hChart, XRT_GRAPH_IMAGE, &val, NULL); return(val); }
    XrtImageLayout GetGraphImageLayout() { XrtImageLayout val; XrtGetValues(m_hChart, XRT_GRAPH_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetGraphImageTransparent() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    TCHAR * GetGraphImageObsolete() { TCHAR * val; XrtGetValues(m_hChart, XRT_GRAPH_IMAGE_OBSOLETE, &val, NULL); return(val); }
    int GetGraphInclination() { int val; XrtGetValues(m_hChart, XRT_GRAPH_INCLINATION, &val, NULL); return(val); }
    int GetGraphMarginBottom() { int val; XrtGetValues(m_hChart, XRT_GRAPH_MARGIN_BOTTOM, &val, NULL); return(val); }
    BOOL GetGraphMarginBottomUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_MARGIN_BOTTOM_USE_DEFAULT, &val, NULL); return(val); }
    int GetGraphMarginLeft() { int val; XrtGetValues(m_hChart, XRT_GRAPH_MARGIN_LEFT, &val, NULL); return(val); }
    BOOL GetGraphMarginLeftUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_MARGIN_LEFT_USE_DEFAULT, &val, NULL); return(val); }
    int GetGraphMarginRight() { int val; XrtGetValues(m_hChart, XRT_GRAPH_MARGIN_RIGHT, &val, NULL); return(val); }
    BOOL GetGraphMarginRightUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_MARGIN_RIGHT_USE_DEFAULT, &val, NULL); return(val); }
    int GetGraphMarginTop() { int val; XrtGetValues(m_hChart, XRT_GRAPH_MARGIN_TOP, &val, NULL); return(val); }
    BOOL GetGraphMarginTopUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_MARGIN_TOP_USE_DEFAULT, &val, NULL); return(val); }
    int GetGraphRotation() { int val; XrtGetValues(m_hChart, XRT_GRAPH_ROTATION, &val, NULL); return(val); }
    BOOL GetGraphShowOutlines() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_SHOW_OUTLINES, &val, NULL); return(val); }
    int GetGraphWidth() { int val; XrtGetValues(m_hChart, XRT_GRAPH_WIDTH, &val, NULL); return(val); }
    BOOL GetGraphWidthUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_WIDTH_USE_DEFAULT, &val, NULL); return(val); }
    int GetGraphX() { int val; XrtGetValues(m_hChart, XRT_GRAPH_X, &val, NULL); return(val); }
    BOOL GetGraphXUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_X_USE_DEFAULT, &val, NULL); return(val); }
    int GetGraphY() { int val; XrtGetValues(m_hChart, XRT_GRAPH_Y, &val, NULL); return(val); }
    BOOL GetGraphYUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_GRAPH_Y_USE_DEFAULT, &val, NULL); return(val); }
    XrtAdjust GetHeaderAdjust() { XrtAdjust val; XrtGetValues(m_hChart, XRT_HEADER_ADJUST, &val, NULL); return(val); }
    COLORREF GetHeaderBackgroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_HEADER_BACKGROUND_COLOR, &val, NULL); return(val); }
    XrtBorder GetHeaderBorder() { XrtBorder val; XrtGetValues(m_hChart, XRT_HEADER_BORDER, &val, NULL); return(val); }
    int GetHeaderBorderWidth() { int val; XrtGetValues(m_hChart, XRT_HEADER_BORDER_WIDTH, &val, NULL); return(val); }
    HFONT GetHeaderFont() { HFONT val; XrtGetValues(m_hChart, XRT_HEADER_FONT, &val, NULL); return(val); }
    COLORREF GetHeaderForegroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_HEADER_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetHeaderHeight() { int val; XrtGetValues(m_hChart, XRT_HEADER_HEIGHT, &val, NULL); return(val); }
    XrtImage GetHeaderImage() { XrtImage val; XrtGetValues(m_hChart, XRT_HEADER_IMAGE, &val, NULL); return(val); }
    XrtImageLayout GetHeaderImageLayout() { XrtImageLayout val; XrtGetValues(m_hChart, XRT_HEADER_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetHeaderImageTransparent() { BOOL val; XrtGetValues(m_hChart, XRT_HEADER_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    BOOL GetHeaderImageMinimumSize() { BOOL val; XrtGetValues(m_hChart, XRT_HEADER_IMAGE_MINIMUM_SIZE, &val, NULL); return(val); }
    TCHAR * GetHeaderImageObsolete() { TCHAR * val; XrtGetValues(m_hChart, XRT_HEADER_IMAGE_OBSOLETE, &val, NULL); return(val); }
    TCHAR ** GetHeaderStrings() { TCHAR ** val; XrtGetValues(m_hChart, XRT_HEADER_STRINGS, &val, NULL); return(val); }
    int GetHeaderWidth() { int val; XrtGetValues(m_hChart, XRT_HEADER_WIDTH, &val, NULL); return(val); }
    int GetHeaderX() { int val; XrtGetValues(m_hChart, XRT_HEADER_X, &val, NULL); return(val); }
    BOOL GetHeaderXUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_HEADER_X_USE_DEFAULT, &val, NULL); return(val); }
    int GetHeaderY() { int val; XrtGetValues(m_hChart, XRT_HEADER_Y, &val, NULL); return(val); }
    BOOL GetHeaderYUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_HEADER_Y_USE_DEFAULT, &val, NULL); return(val); }
    int GetHeight() { int val; XrtGetValues(m_hChart, XRT_HEIGHT, &val, NULL); return(val); }
    BOOL GetHiloCloseShow() { BOOL val; XrtGetValues(m_hChart, XRT_HILO_CLOSE_SHOW, &val, NULL); return(val); }
    BOOL GetHiloOpenCloseFullWidth() { BOOL val; XrtGetValues(m_hChart, XRT_HILO_OPEN_CLOSE_FULL_WIDTH, &val, NULL); return(val); }
    BOOL GetHiloOpenShow() { BOOL val; XrtGetValues(m_hChart, XRT_HILO_OPEN_SHOW, &val, NULL); return(val); }
    XrtImage GetImage() { XrtImage val; XrtGetValues(m_hChart, XRT_IMAGE, &val, NULL); return(val); }
    XrtImageLayout GetImageLayout() { XrtImageLayout val; XrtGetValues(m_hChart, XRT_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetImageTransparent() { BOOL val; XrtGetValues(m_hChart, XRT_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    TCHAR * GetImageObsolete() { TCHAR * val; XrtGetValues(m_hChart, XRT_IMAGE_OBSOLETE, &val, NULL); return(val); }
    BOOL GetInvertOrientation() { BOOL val; XrtGetValues(m_hChart, XRT_INVERT_ORIENTATION, &val, NULL); return(val); }
    BOOL GetIsStacked() { BOOL val; XrtGetValues(m_hChart, XRT_IS_STACKED, &val, NULL); return(val); }
    BOOL GetIsStacked2() { BOOL val; XrtGetValues(m_hChart, XRT_IS_STACKED2, &val, NULL); return(val); }
    XrtAnchor GetLegendAnchor() { XrtAnchor val; XrtGetValues(m_hChart, XRT_LEGEND_ANCHOR, &val, NULL); return(val); }
    COLORREF GetLegendBackgroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_LEGEND_BACKGROUND_COLOR, &val, NULL); return(val); }
    XrtBorder GetLegendBorder() { XrtBorder val; XrtGetValues(m_hChart, XRT_LEGEND_BORDER, &val, NULL); return(val); }
    int GetLegendBorderWidth() { int val; XrtGetValues(m_hChart, XRT_LEGEND_BORDER_WIDTH, &val, NULL); return(val); }
    HFONT GetLegendFont() { HFONT val; XrtGetValues(m_hChart, XRT_LEGEND_FONT, &val, NULL); return(val); }
    COLORREF GetLegendForegroundColor() { COLORREF val; XrtGetValues(m_hChart, XRT_LEGEND_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetLegendHeight() { int val; XrtGetValues(m_hChart, XRT_LEGEND_HEIGHT, &val, NULL); return(val); }
    XrtImage GetLegendImage() { XrtImage val; XrtGetValues(m_hChart, XRT_LEGEND_IMAGE, &val, NULL); return(val); }
    XrtImageLayout GetLegendImageLayout() { XrtImageLayout val; XrtGetValues(m_hChart, XRT_LEGEND_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetLegendImageTransparent() { BOOL val; XrtGetValues(m_hChart, XRT_LEGEND_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    TCHAR * GetLegendImageObsolete() { TCHAR * val; XrtGetValues(m_hChart, XRT_LEGEND_IMAGE_OBSOLETE, &val, NULL); return(val); }
    XrtAlign GetLegendOrientation() { XrtAlign val; XrtGetValues(m_hChart, XRT_LEGEND_ORIENTATION, &val, NULL); return(val); }
    BOOL GetLegendReversed() { BOOL val; XrtGetValues(m_hChart, XRT_LEGEND_REVERSED, &val, NULL); return(val); }
    BOOL GetLegendShow() { BOOL val; XrtGetValues(m_hChart, XRT_LEGEND_SHOW, &val, NULL); return(val); }
    TCHAR * GetLegendTitle() { TCHAR * val; XrtGetValues(m_hChart, XRT_LEGEND_TITLE, &val, NULL); return(val); }
    int GetLegendWidth() { int val; XrtGetValues(m_hChart, XRT_LEGEND_WIDTH, &val, NULL); return(val); }
    int GetLegendX() { int val; XrtGetValues(m_hChart, XRT_LEGEND_X, &val, NULL); return(val); }
    BOOL GetLegendXUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_LEGEND_X_USE_DEFAULT, &val, NULL); return(val); }
    int GetLegendY() { int val; XrtGetValues(m_hChart, XRT_LEGEND_Y, &val, NULL); return(val); }
    BOOL GetLegendYUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_LEGEND_Y_USE_DEFAULT, &val, NULL); return(val); }
    XrtDataStyle * GetMarkerDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_MARKER_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetMarkerDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_MARKER_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    int GetMarkerDataset() { int val; XrtGetValues(m_hChart, XRT_MARKER_DATASET, &val, NULL); return(val); }
    TCHAR * GetName() { TCHAR * val; XrtGetValues(m_hChart, XRT_NAME, &val, NULL); return(val); }
    XrtDataStyle * GetOtherDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_OTHER_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetOtherDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_OTHER_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR * GetOtherLabel() { TCHAR * val; XrtGetValues(m_hChart, XRT_OTHER_LABEL, &val, NULL); return(val); }
    BOOL GetPieMergeMissingSlices() { BOOL val; XrtGetValues(m_hChart, XRT_PIE_MERGE_MISSING_SLICES, &val, NULL); return(val); }
    int GetPieMinSlices() { int val; XrtGetValues(m_hChart, XRT_PIE_MIN_SLICES, &val, NULL); return(val); }
    XrtPieOrder GetPieOrder() { XrtPieOrder val; XrtGetValues(m_hChart, XRT_PIE_ORDER, &val, NULL); return(val); }
    double GetPieStartAngle() { double val; XrtGetValues(m_hChart, XRT_PIE_START_ANGLE, &val, NULL); return(val); }
    XrtPieThresholdMethod GetPieThresholdMethod() { XrtPieThresholdMethod val; XrtGetValues(m_hChart, XRT_PIE_THRESHOLD_METHOD, &val, NULL); return(val); }
    double GetPieThresholdValue() { double val; XrtGetValues(m_hChart, XRT_PIE_THRESHOLD_VALUE, &val, NULL); return(val); }
    TCHAR ** GetPointLabels() { TCHAR ** val; XrtGetValues(m_hChart, XRT_POINT_LABELS, &val, NULL); return(val); }
    TCHAR ** GetPointLabels2() { TCHAR ** val; XrtGetValues(m_hChart, XRT_POINT_LABELS2, &val, NULL); return(val); }
    BOOL GetPolarAxisAllowNegatives() { BOOL val; XrtGetValues(m_hChart, XRT_POLAR_AXIS_ALLOW_NEGATIVES, &val, NULL); return(val); }
    BOOL GetPolarHalfRange() { BOOL val; XrtGetValues(m_hChart, XRT_POLAR_HALF_RANGE, &val, NULL); return(val); }
    BOOL GetRepaint() { BOOL val; XrtGetValues(m_hChart, XRT_REPAINT, &val, NULL); return(val); }
    TCHAR ** GetSetLabels() { TCHAR ** val; XrtGetValues(m_hChart, XRT_SET_LABELS, &val, NULL); return(val); }
    TCHAR ** GetSetLabels2() { TCHAR ** val; XrtGetValues(m_hChart, XRT_SET_LABELS2, &val, NULL); return(val); }
    long GetTimeBase() { long val; XrtGetValues(m_hChart, XRT_TIME_BASE, &val, NULL); return(val); }
    double GetTimeBaseOle() { double val; XrtGetValues(m_hChart, XRT_TIME_BASE_OLE, &val, NULL); return(val); }
    TCHAR * GetTimeFormat() { TCHAR * val; XrtGetValues(m_hChart, XRT_TIME_FORMAT, &val, NULL); return(val); }
    BOOL GetTimeFormatUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_TIME_FORMAT_USE_DEFAULT, &val, NULL); return(val); }
    XrtTimeUnit GetTimeUnit() { XrtTimeUnit val; XrtGetValues(m_hChart, XRT_TIME_UNIT, &val, NULL); return(val); }
    BOOL GetTransposeData() { BOOL val; XrtGetValues(m_hChart, XRT_TRANSPOSE_DATA, &val, NULL); return(val); }
    XrtType GetType() { XrtType val; XrtGetValues(m_hChart, XRT_TYPE, &val, NULL); return(val); }
    XrtType GetType2() { XrtType val; XrtGetValues(m_hChart, XRT_TYPE2, &val, NULL); return(val); }
    int GetWidth() { int val; XrtGetValues(m_hChart, XRT_WIDTH, &val, NULL); return(val); }
    XrtAnnoPlacement GetXAnnoPlacement() { XrtAnnoPlacement val; XrtGetValues(m_hChart, XRT_XANNO_PLACEMENT, &val, NULL); return(val); }
    XrtAnnoMethod GetXAnnotationMethod() { XrtAnnoMethod val; XrtGetValues(m_hChart, XRT_XANNOTATION_METHOD, &val, NULL); return(val); }
    XrtRotate GetXAnnotationRotation() { XrtRotate val; XrtGetValues(m_hChart, XRT_XANNOTATION_ROTATION, &val, NULL); return(val); }
    double GetXAnnotationRotationAngle() { double val; XrtGetValues(m_hChart, XRT_XANNOTATION_ROTATION_ANGLE, &val, NULL); return(val); }
    XrtDataStyle * GetXAxisDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_XAXIS_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetXAxisDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XAXIS_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR * GetXAxisLabelFormat() { TCHAR * val; XrtGetValues(m_hChart, XRT_XAXIS_LABEL_FORMAT, &val, NULL); return(val); }
    BOOL GetXAxisLogarithmic() { BOOL val; XrtGetValues(m_hChart, XRT_XAXIS_LOGARITHMIC, &val, NULL); return(val); }
    double GetXAxisMax() { double val; XrtGetValues(m_hChart, XRT_XAXIS_MAX, &val, NULL); return(val); }
    BOOL GetXAxisMaxUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XAXIS_MAX_USE_DEFAULT, &val, NULL); return(val); }
    double GetXAxisMin() { double val; XrtGetValues(m_hChart, XRT_XAXIS_MIN, &val, NULL); return(val); }
    BOOL GetXAxisMinUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XAXIS_MIN_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetXAxisReversed() { BOOL val; XrtGetValues(m_hChart, XRT_XAXIS_REVERSED, &val, NULL); return(val); }
    BOOL GetXAxisShow() { BOOL val; XrtGetValues(m_hChart, XRT_XAXIS_SHOW, &val, NULL); return(val); }
    double GetXGrid() { double val; XrtGetValues(m_hChart, XRT_XGRID, &val, NULL); return(val); }
    XrtDataStyle * GetXGridDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_XGRID_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetXGridDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XGRID_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetXGridUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XGRID_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR ** GetXLabels() { TCHAR ** val; XrtGetValues(m_hChart, XRT_XLABELS, &val, NULL); return(val); }
    double GetXMarker() { double val; XrtGetValues(m_hChart, XRT_XMARKER, &val, NULL); return(val); }
    XrtDataStyle * GetXMarkerDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_XMARKER_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetXMarkerDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XMARKER_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    XrtMarkerMethod GetXMarkerMethod() { XrtMarkerMethod val; XrtGetValues(m_hChart, XRT_XMARKER_METHOD, &val, NULL); return(val); }
    int GetXMarkerPoint() { int val; XrtGetValues(m_hChart, XRT_XMARKER_POINT, &val, NULL); return(val); }
    int GetXMarkerSet() { int val; XrtGetValues(m_hChart, XRT_XMARKER_SET, &val, NULL); return(val); }
    BOOL GetXMarkerShow() { BOOL val; XrtGetValues(m_hChart, XRT_XMARKER_SHOW, &val, NULL); return(val); }
    double GetXMax() { double val; XrtGetValues(m_hChart, XRT_XMAX, &val, NULL); return(val); }
    BOOL GetXMaxUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XMAX_USE_DEFAULT, &val, NULL); return(val); }
    double GetXMin() { double val; XrtGetValues(m_hChart, XRT_XMIN, &val, NULL); return(val); }
    BOOL GetXMinUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XMIN_USE_DEFAULT, &val, NULL); return(val); }
    double GetXNum() { double val; XrtGetValues(m_hChart, XRT_XNUM, &val, NULL); return(val); }
    XrtNumMethod GetXNumMethod() { XrtNumMethod val; XrtGetValues(m_hChart, XRT_XNUM_METHOD, &val, NULL); return(val); }
    BOOL GetXNumUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XNUM_USE_DEFAULT, &val, NULL); return(val); }
    double GetXOrigin() { double val; XrtGetValues(m_hChart, XRT_XORIGIN, &val, NULL); return(val); }
    double GetXOriginBase() { double val; XrtGetValues(m_hChart, XRT_XORIGIN_BASE, &val, NULL); return(val); }
    XrtOriginPlacement GetXOriginPlacement() { XrtOriginPlacement val; XrtGetValues(m_hChart, XRT_XORIGIN_PLACEMENT, &val, NULL); return(val); }
    BOOL GetXOriginUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XORIGIN_USE_DEFAULT, &val, NULL); return(val); }
    int GetXPrecision() { int val; XrtGetValues(m_hChart, XRT_XPRECISION, &val, NULL); return(val); }
    BOOL GetXPrecisionUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XPRECISION_USE_DEFAULT, &val, NULL); return(val); }
    double GetXTick() { double val; XrtGetValues(m_hChart, XRT_XTICK, &val, NULL); return(val); }
    BOOL GetXTickUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_XTICK_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR * GetXTitle() { TCHAR * val; XrtGetValues(m_hChart, XRT_XTITLE, &val, NULL); return(val); }
    XrtRotate GetXTitleRotation() { XrtRotate val; XrtGetValues(m_hChart, XRT_XTITLE_ROTATION, &val, NULL); return(val); }
    XrtAnnoPlacement GetY2AnnoPlacement() { XrtAnnoPlacement val; XrtGetValues(m_hChart, XRT_Y2ANNO_PLACEMENT, &val, NULL); return(val); }
    double GetY2AnnotationAngle() { double val; XrtGetValues(m_hChart, XRT_Y2ANNOTATION_ANGLE, &val, NULL); return(val); }
    BOOL GetY2AnnotationAngleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2ANNOTATION_ANGLE_USE_DEFAULT, &val, NULL); return(val); }
    XrtAnnoMethod GetY2AnnotationMethod() { XrtAnnoMethod val; XrtGetValues(m_hChart, XRT_Y2ANNOTATION_METHOD, &val, NULL); return(val); }
    XrtRotate GetY2AnnotationRotation() { XrtRotate val; XrtGetValues(m_hChart, XRT_Y2ANNOTATION_ROTATION, &val, NULL); return(val); }
    double GetY2AnnotationRotationAngle() { double val; XrtGetValues(m_hChart, XRT_Y2ANNOTATION_ROTATION_ANGLE, &val, NULL); return(val); }
    BOOL GetY2Axis100Percent() { BOOL val; XrtGetValues(m_hChart, XRT_Y2AXIS_100_PERCENT, &val, NULL); return(val); }
    XrtDataStyle * GetY2AxisDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_Y2AXIS_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetY2AxisDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2AXIS_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR * GetY2AxisLabelFormat() { TCHAR * val; XrtGetValues(m_hChart, XRT_Y2AXIS_LABEL_FORMAT, &val, NULL); return(val); }
    BOOL GetY2AxisLogarithmic() { BOOL val; XrtGetValues(m_hChart, XRT_Y2AXIS_LOGARITHMIC, &val, NULL); return(val); }
    double GetY2AxisMax() { double val; XrtGetValues(m_hChart, XRT_Y2AXIS_MAX, &val, NULL); return(val); }
    BOOL GetY2AxisMaxUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2AXIS_MAX_USE_DEFAULT, &val, NULL); return(val); }
    double GetY2AxisMin() { double val; XrtGetValues(m_hChart, XRT_Y2AXIS_MIN, &val, NULL); return(val); }
    BOOL GetY2AxisMinUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2AXIS_MIN_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetY2AxisReversed() { BOOL val; XrtGetValues(m_hChart, XRT_Y2AXIS_REVERSED, &val, NULL); return(val); }
    BOOL GetY2AxisShow() { BOOL val; XrtGetValues(m_hChart, XRT_Y2AXIS_SHOW, &val, NULL); return(val); }
    double GetY2Grid() { double val; XrtGetValues(m_hChart, XRT_Y2GRID, &val, NULL); return(val); }
    XrtDataStyle * GetY2GridDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_Y2GRID_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetY2GridDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2GRID_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetY2GridUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2GRID_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR ** GetY2Labels() { TCHAR ** val; XrtGetValues(m_hChart, XRT_Y2LABELS, &val, NULL); return(val); }
    double GetY2Max() { double val; XrtGetValues(m_hChart, XRT_Y2MAX, &val, NULL); return(val); }
    BOOL GetY2MaxUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2MAX_USE_DEFAULT, &val, NULL); return(val); }
    double GetY2Min() { double val; XrtGetValues(m_hChart, XRT_Y2MIN, &val, NULL); return(val); }
    BOOL GetY2MinUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2MIN_USE_DEFAULT, &val, NULL); return(val); }
    double GetY2Num() { double val; XrtGetValues(m_hChart, XRT_Y2NUM, &val, NULL); return(val); }
    XrtNumMethod GetY2NumMethod() { XrtNumMethod val; XrtGetValues(m_hChart, XRT_Y2NUM_METHOD, &val, NULL); return(val); }
    BOOL GetY2NumUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2NUM_USE_DEFAULT, &val, NULL); return(val); }
    double GetY2Origin() { double val; XrtGetValues(m_hChart, XRT_Y2ORIGIN, &val, NULL); return(val); }
    XrtOriginPlacement GetY2OriginPlacement() { XrtOriginPlacement val; XrtGetValues(m_hChart, XRT_Y2ORIGIN_PLACEMENT, &val, NULL); return(val); }
    BOOL GetY2OriginUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2ORIGIN_USE_DEFAULT, &val, NULL); return(val); }
    int GetY2Precision() { int val; XrtGetValues(m_hChart, XRT_Y2PRECISION, &val, NULL); return(val); }
    BOOL GetY2PrecisionUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2PRECISION_USE_DEFAULT, &val, NULL); return(val); }
    double GetY2Tick() { double val; XrtGetValues(m_hChart, XRT_Y2TICK, &val, NULL); return(val); }
    BOOL GetY2TickUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_Y2TICK_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR * GetY2Title() { TCHAR * val; XrtGetValues(m_hChart, XRT_Y2TITLE, &val, NULL); return(val); }
    XrtRotate GetY2TitleRotation() { XrtRotate val; XrtGetValues(m_hChart, XRT_Y2TITLE_ROTATION, &val, NULL); return(val); }
    XrtAnnoPlacement GetYAnnoPlacement() { XrtAnnoPlacement val; XrtGetValues(m_hChart, XRT_YANNO_PLACEMENT, &val, NULL); return(val); }
    double GetYAnnotationAngle() { double val; XrtGetValues(m_hChart, XRT_YANNOTATION_ANGLE, &val, NULL); return(val); }
    BOOL GetYAnnotationAngleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YANNOTATION_ANGLE_USE_DEFAULT, &val, NULL); return(val); }
    XrtAnnoMethod GetYAnnotationMethod() { XrtAnnoMethod val; XrtGetValues(m_hChart, XRT_YANNOTATION_METHOD, &val, NULL); return(val); }
    XrtRotate GetYAnnotationRotation() { XrtRotate val; XrtGetValues(m_hChart, XRT_YANNOTATION_ROTATION, &val, NULL); return(val); }
    double GetYAnnotationRotationAngle() { double val; XrtGetValues(m_hChart, XRT_YANNOTATION_ROTATION_ANGLE, &val, NULL); return(val); }
    BOOL GetYAxis100Percent() { BOOL val; XrtGetValues(m_hChart, XRT_YAXIS_100_PERCENT, &val, NULL); return(val); }
    double GetYAxisConst() { double val; XrtGetValues(m_hChart, XRT_YAXIS_CONST, &val, NULL); return(val); }
    XrtDataStyle * GetYAxisDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_YAXIS_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetYAxisDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YAXIS_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR * GetYAxisLabelFormat() { TCHAR * val; XrtGetValues(m_hChart, XRT_YAXIS_LABEL_FORMAT, &val, NULL); return(val); }
    BOOL GetYAxisLogarithmic() { BOOL val; XrtGetValues(m_hChart, XRT_YAXIS_LOGARITHMIC, &val, NULL); return(val); }
    double GetYAxisMax() { double val; XrtGetValues(m_hChart, XRT_YAXIS_MAX, &val, NULL); return(val); }
    BOOL GetYAxisMaxUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YAXIS_MAX_USE_DEFAULT, &val, NULL); return(val); }
    double GetYAxisMin() { double val; XrtGetValues(m_hChart, XRT_YAXIS_MIN, &val, NULL); return(val); }
    BOOL GetYAxisMinUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YAXIS_MIN_USE_DEFAULT, &val, NULL); return(val); }
    double GetYAxisMult() { double val; XrtGetValues(m_hChart, XRT_YAXIS_MULT, &val, NULL); return(val); }
    BOOL GetYAxisReversed() { BOOL val; XrtGetValues(m_hChart, XRT_YAXIS_REVERSED, &val, NULL); return(val); }
    BOOL GetYAxisShow() { BOOL val; XrtGetValues(m_hChart, XRT_YAXIS_SHOW, &val, NULL); return(val); }
    double GetYGrid() { double val; XrtGetValues(m_hChart, XRT_YGRID, &val, NULL); return(val); }
    XrtDataStyle * GetYGridDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_YGRID_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetYGridDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YGRID_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetYGridUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YGRID_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR ** GetYLabels() { TCHAR ** val; XrtGetValues(m_hChart, XRT_YLABELS, &val, NULL); return(val); }
    double GetYMarker() { double val; XrtGetValues(m_hChart, XRT_YMARKER, &val, NULL); return(val); }
    XrtDataStyle * GetYMarkerDataStyle() { XrtDataStyle * val; XrtGetValues(m_hChart, XRT_YMARKER_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetYMarkerDataStyleUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YMARKER_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetYMarkerShow() { BOOL val; XrtGetValues(m_hChart, XRT_YMARKER_SHOW, &val, NULL); return(val); }
    double GetYMax() { double val; XrtGetValues(m_hChart, XRT_YMAX, &val, NULL); return(val); }
    BOOL GetYMaxUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YMAX_USE_DEFAULT, &val, NULL); return(val); }
    double GetYMin() { double val; XrtGetValues(m_hChart, XRT_YMIN, &val, NULL); return(val); }
    BOOL GetYMinUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YMIN_USE_DEFAULT, &val, NULL); return(val); }
    double GetYNum() { double val; XrtGetValues(m_hChart, XRT_YNUM, &val, NULL); return(val); }
    XrtNumMethod GetYNumMethod() { XrtNumMethod val; XrtGetValues(m_hChart, XRT_YNUM_METHOD, &val, NULL); return(val); }
    BOOL GetYNumUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YNUM_USE_DEFAULT, &val, NULL); return(val); }
    double GetYOrigin() { double val; XrtGetValues(m_hChart, XRT_YORIGIN, &val, NULL); return(val); }
    XrtOriginPlacement GetYOriginPlacement() { XrtOriginPlacement val; XrtGetValues(m_hChart, XRT_YORIGIN_PLACEMENT, &val, NULL); return(val); }
    BOOL GetYOriginUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YORIGIN_USE_DEFAULT, &val, NULL); return(val); }
    int GetYPrecision() { int val; XrtGetValues(m_hChart, XRT_YPRECISION, &val, NULL); return(val); }
    BOOL GetYPrecisionUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YPRECISION_USE_DEFAULT, &val, NULL); return(val); }
    double GetYTick() { double val; XrtGetValues(m_hChart, XRT_YTICK, &val, NULL); return(val); }
    BOOL GetYTickUseDefault() { BOOL val; XrtGetValues(m_hChart, XRT_YTICK_USE_DEFAULT, &val, NULL); return(val); }
    TCHAR * GetYTitle() { TCHAR * val; XrtGetValues(m_hChart, XRT_YTITLE, &val, NULL); return(val); }
    XrtRotate GetYTitleRotation() { XrtRotate val; XrtGetValues(m_hChart, XRT_YTITLE_ROTATION, &val, NULL); return(val); }
    long GetOptions() { long val; XrtGetValues(m_hChart, XRT_CHART_OPTIONS, &val, NULL); return(val); }
    HFONT GetAxisTitleFont() { HFONT val; XrtGetValues(m_hChart, XRT_AXIS_TITLE_FONT, &val, NULL); return(val); }

    // Set methods
    void SetAngleUnit(XrtAngleUnit val) { XrtSetValues(m_hChart, XRT_ANGLE_UNIT, val, NULL); }
    void SetAlarmZone(const TCHAR *name, BOOL is_shown, XrtFloat LowerY, XrtFloat UpperY,
                      COLORREF line_color, COLORREF fill_color, XrtFillPattern fill_pattern)
      { XrtSetAlarmZone(m_hChart, name, is_shown, LowerY, UpperY, line_color, fill_color, fill_pattern); }
    void SetAxisBoundingBox(BOOL val) { XrtSetValues(m_hChart, XRT_AXIS_BOUNDING_BOX, val, NULL); }
    void SetAxisFont(HFONT val) { XrtSetValues(m_hChart, XRT_AXIS_FONT, (int)val, NULL); }
    void SetBackgroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_BACKGROUND_COLOR, val, NULL); }
    void SetBarClusterOverlap(int val) { XrtSetValues(m_hChart, XRT_BAR_CLUSTER_OVERLAP, val, NULL); }
    void SetBarClusterWidth(int val) { XrtSetValues(m_hChart, XRT_BAR_CLUSTER_WIDTH, val, NULL); }
    void SetBorder(XrtBorder val) { XrtSetValues(m_hChart, XRT_BORDER, val, NULL); }
    void SetBorderWidth(int val) { XrtSetValues(m_hChart, XRT_BORDER_WIDTH, val, NULL); }
    void SetBubbleMax(double val) { XrtSetValues(m_hChart, XRT_BUBBLE_MAX, val, NULL); }
    void SetBubbleMethod(XrtBubbleMethod val) { XrtSetValues(m_hChart, XRT_BUBBLE_METHOD, val, NULL); }
    void SetBubbleMin(double val) { XrtSetValues(m_hChart, XRT_BUBBLE_MIN, val, NULL); }
    void SetCandleComplex(BOOL val) { XrtSetValues(m_hChart, XRT_CANDLE_COMPLEX, val, NULL); }
	void SetCandleFillFalling(BOOL val) { XrtSetValues(m_hChart, XRT_CANDLE_FILLFALLING, val, NULL); }
    void SetData(XrtDataHandle val) { XrtSetValues(m_hChart, XRT_DATA, val, NULL); }
    void SetData2(XrtDataHandle val) { XrtSetValues(m_hChart, XRT_DATA2, val, NULL); }
    void SetDataAreaBackgroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_DATA_AREA_BACKGROUND_COLOR, val, NULL); }
    void SetDataAreaForegroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_DATA_AREA_FOREGROUND_COLOR, val, NULL); }
    void SetDataAreaImage(XrtImage val) { XrtSetValues(m_hChart, XRT_DATA_AREA_IMAGE, val, NULL); }
    void SetDataAreaImageLayout(XrtImageLayout val) { XrtSetValues(m_hChart, XRT_DATA_AREA_IMAGE_LAYOUT, val, NULL); }
    void SetDataAreaImageObsolete(TCHAR * val) { XrtSetValues(m_hChart, XRT_DATA_AREA_IMAGE_OBSOLETE, val, NULL); }
    void SetDataAreaImageTransparent(BOOL val) { XrtSetValues(m_hChart, XRT_DATA_AREA_IMAGE_TRANSPARENT, val, NULL); }
    void SetDataStyles(XrtDataStyle ** val) { XrtSetValues(m_hChart, XRT_DATA_STYLES, val, NULL); }
    void SetDataStyles2(XrtDataStyle ** val) { XrtSetValues(m_hChart, XRT_DATA_STYLES2, val, NULL); }
    void SetDataStyles2UseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_DATA_STYLES2_USE_DEFAULT, val, NULL); }
    void SetDataStylesUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_DATA_STYLES_USE_DEFAULT, val, NULL); }
    void SetDebug(BOOL val) { XrtSetValues(m_hChart, XRT_DEBUG, val, NULL); }
    void SetDoubleBuffer(BOOL val) { XrtSetValues(m_hChart, XRT_DOUBLE_BUFFER, val, NULL); }
    void SetExtraDefaultDataStyles(BOOL val) { XrtSetValues(m_hChart, XRT_EXTRA_DEFAULT_DATA_STYLES, val, NULL); }
    void SetFooterAdjust(XrtAdjust val) { XrtSetValues(m_hChart, XRT_FOOTER_ADJUST, val, NULL); }
    void SetFooterBackgroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_FOOTER_BACKGROUND_COLOR, val, NULL); }
    void SetFooterBorder(XrtBorder val) { XrtSetValues(m_hChart, XRT_FOOTER_BORDER, val, NULL); }
    void SetFooterBorderWidth(int val) { XrtSetValues(m_hChart, XRT_FOOTER_BORDER_WIDTH, val, NULL); }
    void SetFooterFont(HFONT val) { XrtSetValues(m_hChart, XRT_FOOTER_FONT, (int)val, NULL); }
    void SetFooterForegroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_FOOTER_FOREGROUND_COLOR, val, NULL); }
    void SetFooterHeight(int val) { XrtSetValues(m_hChart, XRT_FOOTER_HEIGHT, val, NULL); }
    void SetFooterImage(XrtImage val) { XrtSetValues(m_hChart, XRT_FOOTER_IMAGE, val, NULL); }
    void SetFooterImageLayout(XrtImageLayout val) { XrtSetValues(m_hChart, XRT_FOOTER_IMAGE_LAYOUT, val, NULL); }
    void SetFooterImageMinimumSize(BOOL val) { XrtSetValues(m_hChart, XRT_FOOTER_IMAGE_MINIMUM_SIZE, val, NULL); }
    void SetFooterImageObsolete(TCHAR * val) { XrtSetValues(m_hChart, XRT_FOOTER_IMAGE_OBSOLETE, val, NULL); }
    void SetFooterImageTransparent(BOOL val) { XrtSetValues(m_hChart, XRT_FOOTER_IMAGE_TRANSPARENT, val, NULL); }
    void SetFooterStrings(TCHAR ** val) { XrtSetValues(m_hChart, XRT_FOOTER_STRINGS, val, NULL); }
    void SetFooterWidth(int val) { XrtSetValues(m_hChart, XRT_FOOTER_WIDTH, val, NULL); }
    void SetFooterX(int val) { XrtSetValues(m_hChart, XRT_FOOTER_X, val, NULL); }
    void SetFooterXUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_FOOTER_X_USE_DEFAULT, val, NULL); }
    void SetFooterY(int val) { XrtSetValues(m_hChart, XRT_FOOTER_Y, val, NULL); }
    void SetFooterYUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_FOOTER_Y_USE_DEFAULT, val, NULL); }
    void SetForegroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_FOREGROUND_COLOR, val, NULL); }
    void SetFrontDataset(int val) { XrtSetValues(m_hChart, XRT_FRONT_DATASET, val, NULL); }
    void SetGraph3dShading(XrtShading val) { XrtSetValues(m_hChart, XRT_GRAPH_3D_SHADING, val, NULL); }
    void SetGraphBackgroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_GRAPH_BACKGROUND_COLOR, val, NULL); }
    void SetGraphBorder(XrtBorder val) { XrtSetValues(m_hChart, XRT_GRAPH_BORDER, val, NULL); }
    void SetGraphBorderWidth(int val) { XrtSetValues(m_hChart, XRT_GRAPH_BORDER_WIDTH, val, NULL); }
    void SetGraphDepth(int val) { XrtSetValues(m_hChart, XRT_GRAPH_DEPTH, val, NULL); }
    void SetGraphForegroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_GRAPH_FOREGROUND_COLOR, val, NULL); }
    void SetGraphHeight(int val) { XrtSetValues(m_hChart, XRT_GRAPH_HEIGHT, val, NULL); }
    void SetGraphHeightUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_HEIGHT_USE_DEFAULT, val, NULL); }
    void SetGraphImage(XrtImage val) { XrtSetValues(m_hChart, XRT_GRAPH_IMAGE, val, NULL); }
    void SetGraphImageLayout(XrtImageLayout val) { XrtSetValues(m_hChart, XRT_GRAPH_IMAGE_LAYOUT, val, NULL); }
    void SetGraphImageObsolete(TCHAR * val) { XrtSetValues(m_hChart, XRT_GRAPH_IMAGE_OBSOLETE, val, NULL); }
    void SetGraphImageTransparent(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_IMAGE_TRANSPARENT, val, NULL); }
    void SetGraphInclination(int val) { XrtSetValues(m_hChart, XRT_GRAPH_INCLINATION, val, NULL); }
    void SetGraphMarginBottom(int val) { XrtSetValues(m_hChart, XRT_GRAPH_MARGIN_BOTTOM, val, NULL); }
    void SetGraphMarginBottomUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_MARGIN_BOTTOM_USE_DEFAULT, val, NULL); }
    void SetGraphMarginLeft(int val) { XrtSetValues(m_hChart, XRT_GRAPH_MARGIN_LEFT, val, NULL); }
    void SetGraphMarginLeftUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_MARGIN_LEFT_USE_DEFAULT, val, NULL); }
    void SetGraphMarginRight(int val) { XrtSetValues(m_hChart, XRT_GRAPH_MARGIN_RIGHT, val, NULL); }
    void SetGraphMarginRightUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_MARGIN_RIGHT_USE_DEFAULT, val, NULL); }
    void SetGraphMarginTop(int val) { XrtSetValues(m_hChart, XRT_GRAPH_MARGIN_TOP, val, NULL); }
    void SetGraphMarginTopUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_MARGIN_TOP_USE_DEFAULT, val, NULL); }
    void SetGraphRotation(int val) { XrtSetValues(m_hChart, XRT_GRAPH_ROTATION, val, NULL); }
    void SetGraphShowOutlines(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_SHOW_OUTLINES, val, NULL); }
    void SetGraphWidth(int val) { XrtSetValues(m_hChart, XRT_GRAPH_WIDTH, val, NULL); }
    void SetGraphWidthUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_WIDTH_USE_DEFAULT, val, NULL); }
    void SetGraphX(int val) { XrtSetValues(m_hChart, XRT_GRAPH_X, val, NULL); }
    void SetGraphXUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_X_USE_DEFAULT, val, NULL); }
    void SetGraphY(int val) { XrtSetValues(m_hChart, XRT_GRAPH_Y, val, NULL); }
    void SetGraphYUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_GRAPH_Y_USE_DEFAULT, val, NULL); }
    void SetHeaderAdjust(XrtAdjust val) { XrtSetValues(m_hChart, XRT_HEADER_ADJUST, val, NULL); }
    void SetHeaderBackgroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_HEADER_BACKGROUND_COLOR, val, NULL); }
    void SetHeaderBorder(XrtBorder val) { XrtSetValues(m_hChart, XRT_HEADER_BORDER, val, NULL); }
    void SetHeaderBorderWidth(int val) { XrtSetValues(m_hChart, XRT_HEADER_BORDER_WIDTH, val, NULL); }
    void SetHeaderFont(HFONT val) { XrtSetValues(m_hChart, XRT_HEADER_FONT, (int)val, NULL); }
    void SetHeaderForegroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_HEADER_FOREGROUND_COLOR, val, NULL); }
    void SetHeaderHeight(int val) { XrtSetValues(m_hChart, XRT_HEADER_HEIGHT, val, NULL); }
    void SetHeaderImage(XrtImage val) { XrtSetValues(m_hChart, XRT_HEADER_IMAGE, val, NULL); }
    void SetHeaderImageLayout(XrtImageLayout val) { XrtSetValues(m_hChart, XRT_HEADER_IMAGE_LAYOUT, val, NULL); }
    void SetHeaderImageMinimumSize(BOOL val) { XrtSetValues(m_hChart, XRT_HEADER_IMAGE_MINIMUM_SIZE, val, NULL); }
    void SetHeaderImageObsolete(TCHAR * val) { XrtSetValues(m_hChart, XRT_HEADER_IMAGE_OBSOLETE, val, NULL); }
    void SetHeaderImageTransparent(BOOL val) { XrtSetValues(m_hChart, XRT_HEADER_IMAGE_TRANSPARENT, val, NULL); }
    void SetHeaderStrings(TCHAR ** val) { XrtSetValues(m_hChart, XRT_HEADER_STRINGS, val, NULL); }
    void SetHeaderWidth(int val) { XrtSetValues(m_hChart, XRT_HEADER_WIDTH, val, NULL); }
    void SetHeaderX(int val) { XrtSetValues(m_hChart, XRT_HEADER_X, val, NULL); }
    void SetHeaderXUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_HEADER_X_USE_DEFAULT, val, NULL); }
    void SetHeaderY(int val) { XrtSetValues(m_hChart, XRT_HEADER_Y, val, NULL); }
    void SetHeaderYUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_HEADER_Y_USE_DEFAULT, val, NULL); }
    void SetHeight(int val) { XrtSetValues(m_hChart, XRT_HEIGHT, val, NULL); }
    void SetHiloCloseShow(BOOL val) { XrtSetValues(m_hChart, XRT_HILO_CLOSE_SHOW, val, NULL); }
    void SetHiloOpenCloseFullWidth(BOOL val) { XrtSetValues(m_hChart, XRT_HILO_OPEN_CLOSE_FULL_WIDTH, val, NULL); }
    void SetHiloOpenShow(BOOL val) { XrtSetValues(m_hChart, XRT_HILO_OPEN_SHOW, val, NULL); }
    void SetImage(XrtImage val) { XrtSetValues(m_hChart, XRT_IMAGE, val, NULL); }
    void SetImageLayout(XrtImageLayout val) { XrtSetValues(m_hChart, XRT_IMAGE_LAYOUT, val, NULL); }
    void SetImageTransparent(BOOL val) { XrtSetValues(m_hChart, XRT_IMAGE_TRANSPARENT, val, NULL); }
    void SetImageObsolete(TCHAR * val) { XrtSetValues(m_hChart, XRT_IMAGE_OBSOLETE, val, NULL); }
    void SetInvertOrientation(BOOL val) { XrtSetValues(m_hChart, XRT_INVERT_ORIENTATION, val, NULL); }
    void SetIsStacked(BOOL val) { XrtSetValues(m_hChart, XRT_IS_STACKED, val, NULL); }
    void SetIsStacked2(BOOL val) { XrtSetValues(m_hChart, XRT_IS_STACKED2, val, NULL); }
    void SetLegendAnchor(XrtAnchor val) { XrtSetValues(m_hChart, XRT_LEGEND_ANCHOR, val, NULL); }
    void SetLegendBackgroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_LEGEND_BACKGROUND_COLOR, val, NULL); }
    void SetLegendBorder(XrtBorder val) { XrtSetValues(m_hChart, XRT_LEGEND_BORDER, val, NULL); }
    void SetLegendBorderWidth(int val) { XrtSetValues(m_hChart, XRT_LEGEND_BORDER_WIDTH, val, NULL); }
    void SetLegendFont(HFONT val) { XrtSetValues(m_hChart, XRT_LEGEND_FONT, (int)val, NULL); }
    void SetLegendForegroundColor(COLORREF val) { XrtSetValues(m_hChart, XRT_LEGEND_FOREGROUND_COLOR, val, NULL); }
    void SetLegendHeight(int val) { XrtSetValues(m_hChart, XRT_LEGEND_HEIGHT, val, NULL); }
    void SetLegendImage(XrtImage val) { XrtSetValues(m_hChart, XRT_LEGEND_IMAGE, val, NULL); }
    void SetLegendImageLayout(XrtImageLayout val) { XrtSetValues(m_hChart, XRT_LEGEND_IMAGE_LAYOUT, val, NULL); }
    void SetLegendImageObsolete(TCHAR * val) { XrtSetValues(m_hChart, XRT_LEGEND_IMAGE_OBSOLETE, val, NULL); }
    void SetLegendImageTransparent(BOOL val) { XrtSetValues(m_hChart, XRT_LEGEND_IMAGE_TRANSPARENT, val, NULL); }
    void SetLegendOrientation(XrtAlign val) { XrtSetValues(m_hChart, XRT_LEGEND_ORIENTATION, val, NULL); }
    void SetLegendReversed(BOOL val) { XrtSetValues(m_hChart, XRT_LEGEND_REVERSED, val, NULL); }
    void SetLegendShow(BOOL val) { XrtSetValues(m_hChart, XRT_LEGEND_SHOW, val, NULL); }
    void SetLegendTitle(TCHAR * val) { XrtSetValues(m_hChart, XRT_LEGEND_TITLE, val, NULL); }
    void SetLegendWidth(int val) { XrtSetValues(m_hChart, XRT_LEGEND_WIDTH, val, NULL); }
    void SetLegendX(int val) { XrtSetValues(m_hChart, XRT_LEGEND_X, val, NULL); }
    void SetLegendXUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_LEGEND_X_USE_DEFAULT, val, NULL); }
    void SetLegendY(int val) { XrtSetValues(m_hChart, XRT_LEGEND_Y, val, NULL); }
    void SetLegendYUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_LEGEND_Y_USE_DEFAULT, val, NULL); }
    void SetMarkerDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_MARKER_DATA_STYLE, val, NULL); }
    void SetMarkerDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_MARKER_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetMarkerDataset(int val) { XrtSetValues(m_hChart, XRT_MARKER_DATASET, val, NULL); }
    void SetName(TCHAR * val) { XrtSetValues(m_hChart, XRT_NAME, val, NULL); }
    void SetOtherDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_OTHER_DATA_STYLE, val, NULL); }
    void SetOtherDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_OTHER_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetOtherLabel(TCHAR * val) { XrtSetValues(m_hChart, XRT_OTHER_LABEL, val, NULL); }
    void SetPieMergeMissingSlices(BOOL val) { XrtSetValues(m_hChart, XRT_PIE_MERGE_MISSING_SLICES, val, NULL); }
    void SetPieMinSlices(int val) { XrtSetValues(m_hChart, XRT_PIE_MIN_SLICES, val, NULL); }
    void SetPieOrder(XrtPieOrder val) { XrtSetValues(m_hChart, XRT_PIE_ORDER, val, NULL); }
    void SetPieStartAngle(double val) { XrtSetValues(m_hChart, XRT_PIE_START_ANGLE, val, NULL); }
    void SetPieThresholdMethod(XrtPieThresholdMethod val) { XrtSetValues(m_hChart, XRT_PIE_THRESHOLD_METHOD, val, NULL); }
    void SetPieThresholdValue(double val) { XrtSetValues(m_hChart, XRT_PIE_THRESHOLD_VALUE, val, NULL); }
    void SetPointLabels(TCHAR ** val) { XrtSetValues(m_hChart, XRT_POINT_LABELS, val, NULL); }
    void SetPointLabels2(TCHAR ** val) { XrtSetValues(m_hChart, XRT_POINT_LABELS2, val, NULL); }
    void SetPolarAxisAllowNegatives(BOOL val) { XrtSetValues(m_hChart, XRT_POLAR_AXIS_ALLOW_NEGATIVES, val, NULL); }
    void SetPolarHalfRange(BOOL val) { XrtSetValues(m_hChart, XRT_POLAR_HALF_RANGE, val, NULL); }
    void SetRepaint(BOOL val) { XrtSetValues(m_hChart, XRT_REPAINT, val, NULL); }
    void SetSetLabels(TCHAR ** val) { XrtSetValues(m_hChart, XRT_SET_LABELS, val, NULL); }
    void SetSetLabels2(TCHAR ** val) { XrtSetValues(m_hChart, XRT_SET_LABELS2, val, NULL); }
    void SetTimeBase(long val) { XrtSetValues(m_hChart, XRT_TIME_BASE, val, NULL); }
    void SetTimeBaseOle(double val) { XrtSetValues(m_hChart, XRT_TIME_BASE_OLE, val, NULL); }
    void SetTimeFormat(TCHAR * val) { XrtSetValues(m_hChart, XRT_TIME_FORMAT, val, NULL); }
    void SetTimeFormatUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_TIME_FORMAT_USE_DEFAULT, val, NULL); }
    void SetTimeUnit(XrtTimeUnit val) { XrtSetValues(m_hChart, XRT_TIME_UNIT, val, NULL); }
    void SetTransposeData(BOOL val) { XrtSetValues(m_hChart, XRT_TRANSPOSE_DATA, val, NULL); }
    void SetType(XrtType val) { XrtSetValues(m_hChart, XRT_TYPE, val, NULL); }
    void SetType2(XrtType val) { XrtSetValues(m_hChart, XRT_TYPE2, val, NULL); }
    void SetWidth(int val) { XrtSetValues(m_hChart, XRT_WIDTH, val, NULL); }
    void SetXAnnoPlacement(XrtAnnoPlacement val) { XrtSetValues(m_hChart, XRT_XANNO_PLACEMENT, val, NULL); }
    void SetXAnnotationMethod(XrtAnnoMethod val) { XrtSetValues(m_hChart, XRT_XANNOTATION_METHOD, val, NULL); }
    void SetXAnnotationRotation(XrtRotate val) { XrtSetValues(m_hChart, XRT_XANNOTATION_ROTATION, val, NULL); }
    void SetXAnnotationRotationAngle(double val) { XrtSetValues(m_hChart, XRT_XANNOTATION_ROTATION_ANGLE, val, NULL); }
    void SetXAxisDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_XAXIS_DATA_STYLE, val, NULL); }
    void SetXAxisDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XAXIS_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetXAxisLabelFormat(TCHAR * val) { XrtSetValues(m_hChart, XRT_XAXIS_LABEL_FORMAT, val, NULL); }
    void SetXAxisLogarithmic(BOOL val) { XrtSetValues(m_hChart, XRT_XAXIS_LOGARITHMIC, val, NULL); }
    void SetXAxisMax(double val) { XrtSetValues(m_hChart, XRT_XAXIS_MAX, val, NULL); }
    void SetXAxisMaxUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XAXIS_MAX_USE_DEFAULT, val, NULL); }
    void SetXAxisMin(double val) { XrtSetValues(m_hChart, XRT_XAXIS_MIN, val, NULL); }
    void SetXAxisMinUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XAXIS_MIN_USE_DEFAULT, val, NULL); }
    void SetXAxisReversed(BOOL val) { XrtSetValues(m_hChart, XRT_XAXIS_REVERSED, val, NULL); }
    void SetXAxisShow(BOOL val) { XrtSetValues(m_hChart, XRT_XAXIS_SHOW, val, NULL); }
    void SetXGrid(double val) { XrtSetValues(m_hChart, XRT_XGRID, val, NULL); }
    void SetXGridDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_XGRID_DATA_STYLE, val, NULL); }
    void SetXGridDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XGRID_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetXGridUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XGRID_USE_DEFAULT, val, NULL); }
    void SetXLabels(TCHAR ** val) { XrtSetValues(m_hChart, XRT_XLABELS, val, NULL); }
    void SetXMarker(double val) { XrtSetValues(m_hChart, XRT_XMARKER, val, NULL); }
    void SetXMarkerDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_XMARKER_DATA_STYLE, val, NULL); }
    void SetXMarkerDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XMARKER_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetXMarkerMethod(XrtMarkerMethod val) { XrtSetValues(m_hChart, XRT_XMARKER_METHOD, val, NULL); }
    void SetXMarkerPoint(int val) { XrtSetValues(m_hChart, XRT_XMARKER_POINT, val, NULL); }
    void SetXMarkerSet(int val) { XrtSetValues(m_hChart, XRT_XMARKER_SET, val, NULL); }
    void SetXMarkerShow(BOOL val) { XrtSetValues(m_hChart, XRT_XMARKER_SHOW, val, NULL); }
    void SetXMax(double val) { XrtSetValues(m_hChart, XRT_XMAX, val, NULL); }
    void SetXMaxUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XMAX_USE_DEFAULT, val, NULL); }
    void SetXMin(double val) { XrtSetValues(m_hChart, XRT_XMIN, val, NULL); }
    void SetXMinUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XMIN_USE_DEFAULT, val, NULL); }
    void SetXNum(double val) { XrtSetValues(m_hChart, XRT_XNUM, val, NULL); }
    void SetXNumMethod(XrtNumMethod val) { XrtSetValues(m_hChart, XRT_XNUM_METHOD, val, NULL); }
    void SetXNumUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XNUM_USE_DEFAULT, val, NULL); }
    void SetXOrigin(double val) { XrtSetValues(m_hChart, XRT_XORIGIN, val, NULL); }
    void SetXOriginBase(double val) { XrtSetValues(m_hChart, XRT_XORIGIN_BASE, val, NULL); }
    void SetXOriginPlacement(XrtOriginPlacement val) { XrtSetValues(m_hChart, XRT_XORIGIN_PLACEMENT, val, NULL); }
    void SetXOriginUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XORIGIN_USE_DEFAULT, val, NULL); }
    void SetXPrecision(int val) { XrtSetValues(m_hChart, XRT_XPRECISION, val, NULL); }
    void SetXPrecisionUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XPRECISION_USE_DEFAULT, val, NULL); }
    void SetXTick(double val) { XrtSetValues(m_hChart, XRT_XTICK, val, NULL); }
    void SetXTickUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_XTICK_USE_DEFAULT, val, NULL); }
    void SetXTitle(TCHAR * val) { XrtSetValues(m_hChart, XRT_XTITLE, val, NULL); }
    void SetXTitleRotation(XrtRotate val) { XrtSetValues(m_hChart, XRT_XTITLE_ROTATION, val, NULL); }
    void SetY2AnnoPlacement(XrtAnnoPlacement val) { XrtSetValues(m_hChart, XRT_Y2ANNO_PLACEMENT, val, NULL); }
    void SetY2AnnotationAngle(double val) { XrtSetValues(m_hChart, XRT_Y2ANNOTATION_ANGLE, val, NULL); }
    void SetY2AnnotationAngleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2ANNOTATION_ANGLE_USE_DEFAULT, val, NULL); }
    void SetY2AnnotationMethod(XrtAnnoMethod val) { XrtSetValues(m_hChart, XRT_Y2ANNOTATION_METHOD, val, NULL); }
    void SetY2AnnotationRotation(XrtRotate val) { XrtSetValues(m_hChart, XRT_Y2ANNOTATION_ROTATION, val, NULL); }
    void SetY2AnnotationRotationAngle(double val) { XrtSetValues(m_hChart, XRT_Y2ANNOTATION_ROTATION_ANGLE, val, NULL); }
    void SetY2Axis100Percent(BOOL val) { XrtSetValues(m_hChart, XRT_Y2AXIS_100_PERCENT, val, NULL); }
    void SetY2AxisDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_Y2AXIS_DATA_STYLE, val, NULL); }
    void SetY2AxisDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2AXIS_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetY2AxisLabelFormat(TCHAR * val) { XrtSetValues(m_hChart, XRT_Y2AXIS_LABEL_FORMAT, val, NULL); }
    void SetY2AxisLogarithmic(BOOL val) { XrtSetValues(m_hChart, XRT_Y2AXIS_LOGARITHMIC, val, NULL); }
    void SetY2AxisMax(double val) { XrtSetValues(m_hChart, XRT_Y2AXIS_MAX, val, NULL); }
    void SetY2AxisMaxUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2AXIS_MAX_USE_DEFAULT, val, NULL); }
    void SetY2AxisMin(double val) { XrtSetValues(m_hChart, XRT_Y2AXIS_MIN, val, NULL); }
    void SetY2AxisMinUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2AXIS_MIN_USE_DEFAULT, val, NULL); }
    void SetY2AxisReversed(BOOL val) { XrtSetValues(m_hChart, XRT_Y2AXIS_REVERSED, val, NULL); }
    void SetY2AxisShow(BOOL val) { XrtSetValues(m_hChart, XRT_Y2AXIS_SHOW, val, NULL); }
    void SetY2Grid(double val) { XrtSetValues(m_hChart, XRT_Y2GRID, val, NULL); }
    void SetY2GridDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_Y2GRID_DATA_STYLE, val, NULL); }
    void SetY2GridDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2GRID_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetY2GridUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2GRID_USE_DEFAULT, val, NULL); }
    void SetY2Labels(TCHAR ** val) { XrtSetValues(m_hChart, XRT_Y2LABELS, val, NULL); }
    void SetY2Max(double val) { XrtSetValues(m_hChart, XRT_Y2MAX, val, NULL); }
    void SetY2MaxUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2MAX_USE_DEFAULT, val, NULL); }
    void SetY2Min(double val) { XrtSetValues(m_hChart, XRT_Y2MIN, val, NULL); }
    void SetY2MinUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2MIN_USE_DEFAULT, val, NULL); }
    void SetY2Num(double val) { XrtSetValues(m_hChart, XRT_Y2NUM, val, NULL); }
    void SetY2NumMethod(XrtNumMethod val) { XrtSetValues(m_hChart, XRT_Y2NUM_METHOD, val, NULL); }
    void SetY2NumUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2NUM_USE_DEFAULT, val, NULL); }
    void SetY2Origin(double val) { XrtSetValues(m_hChart, XRT_Y2ORIGIN, val, NULL); }
    void SetY2OriginPlacement(XrtOriginPlacement val) { XrtSetValues(m_hChart, XRT_Y2ORIGIN_PLACEMENT, val, NULL); }
    void SetY2OriginUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2ORIGIN_USE_DEFAULT, val, NULL); }
    void SetY2Precision(int val) { XrtSetValues(m_hChart, XRT_Y2PRECISION, val, NULL); }
    void SetY2PrecisionUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2PRECISION_USE_DEFAULT, val, NULL); }
    void SetY2Tick(double val) { XrtSetValues(m_hChart, XRT_Y2TICK, val, NULL); }
    void SetY2TickUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_Y2TICK_USE_DEFAULT, val, NULL); }
    void SetY2Title(TCHAR * val) { XrtSetValues(m_hChart, XRT_Y2TITLE, val, NULL); }
    void SetY2TitleRotation(XrtRotate val) { XrtSetValues(m_hChart, XRT_Y2TITLE_ROTATION, val, NULL); }
    void SetYAnnoPlacement(XrtAnnoPlacement val) { XrtSetValues(m_hChart, XRT_YANNO_PLACEMENT, val, NULL); }
    void SetYAnnotationAngle(double val) { XrtSetValues(m_hChart, XRT_YANNOTATION_ANGLE, val, NULL); }
    void SetYAnnotationAngleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YANNOTATION_ANGLE_USE_DEFAULT, val, NULL); }
    void SetYAnnotationMethod(XrtAnnoMethod val) { XrtSetValues(m_hChart, XRT_YANNOTATION_METHOD, val, NULL); }
    void SetYAnnotationRotation(XrtRotate val) { XrtSetValues(m_hChart, XRT_YANNOTATION_ROTATION, val, NULL); }
    void SetYAnnotationRotationAngle(double val) { XrtSetValues(m_hChart, XRT_YANNOTATION_ROTATION_ANGLE, val, NULL); }
    void SetYAxis100Percent(BOOL val) { XrtSetValues(m_hChart, XRT_YAXIS_100_PERCENT, val, NULL); }
    void SetYAxisConst(double val) { XrtSetValues(m_hChart, XRT_YAXIS_CONST, val, NULL); }
    void SetYAxisDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_YAXIS_DATA_STYLE, val, NULL); }
    void SetYAxisDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YAXIS_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetYAxisLabelFormat(TCHAR * val) { XrtSetValues(m_hChart, XRT_YAXIS_LABEL_FORMAT, val, NULL); }
    void SetYAxisLogarithmic(BOOL val) { XrtSetValues(m_hChart, XRT_YAXIS_LOGARITHMIC, val, NULL); }
    void SetYAxisMax(double val) { XrtSetValues(m_hChart, XRT_YAXIS_MAX, val, NULL); }
    void SetYAxisMaxUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YAXIS_MAX_USE_DEFAULT, val, NULL); }
    void SetYAxisMin(double val) { XrtSetValues(m_hChart, XRT_YAXIS_MIN, val, NULL); }
    void SetYAxisMinUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YAXIS_MIN_USE_DEFAULT, val, NULL); }
    void SetYAxisMult(double val) { XrtSetValues(m_hChart, XRT_YAXIS_MULT, val, NULL); }
    void SetYAxisReversed(BOOL val) { XrtSetValues(m_hChart, XRT_YAXIS_REVERSED, val, NULL); }
    void SetYAxisShow(BOOL val) { XrtSetValues(m_hChart, XRT_YAXIS_SHOW, val, NULL); }
    void SetYGrid(double val) { XrtSetValues(m_hChart, XRT_YGRID, val, NULL); }
    void SetYGridDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_YGRID_DATA_STYLE, val, NULL); }
    void SetYGridDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YGRID_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetYGridUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YGRID_USE_DEFAULT, val, NULL); }
    void SetYLabels(TCHAR ** val) { XrtSetValues(m_hChart, XRT_YLABELS, val, NULL); }
    void SetYMarker(double val) { XrtSetValues(m_hChart, XRT_YMARKER, val, NULL); }
    void SetYMarkerDataStyle(XrtDataStyle * val) { XrtSetValues(m_hChart, XRT_YMARKER_DATA_STYLE, val, NULL); }
    void SetYMarkerDataStyleUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YMARKER_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetYMarkerShow(BOOL val) { XrtSetValues(m_hChart, XRT_YMARKER_SHOW, val, NULL); }
    void SetYMax(double val) { XrtSetValues(m_hChart, XRT_YMAX, val, NULL); }
    void SetYMaxUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YMAX_USE_DEFAULT, val, NULL); }
    void SetYMin(double val) { XrtSetValues(m_hChart, XRT_YMIN, val, NULL); }
    void SetYMinUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YMIN_USE_DEFAULT, val, NULL); }
    void SetYNum(double val) { XrtSetValues(m_hChart, XRT_YNUM, val, NULL); }
    void SetYNumMethod(XrtNumMethod val) { XrtSetValues(m_hChart, XRT_YNUM_METHOD, val, NULL); }
    void SetYNumUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YNUM_USE_DEFAULT, val, NULL); }
    void SetYOrigin(double val) { XrtSetValues(m_hChart, XRT_YORIGIN, val, NULL); }
    void SetYOriginPlacement(XrtOriginPlacement val) { XrtSetValues(m_hChart, XRT_YORIGIN_PLACEMENT, val, NULL); }
    void SetYOriginUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YORIGIN_USE_DEFAULT, val, NULL); }
    void SetYPrecision(int val) { XrtSetValues(m_hChart, XRT_YPRECISION, val, NULL); }
    void SetYPrecisionUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YPRECISION_USE_DEFAULT, val, NULL); }
    void SetYTick(double val) { XrtSetValues(m_hChart, XRT_YTICK, val, NULL); }
    void SetYTickUseDefault(BOOL val) { XrtSetValues(m_hChart, XRT_YTICK_USE_DEFAULT, val, NULL); }
    void SetYTitle(TCHAR * val) { XrtSetValues(m_hChart, XRT_YTITLE, val, NULL); }
    void SetYTitleRotation(XrtRotate val) { XrtSetValues(m_hChart, XRT_YTITLE_ROTATION, val, NULL); }
    void SetOptions(long val) { XrtSetValues(m_hChart, XRT_CHART_OPTIONS, val, NULL); }
    void SetAxisTitleFont(HFONT val) { XrtSetValues(m_hChart, XRT_AXIS_TITLE_FONT, (int)val, NULL); }

    // Operations
    int ArrCheckAxisBounds(int dset, int npoints) { return XrtArrCheckAxisBounds(m_hChart, dset, npoints); }
    int ArrDataFastUpdate(int dset, int npoints) { return XrtArrDataFastUpdate(m_hChart, dset, npoints); }
    void ClearAlternateDataStyles(int dset) { XrtClearAlternateDataStyles(m_hChart, dset); }
    BOOL DrawToClipboard(XrtDrawFormat format = XRT_DRAW_BITMAP) { return XrtDrawToClipboard(m_hChart, format); }
    BOOL DrawToDC(HDC hdc, XrtDrawFormat format = XRT_DRAW_ENHMETAFILE, XrtDrawScale scale = XRT_DRAWSCALE_TOFIT, int left = 0, int top = 0, int width = 0, int height = 0) { return XrtDrawToDC(m_hChart, hdc, format, scale, left, top, width, height); }
    BOOL DrawToFile(TCHAR *file, XrtDrawFormat format = XRT_DRAW_BITMAP) { return XrtDrawToFile(m_hChart, file, format); }
    XrtDataStyle **DupDataStyles(XrtDataStyle **ds) { return XrtDupDataStyles(ds); }
    void FreeDataStyles(XrtDataStyle **ds) { XrtFreeDataStyles(ds); }
    int GenDataFastUpdate(int dset, int set, int npoints) { return XrtGenDataFastUpdate(m_hChart, dset, set, npoints); }
    int GenCheckAxisBounds(int dset, int set, int npoints) { return XrtGenCheckAxisBounds(m_hChart, dset, set, npoints); }
    XrtDataStyle *GetAlternateDataStyle(int dset, int set, int point) { return XrtGetAlternateDataStyle(m_hChart, dset, set, point); }
    XrtAlternateDataStyle **GetAlternateDataStyleList(int dset) { return XrtGetAlternateDataStyleList(m_hChart, dset); }
    XrtDataStyle *GetNthDataStyle(int index) { return XrtGetNthDataStyle(m_hChart, index); }
    XrtDataStyle *GetNthDataStyle2(int index) { return XrtGetNthDataStyle2(m_hChart, index); }
    TCHAR *GetNthPointLabel(int index) { return XrtGetNthPointLabel(m_hChart, index); }
    TCHAR *GetNthPointLabel2(int index) { return XrtGetNthPointLabel2(m_hChart, index); }
    TCHAR *GetNthSetLabel(int index) { return XrtGetNthSetLabel(m_hChart, index); }
    TCHAR *GetNthSetLabel2(int index) { return XrtGetNthSetLabel2(m_hChart, index); }
    BOOL GetPropString(XrtProperty res, TCHAR **str) { return XrtGetPropString(m_hChart, res, str); }
    XrtValueLabel *GetValueLabel(XrtAxis axis, XrtValueLabel *vlabel) { return XrtGetValueLabel(m_hChart, axis, vlabel); }
    XrtRegion Map(int axis, int x, int y, XrtMapResult *map) { return XrtMap(m_hChart, axis, x, y, map); }
    XrtRegion Pick(int ds, int x, int y, XrtPickResult *pick, XrtFocus focus) { return XrtPick(m_hChart, ds, x, y, pick, focus); }
    BOOL Print(XrtDrawFormat format = XRT_DRAW_ENHMETAFILE, XrtDrawScale scale = XRT_DRAWSCALE_TOFIT, int left = 0, int top = 0, int width = 0, int height = 0) { return XrtPrint(m_hChart, format, scale, left, top, width, height); }
    void Reinitialize() {XrtReinitialize(m_hChart);}
    void RemoveAlarmZone(const TCHAR *name) { XrtRemoveAlarmZone(m_hChart, name); }
    void RemoveAllAlarmZones() { XrtRemoveAllAlarmZones(m_hChart); }
    HANDLE RenderClipboardFormat(int cf) { return XrtRenderClipboardFormat(m_hChart, cf); }
    BOOL SaveImageAsJpeg(TCHAR *filename, int quality = 80, BOOL grayscale = FALSE, BOOL optimize = FALSE, BOOL progressive = FALSE) { return XrtSaveImageAsJpeg(m_hChart, filename, quality, grayscale, optimize, progressive); }
    BOOL SaveImageAsPng(TCHAR *filename, BOOL interlace = FALSE) { return XrtSaveImageAsPng(m_hChart, filename, interlace); }
    void SetAlternateDataStyle(int dset, int set, int point, XrtDataStyle *ds) { XrtSetAlternateDataStyle(m_hChart, dset, set, point, ds); }
    void SetNthDataStyle(int index, XrtDataStyle *ds) { XrtSetNthDataStyle(m_hChart, index, ds); }
    void SetNthDataStyle2(int index, XrtDataStyle *ds) { XrtSetNthDataStyle2(m_hChart, index, ds); }
    void SetNthSetLabel(int index, TCHAR *string) { XrtSetNthSetLabel(m_hChart, index, string); }
    void SetNthSetLabel2(int index, TCHAR *string) { XrtSetNthSetLabel2(m_hChart, index, string); }
    void SetNthPointLabel(int index, TCHAR *string) { XrtSetNthPointLabel(m_hChart, index, string); }
    void SetNthPointLabel2(int index, TCHAR *string) { XrtSetNthPointLabel2(m_hChart, index, string); }
    BOOL SetPropString(XrtProperty res, TCHAR *str) { return XrtSetPropString(m_hChart, res, str); }
    void SetValueLabel(XrtAxis axis, XrtValueLabel *vlabel) { XrtSetValueLabel(m_hChart, axis, vlabel); }
    void Unmap(int axis, double x, double y, XrtMapResult *map) { XrtUnmap(m_hChart, axis, x, y, map); }
    void Unpick(int ds, int set, int point, XrtPickResult *pick) { XrtUnpick(m_hChart, ds, set, point, pick); }

protected:
    // Implementation
    virtual BOOL OnChildNotify(UINT, WPARAM, LPARAM, LRESULT*) { return(FALSE); }
};


/*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 *
 *  Class CChart2DTextArea
 *
 *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
*/

class CChart2DTextArea {
public:
    // Constructor
    CChart2DTextArea(CChart2D *chart) { m_hText = XrtTextAreaCreate(chart->m_hChart); }

    // Destructor
    ~CChart2DTextArea() { XrtTextAreaDestroy(m_hText); }

    // Get methods
    XrtAdjust GetAdjust() { XrtAdjust val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ADJUST, &val, NULL); return(val); }
    XrtAnchor GetAnchor() { XrtAnchor val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ANCHOR, &val, NULL); return(val); }
    int GetAttachDataset() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ATTACH_DATASET, &val, NULL); return(val); }
    int GetAttachPixelX() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ATTACH_PIXEL_X, &val, NULL); return(val); }
    int GetAttachPixelY() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ATTACH_PIXEL_Y, &val, NULL); return(val); }
    int GetAttachPoint() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ATTACH_POINT, &val, NULL); return(val); }
    int GetAttachSet() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ATTACH_SET, &val, NULL); return(val); }
    XrtAttachType GetAttachType() { XrtAttachType val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ATTACH_TYPE, &val, NULL); return(val); }
    double GetAttachValueX() { double val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ATTACH_VALUE_X, &val, NULL); return(val); }
    double GetAttachValueY() { double val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ATTACH_VALUE_Y, &val, NULL); return(val); }
    COLORREF GetBackgroundColor() { COLORREF val; XrtTextAreaGetValues(m_hText, XRT_TEXT_BACKGROUND_COLOR, &val, NULL); return(val); }
    XrtBorder GetBorder() { XrtBorder val; XrtTextAreaGetValues(m_hText, XRT_TEXT_BORDER, &val, NULL); return(val); }
    int GetBorderWidth() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_BORDER_WIDTH, &val, NULL); return(val); }
    HFONT GetFont() { HFONT val; XrtTextAreaGetValues(m_hText, XRT_TEXT_FONT, &val, NULL); return(val); }
    COLORREF GetForegroundColor() { COLORREF val; XrtTextAreaGetValues(m_hText, XRT_TEXT_FOREGROUND_COLOR, &val, NULL); return(val); }
    int GetHeight() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_HEIGHT, &val, NULL); return(val); }
    XrtImage GetImage() { XrtImage val; XrtTextAreaGetValues(m_hText, XRT_TEXT_IMAGE, &val, NULL); return(val); }
    XrtImageLayout GetImageLayout() { XrtImageLayout val; XrtTextAreaGetValues(m_hText, XRT_TEXT_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetImageTransparent() { BOOL val; XrtTextAreaGetValues(m_hText, XRT_TEXT_IMAGE_TRANSPARENT, &val, NULL); return(val); }
    BOOL GetImageMinimumSize() { BOOL val; XrtTextAreaGetValues(m_hText, XRT_TEXT_IMAGE_MINIMUM_SIZE, &val, NULL); return(val); }
    TCHAR * GetImageObsolete() { TCHAR * val; XrtTextAreaGetValues(m_hText, XRT_TEXT_IMAGE_OBSOLETE, &val, NULL); return(val); }
    BOOL GetIsConnected() { BOOL val; XrtTextAreaGetValues(m_hText, XRT_TEXT_IS_CONNECTED, &val, NULL); return(val); }
    BOOL GetIsShowing() { BOOL val; XrtTextAreaGetValues(m_hText, XRT_TEXT_IS_SHOWING, &val, NULL); return(val); }
    TCHAR * GetName() { TCHAR * val; XrtTextAreaGetValues(m_hText, XRT_TEXT_NAME, &val, NULL); return(val); }
    int GetOffset() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_OFFSET, &val, NULL); return(val); }
    XrtRotate GetRotation() { XrtRotate val; XrtTextAreaGetValues(m_hText, XRT_TEXT_ROTATION, &val, NULL); return(val); }
    TCHAR ** GetStrings() { TCHAR ** val; XrtTextAreaGetValues(m_hText, XRT_TEXT_STRINGS, &val, NULL); return(val); }
    int GetWidth() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_WIDTH, &val, NULL); return(val); }
    int GetX() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_X, &val, NULL); return(val); }
    int GetY() { int val; XrtTextAreaGetValues(m_hText, XRT_TEXT_Y, &val, NULL); return(val); }

    // Set methods
    void SetAdjust(XrtAdjust val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ADJUST, val, NULL); }
    void SetAnchor(XrtAnchor val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ANCHOR, val, NULL); }
    void SetAttachDataset(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ATTACH_DATASET, val, NULL); }
    void SetAttachPixelX(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ATTACH_PIXEL_X, val, NULL); }
    void SetAttachPixelY(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ATTACH_PIXEL_Y, val, NULL); }
    void SetAttachPoint(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ATTACH_POINT, val, NULL); }
    void SetAttachSet(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ATTACH_SET, val, NULL); }
    void SetAttachType(XrtAttachType val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ATTACH_TYPE, val, NULL); }
    void SetAttachValueX(double val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ATTACH_VALUE_X, val, NULL); }
    void SetAttachValueY(double val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ATTACH_VALUE_Y, val, NULL); }
    void SetBackgroundColor(COLORREF val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_BACKGROUND_COLOR, val, NULL); }
    void SetBorder(XrtBorder val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_BORDER, val, NULL); }
    void SetBorderWidth(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_BORDER_WIDTH, val, NULL); }
    void SetFont(HFONT val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_FONT, (int)val, NULL); }
    void SetForegroundColor(COLORREF val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_FOREGROUND_COLOR, val, NULL); }
    void SetHeight(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_HEIGHT, val, NULL); }
    void SetImage(XrtImage val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_IMAGE, val, NULL); }
    void SetImageLayout(XrtImageLayout val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_IMAGE_LAYOUT, val, NULL); }
    void SetImageTransparent(BOOL val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_IMAGE_TRANSPARENT, val, NULL); }
    void SetImageMinimumSize(BOOL val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_IMAGE_MINIMUM_SIZE, val, NULL); }
    void SetImageObsolete(TCHAR * val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_IMAGE_OBSOLETE, val, NULL); }
    void SetIsConnected(BOOL val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_IS_CONNECTED, val, NULL); }
    void SetIsShowing(BOOL val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_IS_SHOWING, val, NULL); }
    void SetName(TCHAR * val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_NAME, val, NULL); }
    void SetOffset(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_OFFSET, val, NULL); }
    void SetRotation(XrtRotate val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_ROTATION, val, NULL); }
    void SetStrings(TCHAR ** val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_STRINGS, val, NULL); }
    void SetWidth(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_WIDTH, val, NULL); }
    void SetX(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_X, val, NULL); }
    void SetY(int val) { XrtTextAreaSetValues(m_hText, XRT_TEXT_Y, val, NULL); }

private:
    XrtTextHandle m_hText;
};


/*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 *
 *  Class CChart2DPointStyle
 *
 *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
*/

class CChart2DPointStyle {
public:
    // Constructor
    CChart2DPointStyle(CChart2D *chart) { m_hStyle = XrtPointStyleCreate(chart->m_hChart); }
    CChart2DPointStyle(CChart2D *chart, int dataset, int set, int point)
    {
        m_hStyle = XrtPointStyleCreate(chart->m_hChart);
        XrtPointStyleSetValues(m_hStyle,
            XRT_POINTSTYLE_DATASET, dataset,
            XRT_POINTSTYLE_SET, set,
            XRT_POINTSTYLE_POINT, point,
            NULL);
    }

    // Destructor
    ~CChart2DPointStyle() { XrtPointStyleDestroy(m_hStyle); }

    // Get methods
    XrtDataStyle * GetDataStyle() { XrtDataStyle * val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_DATA_STYLE, &val, NULL); return(val); }
    BOOL GetDataStyleUseDefault() { BOOL val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_DATA_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    int GetDataset() { int val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_DATASET, &val, NULL); return(val); }
    XrtDisplay GetDisplay() { XrtDisplay val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_DISPLAY, &val, NULL); return(val); }
    XrtDataStyle * GetFillStyle() { XrtDataStyle * val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_FILL_STYLE, &val, NULL); return(val); }
    BOOL GetFillStyleUseDefault() { BOOL val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_FILL_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    XrtDataStyle * GetLineStyle() { XrtDataStyle * val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_LINE_STYLE, &val, NULL); return(val); }
    BOOL GetLineStyleUseDefault() { BOOL val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_LINE_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    COLORREF GetPatternBackgroundColor() { COLORREF val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_PATTERN_BACKGROUND_COLOR, &val, NULL); return(val); }
    int GetPoint() { int val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_POINT, &val, NULL); return(val); }
    int GetSet() { int val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_SET, &val, NULL); return(val); }
    double GetSliceOffset() { double val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_SLICE_OFFSET, &val, NULL); return(val); }
    BOOL GetSliceOffsetUseDefault() { BOOL val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_SLICE_OFFSET_USE_DEFAULT, &val, NULL); return(val); }
    XrtDataStyle * GetSymbolStyle() { XrtDataStyle * val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_SYMBOL_STYLE, &val, NULL); return(val); }
    BOOL GetSymbolStyleUseDefault() { BOOL val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_SYMBOL_STYLE_USE_DEFAULT, &val, NULL); return(val); }
    BOOL GetUseDefault() { BOOL val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_USE_DEFAULT, &val, NULL); return(val); }
    XrtImage GetImage() { XrtImage val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_IMAGE, &val, NULL); return(val); }
    XrtImageLayout GetImageLayout() { XrtImageLayout val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_IMAGE_LAYOUT, &val, NULL); return(val); }
    BOOL GetImageTransparent() { BOOL val; XrtPointStyleGetValues(m_hStyle, XRT_POINTSTYLE_IMAGE_TRANSPARENT, &val, NULL); return(val); }

    // Set methods
    void SetDataStyle(XrtDataStyle * val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_DATA_STYLE, val, NULL); }
    void SetDataStyleUseDefault(BOOL val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_DATA_STYLE_USE_DEFAULT, val, NULL); }
    void SetDataset(int val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_DATASET, val, NULL); }
    void SetDisplay(XrtDisplay val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_DISPLAY, val, NULL); }
    void SetFillStyle(XrtDataStyle * val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_FILL_STYLE, val, NULL); }
    void SetFillStyleUseDefault(BOOL val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_FILL_STYLE_USE_DEFAULT, val, NULL); }
    void SetLineStyle(XrtDataStyle * val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_LINE_STYLE, val, NULL); }
    void SetLineStyleUseDefault(BOOL val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_LINE_STYLE_USE_DEFAULT, val, NULL); }
    void SetPatternBackgroundColor(COLORREF val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_PATTERN_BACKGROUND_COLOR, val, NULL); }
    void SetPoint(int val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_POINT, val, NULL); }
    void SetSet(int val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_SET, val, NULL); }
    void SetSliceOffset(double val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_SLICE_OFFSET, val, NULL); }
    void SetSliceOffsetUseDefault(BOOL val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_SLICE_OFFSET_USE_DEFAULT, val, NULL); }
    void SetSymbolStyle(XrtDataStyle * val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_SYMBOL_STYLE, val, NULL); }
    void SetSymbolStyleUseDefault(BOOL val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_SYMBOL_STYLE_USE_DEFAULT, val, NULL); }
    void SetUseDefault(BOOL val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_USE_DEFAULT, val, NULL); }
    void SetImage(XrtImage val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_IMAGE, val, NULL); }
    void SetImageLayout(XrtImageLayout val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_IMAGE_LAYOUT, val, NULL); }
    void SetImageTransparent(BOOL val) { XrtPointStyleSetValues(m_hStyle, XRT_POINTSTYLE_IMAGE_TRANSPARENT, val, NULL); }

private:
    XrtPointStyleHandle m_hStyle;
};


/*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 *
 *  Class CChart2DData
 *
 *-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
*/

class CChart2DData {
public:
    // Constructors
    CChart2DData(XrtDataType type, int nsets, int npoints) { m_hData = XrtDataCreate(type, nsets, npoints); }
    CChart2DData(TCHAR* name, TCHAR* error = NULL) { m_hData = XrtDataCreateFromFile(name, error); }
    CChart2DData(const CChart2DData& from) { m_hData = XrtDataCopy(from.m_hData); }
    CChart2DData(XrtDataHandle handle) { m_hData = handle; }

    // Destructor
    ~CChart2DData() { XrtDataDestroy(m_hData); }

    // Conversion to handle
    operator XrtDataHandle() { return m_hData; }

    // Get methods
    XrtDisplay GetDisplay(int set) { return XrtDataGetDisplay(m_hData, set); }
    int GetFirstPoint(int set) { return XrtDataGetFirstPoint(m_hData, set); }
    int GetFirstSet() { return XrtDataGetFirstSet(m_hData); }
    double GetHole() { double value; XrtDataGetHoleIndirect(m_hData, &value); return value; }
    int GetLastPoint(int set) { return XrtDataGetLastPoint(m_hData, set); }
    int GetLastSet() { return XrtDataGetLastSet(m_hData); }
    int GetNPoints(int set) { return XrtDataGetNPoints(m_hData, set); }
    int GetNSets() { return XrtDataGetNSets(m_hData); }
    XrtDataType GetType() { return XrtDataGetType(m_hData); }
    double* GetXData(int set) { return XrtDataGetXData(m_hData, set); }
    double GetXElement(int set, int point) { double value; XrtDataGetXElementIndirect(m_hData, set, point, &value); return value; }
    double* GetYData(int set) { return XrtDataGetYData(m_hData, set); }
    double GetYElement(int set, int point) { double value; XrtDataGetYElementIndirect(m_hData, set, point, &value); return value; }
    double GetYMean(int set) { return XrtDataGetYMean(m_hData, set); }
    double GetYMedian(int set) { return XrtDataGetYMedian(m_hData, set); }
    double GetYStdDev(int set) { return XrtDataGetYStdDev(m_hData, set); }
    double GetYAveDev(int set) { return XrtDataGetYAveDev(m_hData, set); }
    long GetYDataCount(int set) { return XrtDataGetYDataCount(m_hData, set); }
    double GetYDataMax(int set) { return XrtDataGetYDataMax(m_hData, set); }
    double GetYDataMin(int set) { return XrtDataGetYDataMin(m_hData, set); }
    BOOL GetSeriesLeastSquaresPoly(int set, int order, double ** coefficents) { return XrtDataGetSeriesLeastSquaresPoly(m_hData, set, order, coefficents); }
    BOOL GetLeastSquaresPoly(const double * xv, const double * yv, int npoints, int order, double ** coefficents) { return XrtDataGetLeastSquaresPoly(m_hData, xv, yv, npoints, order, coefficents); }
    double GetPolynomialValue(double xvalue, int order, double *coefficents) { return XrtDataPolynomialEvaluate(xvalue, order, coefficents); }

    // Set methods
    int SetDisplay(int set, XrtDisplay value) { return XrtDataSetDisplay(m_hData, set, value); }
    int SetFirstPoint(int set, int value) { return XrtDataSetFirstPoint(m_hData, set, value); }
    int SetFirstSet(int value) { return XrtDataSetFirstSet(m_hData, value); }
    int SetHole(double value) { return XrtDataSetHole(m_hData, value); }
    int SetLastPoint(int set, int value) { return XrtDataSetLastPoint(m_hData, set, value); }
    int SetLastSet( int value) { return XrtDataSetLastSet(m_hData, value); }
    int SetNPoints(int set, int value) { return XrtDataSetNPoints(m_hData, set, value); }
    int SetNSets(int value) { return XrtDataSetNSets(m_hData, value); }
    int SetType(XrtDataType value) { return XrtDataSetType(m_hData, value); }
    int SetXData(int set, double* values, int num, int start = 0) { return XrtDataSetXData(m_hData, set, values, num, start); }
    int SetXElement(int set, int point, double value) { return XrtDataSetXElement(m_hData, set, point, value); }
    int SetYData(int set, double* values, int num, int start = 0) { return XrtDataSetYData(m_hData, set, values, num, start); }
    int SetYElement(int set, int point, double value) { return XrtDataSetYElement(m_hData, set, point, value); }

    // Operations
    CChart2DData& operator =(const CChart2DData& from)
        {
            XrtDataHandle handle;
            if(this != &from)
            {
                handle = XrtDataCopy(from.m_hData);
                if(handle)
                {
                    XrtDataDestroy(m_hData);
                    m_hData = handle;
                }
            }
            return *this;
        }
    CChart2DData operator +(const CChart2DData& arg2) { return XrtDataConcat(m_hData, arg2.m_hData); }
    CChart2DData ExtractSet(int set) { return XrtDataExtractSet(m_hData, set); }
    int Release() { return XrtDataRelease(m_hData); }
    int RemoveSet(int set) { return XrtDataRemoveSet(m_hData, set); }
    int SaveToFile(TCHAR* name, TCHAR* error = NULL) { return XrtDataSaveToFile(m_hData, name, error); }
    int Sort() { return XrtDataSort(m_hData); }
    CChart2DData Transpose() { return XrtDataTranspose(m_hData); }

private:
    XrtDataHandle m_hData;
};

#endif
