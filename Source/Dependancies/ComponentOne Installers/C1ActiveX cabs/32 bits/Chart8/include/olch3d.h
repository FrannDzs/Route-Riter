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

#ifndef _XRT3D_WIN_H
#define _XRT3D_WIN_H

/* Map some defines to expected defines. */
#if (defined(_Windows) || defined(__WINDOWS__))
# if !defined(_WINDOWS)
#  define _WINDOWS
# endif
#endif

#if (defined(__WIN32__) || defined(__NT__))
# if !defined( _WIN32 )
#  define _WIN32
# endif
#endif

#if defined(_WINDOWS) || defined(RC_INVOKED)
# define WINDOWS        1
#endif

#include <tchar.h>
#include "olch3dcm.h"

#if defined(__cplusplus)
extern "C" {
#endif

XRTIMP(void)                    Xrt3dAboutDialogBox(HWND);

XRTIMP(HXRT3D)                  Xrt3dCreate(void);
XRTIMP(HXRT3D)                  Xrt3dCreateWindow(LPCTSTR, int, int, int, int, HWND, HANDLE);
XRTIMP(BOOL)                    Xrt3dAttachWindow(HXRT3D, HWND);
XRTIMP(HWND)                    Xrt3dDetachWindow(HXRT3D);
XRTIMP(HWND)                    Xrt3dGetWindow(HXRT3D);
XRTIMP(HXRT3D)                  Xrt3dGetHandle(HWND);
XRTIMP(void)                    Xrt3dReinitialize(HXRT3D);

XRTIMP(void)                    Xrt3dComputePalette(HXRT3D);
XRTIMP(HPALETTE)                Xrt3dGetPalette(HXRT3D);

XRTIMP(BOOL)                    Xrt3dDrawToDC(HXRT3D, HDC, Xrt3dDrawFormat, Xrt3dDrawScale, 
                                              int, int, int, int);
XRTIMP(BOOL)                    Xrt3dPrint(HXRT3D, Xrt3dDrawFormat, Xrt3dDrawScale, 
                                           int, int, int, int);
XRTIMP(BOOL)                    Xrt3dDrawToClipboard(HXRT3D, Xrt3dDrawFormat);
XRTIMP(BOOL)                    Xrt3dDrawToFile(HXRT3D, TCHAR *, Xrt3dDrawFormat);
XRTIMP(BOOL)                    Xrt3dSaveImageAsJpeg(HXRT3D, TCHAR *, int, BOOL, BOOL, BOOL);
XRTIMP(BOOL)                    Xrt3dSaveImageAsPng(HXRT3D, TCHAR *, BOOL);

XRTIMPORT (void) CDECL          Xrt3dSetValues(HXRT3D w, ... );
XRTIMPORT (void) CDECL          Xrt3dGetValues(HXRT3D w, ... );
XRTIMPORT (void) CDECL          Xrt3dTextSetValues(HXRT3DTEXT w, ... );
XRTIMPORT (void) CDECL          Xrt3dTextGetValues(HXRT3DTEXT w, ... );
XRTIMP(HXRT3DTEXT)              Xrt3dTextCreate(HXRT3D w);
XRTIMP(void)                    Xrt3dTextDestroy(HXRT3DTEXT w);
XRTIMP(BOOL)                    Xrt3dGetPropString(HXRT3D, int, TCHAR **);
XRTIMP(BOOL)                    Xrt3dSetPropString(HXRT3D, int, TCHAR *);
XRTIMP(BOOL)                    Xrt3dTextGetPropString(HXRT3DTEXT, int, TCHAR **);
XRTIMP(BOOL)                    Xrt3dTextSetPropString(HXRT3DTEXT, int, TCHAR *);
XRTIMP(void)                    Xrt3dFreePropString(TCHAR *);
XRTIMP(Xrt3dContourStyle *)     Xrt3dGetNthContourStyle(HXRT3D, int, int);
XRTIMP(void)                    Xrt3dSetNthContourStyle(HXRT3D, int,
                                                        Xrt3dContourStyle *, int);
XRTIMP(Xrt3dContourStyle **)    Xrt3dResetContourStyles(HXRT3D);
XRTIMP(Xrt3dContourStyle **)    Xrt3dDupContourStyles(Xrt3dContourStyle **);
XRTIMP(Xrt3dContourStyle **)    Xrt3dContourStylesFromFile(TCHAR *, TCHAR *);
XRTIMP(int)                     Xrt3dContourStylesToFile(Xrt3dContourStyle **, TCHAR *, TCHAR *);
XRTIMP(Xrt3dData *)             Xrt3dMakeGridData(int, int, double, double,
                                                  double, double, double, BOOL);
XRTIMP(Xrt3dData *)             Xrt3dMakeIrGridData(int, int, double, BOOL);
XRTIMP(Xrt3dData *)             Xrt3dMakePointData(int, int, double, BOOL);
XRTIMP(BOOL)                    Xrt3dMakePointDataSeries(Xrt3dPointSeries* series, int npoints);
XRTIMP(void)                    Xrt3dDestroy(HXRT3D);
XRTIMP(void)                    Xrt3dDestroyData(Xrt3dData *, BOOL);
XRTIMP(Xrt3dData *)             Xrt3dMakeDataFromFile(TCHAR *, TCHAR *);
XRTIMP(int)                     Xrt3dSaveDataToFile(Xrt3dData *, TCHAR *, TCHAR *);
XRTIMP(Xrt3dData *)             Xrt3dDataCopy(Xrt3dData *);
XRTIMP(Xrt3dData *)             Xrt3dDataShaded(Xrt3dData *, double, double,
                                                double, double, double);
XRTIMP(void)                    Xrt3dDataSmooth(Xrt3dData *, double);
XRTIMP(int)                     Xrt3dDataUpdateDataValue(Xrt3dData *, double,
                                                         double);
XRTIMP(Xrt3dData *)             Xrt3dDataContours(HXRT3D, int, int, int);
XRTIMP(Xrt3dData *)             Xrt3dDataWindow(Xrt3dData *, double, double,
                                                double, double, int, int,
                                                Xrt3dInterpMethod);
XRTIMP(Xrt3dDistnTable *)       Xrt3dGetDistnTable(HXRT3D);
XRTIMP(Xrt3dDistnTable *)       Xrt3dDupDistnTable(Xrt3dDistnTable *);
XRTIMP(int)                     Xrt3dDistnIndex(HXRT3D, double);
XRTIMP(Xrt3dDistnTable *)       Xrt3dDistnTableFromFile(TCHAR *, TCHAR *);
XRTIMP(int)                     Xrt3dDistnTableToFile(Xrt3dDistnTable *, TCHAR *, TCHAR *);
XRTIMP(Xrt3dRegion)             Xrt3dMap(HXRT3D, XrtPosition, XrtPosition,
                                         Xrt3dMapResult *);
XRTIMP(Xrt3dRegion)             Xrt3dPick(HXRT3D, XrtPosition, XrtPosition,
                                          Xrt3dPickResult *);
XRTIMP(TCHAR **)                 Xrt3dDupStrings(TCHAR **);
XRTIMP(void)                    Xrt3dFreeContourStyles(Xrt3dContourStyle **);
XRTIMP(void)                    Xrt3dFreeDistnTable(Xrt3dDistnTable *);
XRTIMP(void)                    Xrt3dFreeStrings(TCHAR **);
XRTIMP(void)                    Xrt3dUnmap(HXRT3D, double, double, double,
                                           Xrt3dMapResult *);
XRTIMP(void)                    Xrt3dUnpick(HXRT3D, int, int, Xrt3dPickResult *);
XRTIMP(double)                  Xrt3dComputeZValue(HXRT3D, int, int, int, int);
XRTIMP(void)                    Xrt3dComputeZValueIndirect(HXRT3D, int, int, int, int, double*);
XRTIMP(double)                  Xrt3dZInterpolate(HXRT3D, double, double);
XRTIMP(void)                    Xrt3dZInterpolateIndirect(HXRT3D, double, double, double*);
XRTIMP(TCHAR *)                  Xrt3dGetNthDataLabel(HXRT3D, Xrt3dAxis, int);
XRTIMP(TCHAR *)                  Xrt3dGetNthFooterString(HXRT3D, int);
XRTIMP(TCHAR *)                  Xrt3dGetNthHeaderString(HXRT3D, int);
XRTIMP(TCHAR *)                  Xrt3dGetNthLegendString(HXRT3D, int);
XRTIMP(Xrt3dValueLabel *)       Xrt3dGetValueLabel(HXRT3D, Xrt3dAxis, Xrt3dValueLabel *);
XRTIMP(void)                    Xrt3dSetNthDataLabel(HXRT3D, Xrt3dAxis, int, TCHAR *);
XRTIMP(void)                    Xrt3dSetNthFooterString(HXRT3D, int, TCHAR *);
XRTIMP(void)                    Xrt3dSetNthHeaderString(HXRT3D, int, TCHAR *);
XRTIMP(void)                    Xrt3dSetNthLegendString(HXRT3D, int, TCHAR *);
XRTIMP(void)                    Xrt3dSetValueLabel(HXRT3D, Xrt3dAxis, Xrt3dValueLabel *);
XRTIMP(XrtColor)                Xrt3dGetXYColor(HXRT3D, int, int);
XRTIMP(void)                    Xrt3dSetXYColor(HXRT3D, int, int, XrtColor);
XRTIMP(void)                    Xrt3dFreeValueLabels(Xrt3dValueLabel **);
XRTIMP(Xrt3dValueLabel **)      Xrt3dDupValueLabels(Xrt3dValueLabel **);
XRTIMP(void)                    Xrt3dFreeXYColors(Xrt3dXYColor **);
XRTIMP(Xrt3dXYColor **)         Xrt3dDupXYColors(Xrt3dXYColor **);
XRTIMP(Xrt3dDataStyle **)       Xrt3dDupDataStyles(Xrt3dDataStyle **dstyles);
XRTIMP(void)                    Xrt3dFreeDataStyles(Xrt3dDataStyle **dstyles);
XRTIMP(Xrt3dDataStyle *)        Xrt3dGetNthDataStyle(HXRT3D, int);
XRTIMP(int)                     Xrt3dInsertNthDataStyle(HXRT3D, int, Xrt3dDataStyle *);
XRTIMP(int)                     Xrt3dRemoveNthDataStyle(HXRT3D, int);
XRTIMP(void)                    Xrt3dSetNthDataStyle(HXRT3D, int, Xrt3dDataStyle *);
XRTIMP(Xrt3dAction)             Xrt3dGetAction(HXRT3D widget, UINT msg, UINT modifier, UINT keycode);
XRTIMP(void)                    Xrt3dSetAction(HXRT3D widget, UINT msg, UINT modifier, UINT keycode, Xrt3dAction action);
XRTIMP(Xrt3dActionItem *)       Xrt3dGetActionList(HXRT3D widget);
XRTIMP(void)                    Xrt3dCallAction(HXRT3D widget, Xrt3dAction action, XrtPosition x, XrtPosition y);
XRTIMP(void)                    Xrt3dRemoveAllActions(HXRT3D widget);
XRTIMP(void)                    Xrt3dResetAllActions(HXRT3D widget);

XRTIMP(void)                    Xrt3dClearVersionInfo(Xrt3dVersionInfo *vinfo);
XRTIMP(HINSTANCE)               Xrt3dGetLocalizedResourceHandle(HXRT3D hXrt3d);
XRTIMP(BOOL)                    Xrt3dGetModuleVersionInfo(HINSTANCE hInst, Xrt3dVersionInfo *vinfo);
XRTIMP(TCHAR*)                   Xrt3dLoadResourceString(HXRT3D hXrt3d, long strID, TCHAR * strbuf, int strbufLen);
XRTIMP(void)                    Xrt3dReleaseLocalizedResourceHandle(HXRT3D hXrt3d);

XRTIMP(DWORD)                Xrt3dSaveImageAsJpegBytes(HXRT3D, LPVOID, DWORD, int, BOOL, BOOL, BOOL);
XRTIMP(DWORD)                Xrt3dSaveImageAsPngBytes(HXRT3D, LPVOID, DWORD, BOOL);
XRTIMP(DWORD)                Xrt3dSaveImageAsDibBytes(HXRT3D, LPVOID, DWORD);
#if defined(__cplusplus)
}
#endif

#endif /* _XRT3D_WIN_H */
