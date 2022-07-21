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

#ifndef _XRTNT_H
#define _XRTNT_H

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
#include "olch2dcm.h"

#if defined(__cplusplus)
extern "C" {
#endif

XRTIMP(void)                XrtAboutDialogBox(HWND);
XRTIMP(void)                XrtClearVersionInfo(XrtVersionInfo *vinfo);
XRTIMP(HXRT2D)              XrtCreate(void);
XRTIMP(HXRT2D)              XrtCreateWindow(LPCTSTR, int, int, int, int, HWND, HANDLE);
XRTIMP(BOOL)                XrtAttachWindow(HXRT2D, HWND);
XRTIMP(void)                XrtComputePalette(HXRT2D);
XRTIMP(HWND)                XrtDetachWindow(HXRT2D);
XRTIMP(HWND)                XrtGetWindow(HXRT2D);
XRTIMP(HXRT2D)              XrtGetHandle(HWND);
XRTIMP(int)                 XrtArrCheckAxisBounds(HXRT2D, int, int);
XRTIMP(int)                 XrtArrDataAppendPts(XrtDataHandle, double, double   *);
XRTIMP(int)                 XrtArrDataFastUpdate(HXRT2D, int, int);
XRTIMP(int)                 XrtArrDataRemovePts(XrtDataHandle, int);
XRTIMP(int)                 XrtArrDataShiftPts(XrtDataHandle, int, int, int);
XRTIMP(void)                XrtCallAction(HXRT2D widget, XrtAction action, XrtPosition x, XrtPosition y);
XRTIMP(void)                XrtClearAlternateDataStyles(HXRT2D hXrt2D, int dataset);
XRTIMP(int)                 XrtCustomFormatValidate(TCHAR* format);
XRTIMP(XrtDataHandle)       XrtDataConcat(XrtDataHandle, XrtDataHandle);
XRTIMP(XrtDataHandle)       XrtDataCopy(XrtDataHandle);
XRTIMP(XrtDataHandle)       XrtDataConvertToHandle(XrtData * data);
XRTIMP(XrtDataHandle)       XrtDataCreate(XrtDataType, int, int);
XRTIMP(XrtDataHandle)       XrtDataCreateFromFile(TCHAR *, TCHAR *);
XRTIMP(void)                XrtDataDestroy(XrtDataHandle);
XRTIMP(XrtDataHandle)       XrtDataExtractSet(XrtDataHandle, int);
XRTIMP(XrtDisplay)          XrtDataGetDisplay(XrtDataHandle hData, int set);
XRTIMP(int)                 XrtDataGetFirstPoint(XrtDataHandle hData, int set);
XRTIMP(int)                 XrtDataGetFirstSet(XrtDataHandle hData);
XRTIMP(double)              XrtDataGetHole(XrtDataHandle hData);
XRTIMP(void)                XrtDataGetHoleIndirect(XrtDataHandle hData, double* value);
XRTIMP(int)                 XrtDataGetLastPoint(XrtDataHandle hData, int set);
XRTIMP(int)                 XrtDataGetLastSet(XrtDataHandle hData);
XRTIMP(int)                 XrtDataGetNPoints(XrtDataHandle hData, int set);
XRTIMP(int)                 XrtDataGetNSets(XrtDataHandle hData);
XRTIMP(XrtDataType)         XrtDataGetType(XrtDataHandle hData);
XRTIMP(double   *)          XrtDataGetXData(XrtDataHandle hData, int set);
XRTIMP(double)              XrtDataGetXElement(XrtDataHandle hData, int set, int point);
XRTIMP(void)                XrtDataGetXElementIndirect(XrtDataHandle hData, int set, int point, double* value);
XRTIMP(double   *)          XrtDataGetYData(XrtDataHandle hData, int set);
XRTIMP(double)              XrtDataGetYElement(XrtDataHandle hData, int set, int point);
XRTIMP(void)                XrtDataGetYElementIndirect(XrtDataHandle hData, int set, int point, double* value);
XRTIMP(double)              XrtDataGetYMean(XrtDataHandle hData, int set);
XRTIMP(double)              XrtDataGetYMedian(XrtDataHandle hData, int set);
XRTIMP(double)              XrtDataGetYStdDev(XrtDataHandle hData, int set);
XRTIMP(double)              XrtDataGetYAveDev(XrtDataHandle hData, int set);
XRTIMP(long)                XrtDataGetYDataCount(XrtDataHandle hData, int set);
XRTIMP(double)              XrtDataGetYDataMax(XrtDataHandle hData, int set);
XRTIMP(double)              XrtDataGetYDataMin(XrtDataHandle hData, int set);
XRTIMP(BOOL)                XrtDataGetSeriesLeastSquaresPoly(XrtDataHandle hData, int set, int order, double ** coefficents);
XRTIMP(double)              XrtDataPolynomialEvaluate(double xvalue, int order, double * coefficents);
XRTIMP(BOOL)                XrtDataGetLeastSquaresPoly(XrtDataHandle hData, const double * xv, const double *yv, int npoints, int order, double ** coefficients);
XRTIMP(int)                 XrtDataRelease(XrtDataHandle hData);
XRTIMP(int)                 XrtDataRemoveSet(XrtDataHandle, int);
XRTIMP(int)                 XrtDataSaveToFile(XrtDataHandle, TCHAR *, TCHAR *);
XRTIMP(int)                 XrtDataSetDisplay(XrtDataHandle hData, int set, XrtDisplay display);
XRTIMP(int)                 XrtDataSetFirstPoint(XrtDataHandle hData, int set, int point);
XRTIMP(int)                 XrtDataSetFirstSet(XrtDataHandle hData, int set);
XRTIMP(int)                 XrtDataSetHole(XrtDataHandle hData, double hole);
XRTIMP(int)                 XrtDataSetLastPoint(XrtDataHandle hData, int set, int point);
XRTIMP(int)                 XrtDataSetLastSet(XrtDataHandle hData, int set);
XRTIMP(int)                 XrtDataSetNPoints(XrtDataHandle hData, int set, int npoints);
XRTIMP(int)                 XrtDataSetNSets(XrtDataHandle hData, int nsets);
XRTIMP(int)                 XrtDataSetType(XrtDataHandle hData, XrtDataType type);
XRTIMP(int)                 XrtDataSetXData(XrtDataHandle hData, int set, double   *xvalues, int n, int index);
XRTIMP(int)                 XrtDataSetXElement(XrtDataHandle hData, int set, int point, double x);
XRTIMP(int)                 XrtDataSetYData(XrtDataHandle hData, int set, double   *yvalues, int n, int index);
XRTIMP(int)                 XrtDataSetYElement(XrtDataHandle hData, int set, int point, double y);
XRTIMP(int)                 XrtDataSort(XrtDataHandle);
XRTIMP(int)                 XrtDataUpdateDataValue(XrtDataHandle hData, double   oldValue, double   newValue);
XRTIMP(XrtDataHandle)       XrtDataTranspose(XrtDataHandle);
XRTIMP(BOOL)                XrtDeletePointLabel(HXRT2D, int);
XRTIMP(BOOL)                XrtDeletePointLabel2(HXRT2D, int);
XRTIMP(BOOL)                XrtDeleteSetLabel(HXRT2D, int);
XRTIMP(BOOL)                XrtDeleteSetLabel2(HXRT2D, int);
XRTIMP(void)                XrtDestroy(HXRT2D);
XRTIMP(void)                XrtDestroyData(XrtData *, int);
XRTIMP(BOOL)                XrtDrawToClipboard(HXRT2D, XrtDrawFormat);
XRTIMP(BOOL)                XrtDrawToDC(HXRT2D, HDC, XrtDrawFormat, XrtDrawScale, int, int, int, int);
XRTIMP(BOOL)                XrtDrawToFile(HXRT2D, TCHAR *, XrtDrawFormat);
XRTIMP(XrtDataStyle **)     XrtDupDataStyles(XrtDataStyle **);
XRTIMP(TCHAR **)             XrtDupStrings(TCHAR **);
XRTIMP(XrtValueLabel **)    XrtDupValueLabels(XrtValueLabel **);
XRTIMP(void)                XrtForceRedraw(HXRT2D hXrt2D);
XRTIMP(void)                XrtFreeDataStyles(XrtDataStyle **);
XRTIMP(void)                XrtFreePropString(TCHAR *);
XRTIMP(void)                XrtFreeStrings(TCHAR **);
XRTIMP(void)                XrtFreeTextHandles(XrtTextHandle *);
XRTIMP(void)                XrtFreeValueLabels(XrtValueLabel **);
XRTIMP(int)                 XrtGenCheckAxisBounds(HXRT2D, int, int, int);
XRTIMP(int)                 XrtGenDataAppendPt(XrtDataHandle, int, double, double);
XRTIMP(int)                 XrtGenDataFastUpdate(HXRT2D, int, int, int);
XRTIMP(int)                 XrtGenDataRemovePt(XrtDataHandle, int, int);
XRTIMP(int)                 XrtGenDataShiftPts(XrtDataHandle, int, int, int, int);
XRTIMP(XrtAction)           XrtGetAction(HXRT2D widget, UINT msg, UINT modifier, UINT keycode);
XRTIMP(XrtActionItem *)     XrtGetActionList(HXRT2D widget);
XRTIMP(XrtBoolean)          XrtAddAlarmZone(HXRT2D widget, const TCHAR *name);
XRTIMP(XrtBoolean)          XrtAddAlarmZoneBefore(HXRT2D widget, const char *name, const char *before);
XRTIMP(XrtBoolean)          XrtAlarmZoneValid(HXRT2D widget, XrtAlarmZone *zone);
XRTIMP(XrtAlarmZone *)      XrtGetAlarmZone(HXRT2D widget, const TCHAR *name);
XRTIMP(XrtAlarmZone *)      XrtGetAlarmZoneList(HXRT2D widget);
XRTIMP(XrtDataStyle *)      XrtGetAlternateDataStyle(HXRT2D hXrt2D, int dataset, int set, int point);
XRTIMP(XrtAlternateDataStyle **) XrtGetAlternateDataStyleList(HXRT2D hXrt2D, int dataset);
XRTIMP(HINSTANCE)           XrtGetLocalizedResourceHandle(HXRT2D hXrt2d);
XRTIMP(BOOL)                XrtGetModuleVersionInfo(HINSTANCE hInst, XrtVersionInfo *vinfo);
XRTIMP(XrtDataStyle *)      XrtGetNthDataStyle(HXRT2D, int);
XRTIMP(XrtDataStyle *)      XrtGetNthDataStyle2(HXRT2D, int);
XRTIMP(TCHAR *)              XrtGetNthFooterString(HXRT2D, int);
XRTIMP(TCHAR *)              XrtGetNthHeaderString(HXRT2D, int);
XRTIMP(TCHAR *)              XrtGetNthPointLabel(HXRT2D, int);
XRTIMP(TCHAR *)              XrtGetNthPointLabel2(HXRT2D, int);
XRTIMP(TCHAR *)              XrtGetNthSetLabel(HXRT2D, int);
XRTIMP(TCHAR *)              XrtGetNthSetLabel2(HXRT2D, int);
XRTIMP(int *)               XrtGetOtherSliceSets(HXRT2D, int, int);
XRTIMP(HPALETTE)            XrtGetPalette(HXRT2D);
XRTIMP(int)                 XrtGetPlotAreaLeftExtent(HXRT2D);
XRTIMP(int)                 XrtGetPlotAreaRightExtent(HXRT2D);
XRTIMP(BOOL)                XrtGetPropString(HXRT2D, XrtProperty, TCHAR **);
XRTIMP(int)                 XrtGetTextHandles(HXRT2D, XrtTextHandle **);
XRTIMP(XrtValueLabel *)     XrtGetValueLabel(HXRT2D, XrtAxis, XrtValueLabel *);
XRTIMPORT(void) CDECL       XrtGetValues(HXRT2D w, ... );
XRTIMP(int)                 XrtInsertNthDataStyle(HXRT2D, int, XrtDataStyle *, int);
XRTIMP(BOOL)                XrtInsertPointLabel(HXRT2D, int, TCHAR *);
XRTIMP(BOOL)                XrtInsertPointLabel2(HXRT2D, int, TCHAR *);
XRTIMP(BOOL)                XrtInsertSetLabel(HXRT2D, int, TCHAR *);
XRTIMP(BOOL)                XrtInsertSetLabel2(HXRT2D, int, TCHAR *);
XRTIMP(TCHAR*)               XrtLoadResourceString(HXRT2D hXrt2d, long strID, TCHAR * strbuf, int strbufLen);
XRTIMP(XrtData *)           XrtMakeData(XrtDataType, int, int, int);
XRTIMP(XrtData *)           XrtMakeDataFromFile(TCHAR *, TCHAR *);
XRTIMP(long)                XrtMakeTime(int, int, int, int, int, int);
XRTIMP(XrtRegion)           XrtMap(HXRT2D, int, int, int, XrtMapResult *);
XRTIMP(XrtRegion)           XrtPick(HXRT2D, int, int, int, XrtPickResult *, XrtFocus);
XRTIMP(XrtPointStyleHandle) XrtPointStyleCreate(HXRT2D);
XRTIMP(void)                XrtPointStyleDestroy(XrtPointStyleHandle);
XRTIMPORT(void) CDECL       XrtPointStyleGetValues(XrtPointStyleHandle, ... );
XRTIMP(XrtPointStyleHandle) XrtPointStyleFindExact(HXRT2D, int, int, int);
XRTIMPORT(void) CDECL       XrtPointStyleSetValues(XrtPointStyleHandle, ... );
XRTIMP(BOOL)                XrtPrint(HXRT2D, XrtDrawFormat, XrtDrawScale, int, int, int, int);
XRTIMP(void)                XrtReinitialize(HXRT2D widget);
XRTIMP(void)                XrtRemoveAllActions(HXRT2D widget);
XRTIMP(int)                 XrtRemoveNthDataStyle(HXRT2D, int, int);
XRTIMP(void)                XrtReleaseLocalizedResourceHandle(HXRT2D hXrt2d);
XRTIMP(void)                XrtRemoveAlarmZone(HXRT2D widget, const TCHAR *name);
XRTIMP(void)                XrtRemoveAllActions(HXRT2D widget);
XRTIMP(void)                XrtRemoveAllAlarmZones(HXRT2D widget);
XRTIMP(int)                 XrtRemoveNthDataStyle(HXRT2D, int, int);
XRTIMP(XrtBoolean)          XrtRenameAlarmZone(HXRT2D, XrtAlarmZone*, const TCHAR*);
XRTIMP(HANDLE)              XrtRenderClipboardFormat(HXRT2D widget, int cf);
XRTIMP(void)                XrtResetAllActions(HXRT2D widget);
XRTIMP(int)                 XrtSaveDataToFile(XrtData *, TCHAR *, TCHAR *);
XRTIMP(BOOL)                XrtSaveImageAsJpeg(HXRT2D, TCHAR *, int, BOOL, BOOL, BOOL);
XRTIMP(BOOL)                XrtSaveImageAsPng(HXRT2D, TCHAR *, BOOL);
XRTIMP(void)                XrtSetAction(HXRT2D widget, UINT msg, UINT modifier, UINT keycode, XrtAction action);
XRTIMP(void)                XrtSetAlarmZone(HXRT2D widget, const TCHAR *name, BOOL is_shown,
                                            XrtFloat LowerY, XrtFloat UpperY, COLORREF line_color,
                                            COLORREF fill_color, XrtFillPattern fill_pattern);
XRTIMP(void)                XrtSetWorkColor(HXRT2D, XrtColor);
XRTIMP(void)                XrtSetAlternateDataStyle(HXRT2D hXrt2D, int dataset, int set, int point, XrtDataStyle *ds);
XRTIMP(void)                XrtSetNthDataStyle(HXRT2D, int, XrtDataStyle *);
XRTIMP(void)                XrtSetNthDataStyle2(HXRT2D, int, XrtDataStyle *);
XRTIMP(void)                XrtSetNthFooterString(HXRT2D, int, TCHAR *);
XRTIMP(void)                XrtSetNthHeaderString(HXRT2D, int, TCHAR *);
XRTIMP(void)                XrtSetNthPointLabel(HXRT2D, int, TCHAR *);
XRTIMP(void)                XrtSetNthPointLabel2(HXRT2D, int, TCHAR *);
XRTIMP(void)                XrtSetNthSetLabel(HXRT2D, int, TCHAR *);
XRTIMP(void)                XrtSetNthSetLabel2(HXRT2D, int, TCHAR *);
XRTIMP(BOOL)                XrtSetPropString(HXRT2D, XrtProperty, TCHAR *);
XRTIMP(void)                XrtSetValueLabel(HXRT2D, XrtAxis, XrtValueLabel *);
XRTIMPORT(void) CDECL       XrtSetValues(HXRT2D w, ... );
XRTIMP(XrtTextHandle)       XrtTextAreaCreate(HXRT2D);
XRTIMP(void)                XrtTextAreaDestroy(XrtTextHandle);
XRTIMP(BOOL)                XrtTextAreaGetPropString(XrtTextHandle, XrtProperty, TCHAR **);
XRTIMPORT(void) CDECL       XrtTextAreaGetValues(XrtTextHandle, ... );
XRTIMP(BOOL)                XrtTextAreaSetPropString(XrtTextHandle, XrtProperty, TCHAR *);
XRTIMPORT(void) CDECL       XrtTextAreaSetValues(XrtTextHandle, ... );
XRTIMP(void)                XrtTextAttach(HXRT2D, XrtTextHandle);
XRTIMP(XrtTextHandle)       XrtTextCreate(HXRT2D, XrtTextDesc *);
XRTIMP(void)                XrtTextDestroy(HXRT2D, XrtTextHandle);
XRTIMP(void)                XrtTextDetach(HXRT2D, XrtTextHandle);
XRTIMP(int)                 XrtTextDetail(HXRT2D, XrtTextHandle, XrtTextDesc *);
XRTIMP(void)                XrtTextUpdate(HXRT2D, XrtTextHandle, XrtTextDesc *);
XRTIMP(void)                XrtUnmap(HXRT2D, int, double, double, XrtMapResult *);
XRTIMP(void)                XrtUnpick(HXRT2D, int, int, int, XrtPickResult *);
XRTIMP(time_t)              XrtValueToTime(HXRT2D w, double   value);
XRTIMP(double  )            XrtTimeToValue(HXRT2D w, time_t tvalue);

XRTEXP(BOOL)				XrtLoadPropertiesFromMemory(HXRT2D, HINSTANCE, LPCTSTR, BOOL);
XRTEXP(BOOL)				XrtLoadPropertiesFromFile(HXRT2D, TCHAR *, BOOL);
XRTEXP(BOOL)				XrtSavePropertiesToFile(HXRT2D, TCHAR *);

XRTIMP(DWORD)                XrtSaveImageAsJpegBytes(HXRT2D, LPVOID, DWORD, int, BOOL, BOOL, BOOL);
XRTIMP(DWORD)                XrtSaveImageAsPngBytes(HXRT2D, LPVOID, DWORD, BOOL);
XRTIMP(DWORD)                XrtSaveImageAsDibBytes(HXRT2D, LPVOID, DWORD);

XRTIMP(BOOL)                XrtCheckBaseTime(time_t tvalue);
XRTIMP(BOOL)                XrtTimeToStandardFormat(time_t, struct tm * tmDest);
XRTEXP(double)			    XrtMakeTimeOle(int yr,int mon,int day,int hr,int min,int sec);
XRTIMP(double)              XrtValueToTimeOle(HXRT2D w, double value);
XRTIMP(double)              XrtTimeToValueOle(HXRT2D w, double tvalue);
XRTIMP(BOOL)                XrtCheckBaseTimeOle(DATE dtBase);
XRTIMP(BOOL)                XrtTimeToStandardFormatOle(DATE dt, struct tm * tmDest);

#if defined(__cplusplus)
}
#endif

#endif
