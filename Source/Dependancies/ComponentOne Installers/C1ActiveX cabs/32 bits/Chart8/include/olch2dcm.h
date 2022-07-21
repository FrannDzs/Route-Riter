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

/** \file xrtgfcom.h
 *
 *  This include file is to contain all the definitions which are
 *  necessary for both the API level and the XRT implementation level;
 *  include files at both of these levels include this file.
 *
 */

#if !defined(LONG_PTR) && defined(_MSC_VER) && _MSC_VER < 1400
#define LONG_PTR long
#endif

#ifndef xrt_com_DEFINED
#define xrt_com_DEFINED

#ifndef RC_INVOKED
#include <math.h>
#endif
#include <time.h>

#ifdef __STDC__
#include <float.h>
#define XRT_HUGE_VAL                                  DBL_MAX

#else /* !__STDC__ */

#include <float.h>
/* DO NOT USE DBL_MAX!  VB can't handle it - this is big enough */
#define XRT_HUGE_VAL                                  1.0E308


#endif /* __STDC__ */

# if defined(_WIN32)
#  include "pshpack8.h"
# else
#  if !defined(RC_INVOKED)
#   if defined(_MSC_VER)
#     pragma pack(2)
#   else
#    include "pshpack2.h"
#   endif
#  endif
# endif


#ifndef __wtypes_h__
#ifndef DATE
typedef double DATE;		// defined in wtypes.h
#endif
#endif

/*
 *  Version numbers and strings, in case developers need
 *  to have any version dependent code.  This scheme allows
 *  for 99 revisions, 9 point updates.  Also provide a
 *  string, for printing purposes.
*/
#define XRT_VERSION      4
#define XRT_REVISION     0
#define XRT_POINT        0
#define XRT_RELEASE      \
                    (XRT_VERSION * 1000 + XRT_REVISION*10 + XRT_POINT)
#define XRT_RELEASE_STRING   _T("4.0.0")

#define XRT_DEFAULT_FLOAT               XRT_HUGE_VAL


#define XRT_MAX_PRECISION               14

#define arr_nsets                       a.nsets
#define arr_npoints                     a.npoints
#define arr_data                        a.data
#define arr_hole                        a.hole
#define arr_xdata                       arr_data.xp
#define arr_ydata                       arr_data.yp
#define arr_xel(j)                      arr_xdata[j]
#define arr_yel(i, j)                   arr_ydata[i][j]

#define gen_nsets                       g.nsets
#define gen_data                        g.data
#define gen_hole                        g.hole
#define gen_npoints(i)                  gen_data[i].npoints
#define gen_xdata(i)                    gen_data[i].xp
#define gen_ydata(i)                    gen_data[i].yp
#define gen_xel(i, j)                   gen_xdata(i)[j]
#define gen_yel(i, j)                   gen_ydata(i)[j]

#define XRT_DEFAULT_POINT_SIZE          7

#define XRT_NORTHSOUTH_MASK             0xf0
#define XRT_EASTWEST_MASK               0x0f

#define XRT_DEFAULT_BORDER_WIDTH        2
#define XRT_MAX_BORDER_WIDTH            20

#define XRT_MAX_BAR_CLUSTER_OVERLAP     200
#define XRT_DEFAULT_BAR_CLUSTER_OVERLAP 100
#define XRT_MAX_BAR_CLUSTER_WIDTH       100
#define XRT_DEFAULT_BAR_CLUSTER_WIDTH   50

#define XRT_MAX_DEPTH                   500
#define XRT_MAX_ROTATION                45
#define XRT_MAX_INCLINATION             45
#define XRT_DEFAULT_START_ANGLE			135

#define XRT_MIN_MARGIN                  -30000
#define XRT_MAX_MARGIN                  30000

#define XRT_XFOCUS                      0x1
#define XRT_YFOCUS                      0x2

#define XRT_DATASET1                    0x1
#define XRT_DATASET2                    0x2
#define XRT_ALL_DATA                    (XRT_DATASET1 | XRT_DATASET2)

#define XRT_OTHER_SLICE                 -10
#define XRT_DEFAULT_OTHER_LABEL         _T("Other")

#define XRT_DEFAULT_BUBBLE_MIN			5
#define XRT_DEFAULT_BUBBLE_MAX			20

#define XRT_LTX                         0x01
#define XRT_GTX                         0x02
#define XRT_LTY                         0x04
#define XRT_GTY                         0x08

#define XRT_SET_NULL                      -1
#define XRT_SET_ALL                       -2
#define XRT_SET_OTHER_SLICE  XRT_OTHER_SLICE   /* -10 */

#define XRT_POINT_NULL                    -1
#define XRT_POINT_ALL                     -2
#define XRT_POINT_LEGEND                  -3

/*
 *  Aliases for pre-2.3 names; these might disappear
 *  in a future release, so don't depend on them.
*/

#define     XRT_PLOT                XRT_TYPE_PLOT
#define     XRT_BAR                 XRT_TYPE_BAR
#define     XRT_PIE                 XRT_TYPE_PIE
#define     XRT_STACKING_BAR        XRT_TYPE_STACKING_BAR
#define     XRT_AREA                XRT_TYPE_AREA

#define     XRT_ARRAY               XRT_DATA_ARRAY
#define     XRT_GENERAL             XRT_DATA_GENERAL

#define     XRT_LEFT                XRT_ADJUST_LEFT
#define     XRT_RIGHT               XRT_ADJUST_RIGHT
#define     XRT_CENTER              XRT_ADJUST_CENTER

#define     XRT_ASCENDING           XRT_PIEORDER_ASCENDING
#define     XRT_DESCENDING          XRT_PIEORDER_DESCENDING
#define     XRT_DATA_ORDER          XRT_PIEORDER_DATA_ORDER

#define     XRT_NORTH               XRT_ANCHOR_NORTH
#define     XRT_SOUTH               XRT_ANCHOR_SOUTH
#define     XRT_EAST                XRT_ANCHOR_EAST
#define     XRT_WEST                XRT_ANCHOR_WEST
#define     XRT_NORTHEAST           XRT_ANCHOR_NORTHEAST
#define     XRT_NORTHWEST           XRT_ANCHOR_NORTHWEST
#define     XRT_SOUTHEAST           XRT_ANCHOR_SOUTHEAST
#define     XRT_SOUTHWEST           XRT_ANCHOR_SOUTHWEST
#define     XRT_HOME                XRT_ANCHOR_HOME
#define     XRT_BEST                XRT_ANCHOR_BEST

#define     XRT_VERTICAL            XRT_ALIGN_VERTICAL
#define     XRT_HORIZONTAL          XRT_ALIGN_HORIZONTAL

#define     XRT_IN_GRAPH            XRT_RGN_IN_GRAPH
#define     XRT_IN_LEGEND           XRT_RGN_IN_LEGEND
#define     XRT_IN_HEADER           XRT_RGN_IN_HEADER
#define     XRT_IN_FOOTER           XRT_RGN_IN_FOOTER
#define     XRT_NOWHERE             XRT_RGN_NOWHERE

#define     XRT_USE_XVALUES         XRT_XMETHOD_XVALUES
#define     XRT_USE_POINT_LABELS    XRT_XMETHOD_POINT_LABELS
#define     XRT_USE_XLABELS         XRT_XMETHOD_XLABELS

/* Change statuses, for handling repaints */

#define XRT_GRAPH_CHANGED       0x01
#define XRT_HEADER_CHANGED      0x02
#define XRT_FOOTER_CHANGED      0x04
#define XRT_LEGEND_CHANGED      0x08
#define XRT_MARKER_CHANGED      0x10
#define XRT_DATA_CHANGED        (XRT_GRAPH_CHANGED | XRT_LEGEND_CHANGED)
#define XRT_ALL_CHANGED         0xFF        /* all of the above        */

# ifdef XRT_STATICLIB
#  define XRTEXPORT(type)   extern type
#  define XRTIMPORT(type)   extern type
#  define XRTEXP(type)      XRTEXPORT(type)
#  define XRTIMP(type)      XRTIMPORT(type)
# else
#  define XRTEXPORT(type)  extern __declspec( dllexport ) type
#  define XRTIMPORT(type)  extern __declspec( dllimport ) type
#  define XRTEXP(type)     XRTEXPORT(type)
#  define XRTIMP(type)     XRTIMPORT(type)
# endif

#ifndef XRT_TYPES_DEFINED
/* Portability section - use these typedefs on Windows or Motif */


typedef COLORREF            XrtColor;      /* colors specified as COLORREF */
typedef BOOL                XrtBoolean;
typedef int                 XrtDimension;  /* 16-bit GDI in Win16, 32-bit GDI in Win32 */
typedef int                 XrtPosition;   /* 16-bit GDI in Win16, 32-bit GDI in Win32 */
typedef HFONT               XrtFont;
typedef double              XrtFloat;      /* double for all but Motif */

#define XRT_TYPES_DEFINED
#endif /* XRT_TYPES_DEFINED */


#ifndef XRT2D_TYPES_DEFINED

typedef TCHAR *              XrtImage;

#define XRT2D_TYPES_DEFINED
#endif /* XRT2D_TYPES_DEFINED */


/*--------------------------------------------------
 * Windows specific definitions
 *-------------------------------------------------*/

#ifndef MK_ALT
#define MK_ALT 0x20     /* Win32 defines this, Win16 doesn't */
#endif

 #ifdef _UNICODE
  #ifdef XRT_STATICLIB
   #define XRT2DWNDCLASS "C1Chart2DSU8"		/*NO UNICODE CONVERSION HERE*/
  #else
   #define XRT2DWNDCLASS "C1Chart2DU8"		/*NO UNICODE CONVERSION HERE*/
  #endif
 #else
  #ifdef XRT_STATICLIB
   #define XRT2DWNDCLASS "C1Chart2DS8"		/*NO UNICODE CONVERSION HERE*/
  #else
   #define XRT2DWNDCLASS "C1Chart2D8"		/*NO UNICODE CONVERSION HERE*/
  #endif
 #endif
 #define XRT2D _T(XRT2DWNDCLASS)

typedef struct {
    XrtPosition     x;
    XrtPosition     y;
    XrtDimension    width;
    XrtDimension    height;
} XrtRectangle;

#define XRT_DEFAULT_COLOR   0xffffffff

#if defined( STRICT )
typedef struct HXRT2D__ { int unused; } FAR* HXRT2D;
typedef struct XrtTextHandle__ { int unused; } FAR* XrtTextHandle;
typedef struct XrtDataHandle__ { int unused; } FAR* XrtDataHandle;
typedef struct XrtPointStyleHandle__ { int unused; } FAR* XrtPointStyleHandle;
#else
typedef void FAR* HXRT2D;
typedef void FAR* XrtTextHandle;
typedef void FAR* XrtDataHandle;
typedef void FAR* XrtPointStyleHandle;
#endif



/*
 *  All XRT_* properties are defined as a continuous sequence of
 *  numbers within each representation type
 */

#if ((defined(sun) && !defined(SOLARIS)) || defined(_DGUX_SOURCE)) && !defined(__STDC__) && !defined(__cplusplus)
#define XRT_PROPERTY(type, index)       ((unsigned int) (((XRTT_/**/type) << 8) | (index)))
#else
#define XRT_PROPERTY(type, index)       ((unsigned int) (((XRTT_##type) << 8) | (index)))
#endif
#define XRT_PROPERTY_TYPE(property)     ((((unsigned int) (property)) >> 8) & 0xFF)
#define XRT_PROPERTY_INDEX(property)    (((unsigned int) (property)) & 0xFF)

/** \internal All types over ENUM_BASE are enums */
#define XRTT_ENUM_BASE      64      /* keep value positive : 128
                                       introduces sign extension problems */
#define XRTT_IS_ENUM(type)  ((type) >= XRTT_ENUM_BASE)

/** \internal
 * XRT_* properties. All XRT_* properties are defined as a continuous sequence
 * of numbers within each representation type
 */
typedef enum {
    XRTT_NONE = 0,          /**< unused */
    XRTT_BOOLVAL,
    XRTT_COLOR,
    XRTT_DATA,
    XRTT_DATA_STYLE,
    XRTT_DATA_STYLES,
    XRTT_DIMENSION,
    XRTT_FLOAT,
    XRTT_FONT,
    XRTT_INT,
    XRTT_LONG,
    XRTT_POSITION,
    XRTT_STRING,
    XRTT_STRINGS,
    XRTT_USE_DEFAULT,
    XRTT_VALUE_LABELS,
    XRTT_IMAGE,

    XRTT_ADJUST = XRTT_ENUM_BASE,
    XRTT_ALIGN,
    XRTT_ANCHOR,
    XRTT_ANNO_METHOD,
    XRTT_ANNO_PLACEMENT,
    XRTT_BORDER,
    XRTT_DATASET,
    XRTT_DATA_TYPE,
    XRTT_FILL_PATTERN,
    XRTT_LINE_PATTERN,
    XRTT_NUM_METHOD,
    XRTT_ORIGIN_PLACEMENT,
    XRTT_PIE_ORDER,
    XRTT_PIE_THRESHOLD_METHOD,
    XRTT_POINT,
    XRTT_ROTATE,
    XRTT_TIME_UNIT,
    XRTT_TYPE,
    XRTT_MARKER_METHOD,
    XRTT_TEXT_ATTACH,
	XRTT_SHADING,
    XRTT_ANGLE_UNIT,
    XRTT_IMAGE_LAYOUT,
    XRTT_BUBBLE_METHOD,
    XRTT_DISPLAY
} XrtPropertyType;

#define XRTT_LAST_TYPE       XRTT_DISPLAY       /**< \internal last type in list */

typedef enum {
    /* XrtBoolean */
    XRT_DOUBLE_BUFFER                  = XRT_PROPERTY(BOOLVAL, 0),
    XRT_LEGEND_SHOW,
    XRT_XAXIS_SHOW,
    XRT_YAXIS_SHOW,
    XRT_Y2AXIS_SHOW,
    XRT_INVERT_ORIENTATION,
    XRT_TRANSPOSE_DATA,             /**< old -- this resource ignored */
    XRT_XMARKER_SHOW,
    XRT_YMARKER_SHOW,
    XRT_AXIS_BOUNDING_BOX,
    XRT_XAXIS_LOGARITHMIC,
    XRT_YAXIS_LOGARITHMIC,
    XRT_Y2AXIS_LOGARITHMIC,
    XRT_XAXIS_REVERSED,
    XRT_YAXIS_REVERSED,
    XRT_Y2AXIS_REVERSED,
    XRT_DEBUG,
    XRT_REPAINT,
	XRT_CANDLE_COMPLEX,
	XRT_HILO_CLOSE_SHOW,
	XRT_HILO_OPEN_SHOW,
	XRT_HILO_OPEN_CLOSE_FULL_WIDTH,
    XRT_POLAR_HALF_RANGE,
    XRT_YAXIS_100_PERCENT,
    XRT_Y2AXIS_100_PERCENT,
    XRT_GRAPH_SHOW_OUTLINES,
    XRT_TEXT_IS_CONNECTED,
    XRT_TEXT_IS_SHOWING,
    XRT_IS_STACKED,
    XRT_IS_STACKED2,

    XRT_TEXT_IMAGE_MINIMUM_SIZE,
    XRT_HEADER_IMAGE_MINIMUM_SIZE,
    XRT_FOOTER_IMAGE_MINIMUM_SIZE,

    XRT_EXTRA_DEFAULT_DATA_STYLES,
    XRT_POLAR_AXIS_ALLOW_NEGATIVES,
    XRT_LEGEND_REVERSED,
    XRT_LEGEND_ISSHOWING3D,
    XRT_PIE_MERGE_MISSING_SLICES,

    XRT_IMAGE_TRANSPARENT,
    XRT_GRAPH_IMAGE_TRANSPARENT,
    XRT_HEADER_IMAGE_TRANSPARENT,
    XRT_FOOTER_IMAGE_TRANSPARENT,
    XRT_LEGEND_IMAGE_TRANSPARENT,
    XRT_DATA_AREA_IMAGE_TRANSPARENT,
    XRT_TEXT_IMAGE_TRANSPARENT,

	XRT_CANDLE_FILLFALLING,

    /* XrtPosition */
    XRT_GRAPH_X                        = XRT_PROPERTY(POSITION, 0),
    XRT_GRAPH_Y,
    XRT_HEADER_X,
    XRT_HEADER_Y,
    XRT_FOOTER_X,
    XRT_FOOTER_Y,
    XRT_LEGEND_X,
    XRT_LEGEND_Y,
    XRT_GRAPH_MARGIN_BOTTOM,
    XRT_GRAPH_MARGIN_LEFT,
    XRT_GRAPH_MARGIN_RIGHT,
    XRT_GRAPH_MARGIN_TOP,
    XRT_TEXT_X,
    XRT_TEXT_Y,
    XRT_TEXT_ATTACH_PIXEL_X,
    XRT_TEXT_ATTACH_PIXEL_Y,

    /* XrtDimension */
    XRT_GRAPH_WIDTH                    = XRT_PROPERTY(DIMENSION, 0),
    XRT_GRAPH_HEIGHT,
    XRT_HEADER_WIDTH,
    XRT_HEADER_HEIGHT,
    XRT_FOOTER_WIDTH,
    XRT_FOOTER_HEIGHT,
    XRT_LEGEND_WIDTH,
    XRT_LEGEND_HEIGHT,
    XRT_BORDER_WIDTH,
    XRT_HEADER_BORDER_WIDTH,
    XRT_FOOTER_BORDER_WIDTH,
    XRT_GRAPH_BORDER_WIDTH,
    XRT_LEGEND_BORDER_WIDTH,
    XRT_WIDTH,
    XRT_HEIGHT,
    XRT_TEXT_WIDTH,
    XRT_TEXT_HEIGHT,
    XRT_TEXT_BORDER_WIDTH,

    /* int */
    XRT_XPRECISION                     = XRT_PROPERTY(INT, 0),
    XRT_YPRECISION,
    XRT_Y2PRECISION,
    XRT_PIE_MIN_SLICES,
    XRT_XMARKER_SET,
    XRT_XMARKER_POINT,
    XRT_BAR_CLUSTER_OVERLAP,
    XRT_BAR_CLUSTER_WIDTH,
    XRT_GRAPH_DEPTH,
    XRT_GRAPH_ROTATION,
    XRT_GRAPH_INCLINATION,
    XRT_TEXT_ATTACH_SET,
    XRT_TEXT_ATTACH_POINT,
    XRT_TEXT_OFFSET,

    /* long */
    XRT_TIME_BASE                       = XRT_PROPERTY(LONG, 0),
	XRT_CHART_OPTIONS,

    /* color */
    XRT_BACKGROUND_COLOR                = XRT_PROPERTY(COLOR, 0),
    XRT_FOREGROUND_COLOR,
    XRT_GRAPH_BACKGROUND_COLOR,
    XRT_GRAPH_FOREGROUND_COLOR,
    XRT_HEADER_BACKGROUND_COLOR,
    XRT_HEADER_FOREGROUND_COLOR,
    XRT_FOOTER_BACKGROUND_COLOR,
    XRT_FOOTER_FOREGROUND_COLOR,
    XRT_LEGEND_BACKGROUND_COLOR,
    XRT_LEGEND_FOREGROUND_COLOR,
    XRT_DATA_AREA_BACKGROUND_COLOR,
    XRT_DATA_AREA_FOREGROUND_COLOR,
    XRT_TEXT_BACKGROUND_COLOR,
    XRT_TEXT_FOREGROUND_COLOR,

    /* enum TYPE */
    XRT_TYPE                            = XRT_PROPERTY(TYPE, 0),
    XRT_TYPE2,

    /* enum ADJUST */
    XRT_HEADER_ADJUST                   = XRT_PROPERTY(ADJUST, 0),
    XRT_FOOTER_ADJUST,
    XRT_TEXT_ADJUST,

    /* enum PIE_ORDER */
    XRT_PIE_ORDER                       = XRT_PROPERTY(PIE_ORDER, 0),

    /* enum PIE_METHOD */
    XRT_PIE_THRESHOLD_METHOD            = XRT_PROPERTY(PIE_THRESHOLD_METHOD, 0),

    /* enum ANCHOR */
    XRT_LEGEND_ANCHOR                   = XRT_PROPERTY(ANCHOR, 0),
    XRT_TEXT_ANCHOR,

    /* enum ALIGN */
    XRT_LEGEND_ORIENTATION              = XRT_PROPERTY(ALIGN, 0),

    /* enum BORDER */
    XRT_BORDER                          = XRT_PROPERTY(BORDER, 0),
    XRT_HEADER_BORDER,
    XRT_FOOTER_BORDER,
    XRT_LEGEND_BORDER,
    XRT_GRAPH_BORDER,
    XRT_TEXT_BORDER,

    /* enum ANNO */
    XRT_XANNOTATION_METHOD              = XRT_PROPERTY(ANNO_METHOD, 0),
    XRT_YANNOTATION_METHOD,
    XRT_Y2ANNOTATION_METHOD,

    /* enum PLACEMENT */
    XRT_XANNO_PLACEMENT                 = XRT_PROPERTY(ANNO_PLACEMENT, 0),
    XRT_YANNO_PLACEMENT,
    XRT_Y2ANNO_PLACEMENT,

    /* enum NUM */
    XRT_XNUM_METHOD                     = XRT_PROPERTY(NUM_METHOD, 0),
    XRT_YNUM_METHOD,
    XRT_Y2NUM_METHOD,

    /* enum ORIGIN */
    XRT_XORIGIN_PLACEMENT               = XRT_PROPERTY(ORIGIN_PLACEMENT, 0),
    XRT_YORIGIN_PLACEMENT,
    XRT_Y2ORIGIN_PLACEMENT,

    /* enum TIMEUNIT */
    XRT_TIME_UNIT                       = XRT_PROPERTY(TIME_UNIT, 0),

    /* enum ROTATE */
    XRT_XTITLE_ROTATION                 = XRT_PROPERTY(ROTATE, 0),
    XRT_YTITLE_ROTATION,
    XRT_Y2TITLE_ROTATION,
    XRT_XANNOTATION_ROTATION,
    XRT_YANNOTATION_ROTATION,
    XRT_Y2ANNOTATION_ROTATION,
    XRT_TEXT_ROTATION,

    /* enum DATASET */
    XRT_FRONT_DATASET                   = XRT_PROPERTY(DATASET, 0),
    XRT_MARKER_DATASET,
    XRT_TEXT_ATTACH_DATASET,

    /* enum MARKER_METHOD */
    XRT_XMARKER_METHOD                  = XRT_PROPERTY(MARKER_METHOD, 0),

    /* enum TEXT_ATTACH */
    XRT_TEXT_ATTACH_TYPE                = XRT_PROPERTY(TEXT_ATTACH, 0),

    /* enum SHADING */
    XRT_GRAPH_3D_SHADING                = XRT_PROPERTY(SHADING, 0),

    /* enum ANGLE_UNIT */
    XRT_ANGLE_UNIT                      = XRT_PROPERTY(ANGLE_UNIT, 0),

    /* enum IMAGE_LAYOUT */
    XRT_IMAGE_LAYOUT                    = XRT_PROPERTY(IMAGE_LAYOUT, 0),
    XRT_GRAPH_IMAGE_LAYOUT,
    XRT_HEADER_IMAGE_LAYOUT,
    XRT_FOOTER_IMAGE_LAYOUT,
    XRT_LEGEND_IMAGE_LAYOUT,
    XRT_DATA_AREA_IMAGE_LAYOUT,
    XRT_TEXT_IMAGE_LAYOUT,

    /* enum BUBBLE_METHOD */
    XRT_BUBBLE_METHOD                   = XRT_PROPERTY(BUBBLE_METHOD, 0),

    /* string */
    XRT_XTITLE                          = XRT_PROPERTY(STRING, 0),
    XRT_YTITLE,
    XRT_Y2TITLE,
    XRT_OTHER_LABEL,
    XRT_TIME_FORMAT,
    XRT_NAME,
	XRT_XAXIS_LABEL_FORMAT,
	XRT_YAXIS_LABEL_FORMAT,
	XRT_Y2AXIS_LABEL_FORMAT,
    XRT_TEXT_NAME,
    XRT_IMAGE_OBSOLETE,
    XRT_GRAPH_IMAGE_OBSOLETE,
    XRT_HEADER_IMAGE_OBSOLETE,
    XRT_FOOTER_IMAGE_OBSOLETE,
    XRT_LEGEND_IMAGE_OBSOLETE,
    XRT_DATA_AREA_IMAGE_OBSOLETE,
    XRT_TEXT_IMAGE_OBSOLETE,
    XRT_LEGEND_TITLE,

    /* DATA */
    XRT_DATA                            = XRT_PROPERTY(DATA, 0),
    XRT_DATA2,

    /* FONT */
    XRT_AXIS_FONT                       = XRT_PROPERTY(FONT, 0),
    XRT_HEADER_FONT,
    XRT_FOOTER_FONT,
    XRT_LEGEND_FONT,
    XRT_TEXT_FONT,
    XRT_AXIS_TITLE_FONT,

    /* DATA_STYLE */
    XRT_OTHER_DATA_STYLE                = XRT_PROPERTY(DATA_STYLE, 0),
    XRT_MARKER_DATA_STYLE,
    XRT_XGRID_DATA_STYLE,
    XRT_YGRID_DATA_STYLE,
	XRT_Y2GRID_DATA_STYLE,
    XRT_XMARKER_DATA_STYLE,
    XRT_YMARKER_DATA_STYLE,
    XRT_XAXIS_DATA_STYLE,
    XRT_YAXIS_DATA_STYLE,
	XRT_Y2AXIS_DATA_STYLE,

    /* double   */
    XRT_XMIN                            = XRT_PROPERTY(FLOAT, 0),
    XRT_YMIN,
    XRT_Y2MIN,
    XRT_XMAX,
    XRT_YMAX,
    XRT_Y2MAX,
    XRT_XNUM,
    XRT_YNUM,
    XRT_Y2NUM,
    XRT_XTICK,
    XRT_YTICK,
    XRT_Y2TICK,
    XRT_XORIGIN,
    XRT_YORIGIN,
    XRT_PIE_THRESHOLD_VALUE,
    XRT_XMARKER,
    XRT_YMARKER,
    XRT_XGRID,
    XRT_YGRID,
    XRT_XAXIS_MAX,
    XRT_YAXIS_MAX,
    XRT_Y2AXIS_MAX,
    XRT_XAXIS_MIN,
    XRT_YAXIS_MIN,
    XRT_Y2AXIS_MIN,
    XRT_YAXIS_MULT,
    XRT_YAXIS_CONST,
	XRT_Y2ORIGIN,
	XRT_Y2GRID,
    XRT_TEXT_ATTACH_VALUE_X,
    XRT_TEXT_ATTACH_VALUE_Y,
    XRT_XORIGIN_BASE,
    XRT_YANNOTATION_ANGLE,
    XRT_Y2ANNOTATION_ANGLE,
    XRT_BUBBLE_MIN,
    XRT_BUBBLE_MAX,
    XRT_PIE_START_ANGLE,
    XRT_XANNOTATION_ROTATION_ANGLE,
    XRT_YANNOTATION_ROTATION_ANGLE,
    XRT_Y2ANNOTATION_ROTATION_ANGLE,
	XRT_TIME_BASE_OLE,

    /* Data styles array */
    XRT_DATA_STYLES                     = XRT_PROPERTY(DATA_STYLES, 0),
    XRT_DATA_STYLES2,

    /* String array */
    XRT_SET_LABELS                      = XRT_PROPERTY(STRINGS, 0),
    XRT_SET_LABELS2,
    XRT_POINT_LABELS,
    XRT_POINT_LABELS2,
    XRT_HEADER_STRINGS,
    XRT_FOOTER_STRINGS,
    XRT_TEXT_STRINGS,

    /* Value labels */
    XRT_XLABELS                         = XRT_PROPERTY(VALUE_LABELS, 0),
    XRT_YLABELS,
    XRT_Y2LABELS,

	/* XrtImage */
    XRT_IMAGE                           = XRT_PROPERTY(IMAGE, 0),
    XRT_DATA_AREA_IMAGE,
    XRT_FOOTER_IMAGE,
    XRT_GRAPH_IMAGE,
    XRT_HEADER_IMAGE,
    XRT_LEGEND_IMAGE,
    XRT_TEXT_IMAGE,
    XRT_SYMBOL_IMAGE,
	/* Functions */

    /* put use_default resources last */
    XRT_GRAPH_X_USE_DEFAULT             = XRT_PROPERTY(USE_DEFAULT, 0),
    XRT_GRAPH_Y_USE_DEFAULT,
    XRT_GRAPH_WIDTH_USE_DEFAULT,
    XRT_GRAPH_HEIGHT_USE_DEFAULT,
    XRT_HEADER_X_USE_DEFAULT,
    XRT_HEADER_Y_USE_DEFAULT,
    XRT_FOOTER_X_USE_DEFAULT,
    XRT_FOOTER_Y_USE_DEFAULT,
    XRT_LEGEND_X_USE_DEFAULT,
    XRT_LEGEND_Y_USE_DEFAULT,
    XRT_XMAX_USE_DEFAULT,
    XRT_YMAX_USE_DEFAULT,
    XRT_Y2MAX_USE_DEFAULT,
    XRT_XMIN_USE_DEFAULT,
    XRT_YMIN_USE_DEFAULT,
    XRT_Y2MIN_USE_DEFAULT,
    XRT_XTICK_USE_DEFAULT,
    XRT_YTICK_USE_DEFAULT,
    XRT_Y2TICK_USE_DEFAULT,
    XRT_XNUM_USE_DEFAULT,
    XRT_YNUM_USE_DEFAULT,
    XRT_Y2NUM_USE_DEFAULT,
    XRT_XORIGIN_USE_DEFAULT,
    XRT_YORIGIN_USE_DEFAULT,
    XRT_DATA_STYLES_USE_DEFAULT,
    XRT_DATA_STYLES2_USE_DEFAULT,
    XRT_OTHER_DATA_STYLE_USE_DEFAULT,
    XRT_MARKER_DATA_STYLE_USE_DEFAULT,
    XRT_XGRID_DATA_STYLE_USE_DEFAULT,
    XRT_YGRID_DATA_STYLE_USE_DEFAULT,
    XRT_XGRID_USE_DEFAULT,
    XRT_YGRID_USE_DEFAULT,
    XRT_XPRECISION_USE_DEFAULT,
    XRT_YPRECISION_USE_DEFAULT,
    XRT_Y2PRECISION_USE_DEFAULT,
    XRT_GRAPH_MARGIN_BOTTOM_USE_DEFAULT,
    XRT_GRAPH_MARGIN_LEFT_USE_DEFAULT,
    XRT_GRAPH_MARGIN_RIGHT_USE_DEFAULT,
    XRT_GRAPH_MARGIN_TOP_USE_DEFAULT,
    XRT_XAXIS_MAX_USE_DEFAULT,
    XRT_YAXIS_MAX_USE_DEFAULT,
    XRT_Y2AXIS_MAX_USE_DEFAULT,
    XRT_XAXIS_MIN_USE_DEFAULT,
    XRT_YAXIS_MIN_USE_DEFAULT,
    XRT_Y2AXIS_MIN_USE_DEFAULT,
    XRT_TIME_FORMAT_USE_DEFAULT,
    XRT_XMARKER_DATA_STYLE_USE_DEFAULT,
    XRT_YMARKER_DATA_STYLE_USE_DEFAULT,
	XRT_Y2ORIGIN_USE_DEFAULT,
	XRT_Y2GRID_USE_DEFAULT,
	XRT_Y2GRID_DATA_STYLE_USE_DEFAULT,
    XRT_YANNOTATION_ANGLE_USE_DEFAULT,
    XRT_Y2ANNOTATION_ANGLE_USE_DEFAULT,
    XRT_XAXIS_DATA_STYLE_USE_DEFAULT,
    XRT_YAXIS_DATA_STYLE_USE_DEFAULT,
	XRT_Y2AXIS_DATA_STYLE_USE_DEFAULT,
} XrtProperty;

typedef enum
{
    /* int */
    XRT_POINTSTYLE_SET                      = XRT_PROPERTY(INT, 0),
    XRT_POINTSTYLE_POINT,

    /* color */
    XRT_POINTSTYLE_PATTERN_BACKGROUND_COLOR = XRT_PROPERTY(COLOR, 0),

    /* enum DATASET */
    XRT_POINTSTYLE_DATASET                  = XRT_PROPERTY(DATASET, 0),

    /* DISPLAY */
    XRT_POINTSTYLE_DISPLAY                  = XRT_PROPERTY(DISPLAY, 0),

    /* DATA_STYLE */
    XRT_POINTSTYLE_DATA_STYLE               = XRT_PROPERTY(DATA_STYLE, 0),
    XRT_POINTSTYLE_FILL_STYLE,
    XRT_POINTSTYLE_LINE_STYLE,
    XRT_POINTSTYLE_SYMBOL_STYLE,

    /* double   */
    XRT_POINTSTYLE_SLICE_OFFSET             = XRT_PROPERTY(FLOAT, 0),

    /* image   */
    XRT_POINTSTYLE_IMAGE                    = XRT_PROPERTY(IMAGE, 0),
    XRT_POINTSTYLE_IMAGE_LAYOUT             = XRT_PROPERTY(IMAGE_LAYOUT, 0),
    XRT_POINTSTYLE_IMAGE_TRANSPARENT        = XRT_PROPERTY(BOOLVAL, 0),
    /* put use_default resources last */
    XRT_POINTSTYLE_USE_DEFAULT              = XRT_PROPERTY(USE_DEFAULT, 0),
    XRT_POINTSTYLE_DATA_STYLE_USE_DEFAULT,
    XRT_POINTSTYLE_FILL_STYLE_USE_DEFAULT,
    XRT_POINTSTYLE_LINE_STYLE_USE_DEFAULT,
    XRT_POINTSTYLE_SYMBOL_STYLE_USE_DEFAULT,
    XRT_POINTSTYLE_SLICE_OFFSET_USE_DEFAULT
} XrtPointStyleProperty;

/* notifications sent back to parent */
#define XRTN_RESIZED                        (WM_USER + 3000)
#define XRTN_REPAINTED                      (WM_USER + 3001)
#define XRTN_PALETTECHANGED                 (WM_USER + 3002)
#define XRTN_MODIFY_START                   (WM_USER + 3003)
#define XRTN_MODIFY_END                     (WM_USER + 3004)
#define XRTN_ROTATE                         (WM_USER + 3005)
#define XRTN_TRANSFORM                      (WM_USER + 3006)
#define XRTN_PROPERTIES                     (WM_USER + 3007)
#define XRTN_ZOOM_AXIS                      (WM_USER + 3008)
#define XRTN_FORMAT_AXIS_ANNO               (WM_USER + 3009)
#define XRTN_ZOOM_TRANSFORM_RESET           (WM_USER + 3010)

typedef enum {
    XRT_ACTION_NONE,
    XRT_ACTION_MODIFY_START,
    XRT_ACTION_MODIFY_END,
    XRT_ACTION_ROTATE,
    XRT_ACTION_SCALE,
    XRT_ACTION_TRANSLATE,
    XRT_ACTION_ZOOM_START,
    XRT_ACTION_ZOOM_UPDATE,
    XRT_ACTION_ZOOM_END,
    XRT_ACTION_ZOOM_CANCEL,
    XRT_ACTION_RESET,
    XRT_ACTION_PROPERTIES,
    XRT_ACTION_ZOOM_AXIS
} XrtAction;

typedef struct tag_XrtActionItem {
    UINT        msg;
    UINT        modifier;
    UINT        keycode;
    XrtAction   action;
    struct tag_XrtActionItem *next;
} XrtActionItem;

#ifdef NDEBUG
#define XRT_CHECK()
#else
#define XRT_CHECK() xrt_check(_T(__FILE__), __LINE__)
#endif

typedef int     XrtFocus;
typedef int     XrtDsGroup;

/*---------------------------------------------------------------
 *
 *  Enumerated types used throughout XRT/graph.
 *
 *---------------------------------------------------------------
*/

typedef enum {
    XRT_ADJUST_LEFT = 1,
    XRT_ADJUST_RIGHT,
    XRT_ADJUST_CENTER
} XrtAdjust;

typedef enum {
    XRT_ALIGN_VERTICAL = 1,
    XRT_ALIGN_HORIZONTAL
} XrtAlign;

typedef enum {
	XRT_ALT_DS_ALL          = -1,
	XRT_ALT_DS_ALL_DATA     = -2,
	XRT_ALT_DS_LEGEND       = -3,
	XRT_ALT_DS_OTHER_SLICE  = -4,
	XRT_ALT_DONT_USE        = -5
} XrtAltStyleMethod;

typedef enum {
    XRT_ANCHOR_NORTH        = 0x10,
    XRT_ANCHOR_SOUTH        = 0x20,
    XRT_ANCHOR_EAST         = 0x01,
    XRT_ANCHOR_WEST         = 0x02,
    XRT_ANCHOR_NORTHEAST    = 0x11,
    XRT_ANCHOR_NORTHWEST    = 0x12,
    XRT_ANCHOR_SOUTHEAST    = 0x21,
    XRT_ANCHOR_SOUTHWEST    = 0x22,
    /* for text areas only, not for legend */
    XRT_ANCHOR_HOME         = 0x00,
    XRT_ANCHOR_BEST         = 0x100
} XrtAnchor;

typedef enum {
    XRT_ANGLE_DEGREES = 1,
    XRT_ANGLE_RADIANS,
    XRT_ANGLE_GRADS
} XrtAngleUnit;

typedef enum {
    XRT_ANNO_VALUES = 0,
    XRT_ANNO_POINT_LABELS,
    XRT_ANNO_VALUE_LABELS,
    XRT_ANNO_TIME_LABELS,
    XRT_ANNO_VALUES_EVENT,
    XRT_ANNO_TIME_LABELS_EVENT
} XrtAnnoMethod;

typedef enum {
    XRT_ANNO_AUTO = 0,
    XRT_ANNO_ORIGIN,
    XRT_ANNO_MIN,
    XRT_ANNO_MAX
} XrtAnnoPlacement;

typedef enum {
    XRT_TEXT_ATTACH_PIXEL = 0,
    XRT_TEXT_ATTACH_VALUE,
    XRT_TEXT_ATTACH_DATA,
    XRT_TEXT_ATTACH_DATA_VALUE
} XrtAttachType;

typedef enum {
    XRT_AXIS_X = 0,
    XRT_AXIS_Y,
    XRT_AXIS_Y2
} XrtAxis;

typedef enum {
    XRT_BORDER_NONE = 0,
    XRT_BORDER_3D_OUT,
    XRT_BORDER_3D_IN,
    XRT_BORDER_SHADOW,
    XRT_BORDER_PLAIN,
    XRT_BORDER_ETCHED_IN,
    XRT_BORDER_ETCHED_OUT,
	XRT_BORDER_FRAME_IN,
	XRT_BORDER_FRAME_OUT,
	XRT_BORDER_BEVEL
} XrtBorder;

typedef enum {
    XRT_BUBBLEMETHOD_DIAMETER = 1,
    XRT_BUBBLEMETHOD_AREA
} XrtBubbleMethod;

/**
 * Possible data types stored in and XrtDataHandle.
 * For more information on data types, refer to the page \ref Data.
 */
typedef enum {
    XRT_DATA_ARRAY = 1,			/**< Type is XrtArray */
    XRT_DATA_GENERAL			/**< Type is XrtGeneral */
} XrtDataType;

typedef enum {
	XRT_DISPLAY_SHOW = 1,
	XRT_DISPLAY_HIDE,
	XRT_DISPLAY_EXCLUDE
} XrtDisplay;

/**
 * Supported fill patterns.
 */
typedef enum {
    XRT_FPAT_NONE = 1,
    XRT_FPAT_SOLID,
    XRT_FPAT_25_PERCENT,
    XRT_FPAT_50_PERCENT,
    XRT_FPAT_75_PERCENT,
    XRT_FPAT_HORIZ_STRIPE,
    XRT_FPAT_VERT_STRIPE,
    XRT_FPAT_45_STRIPE,
    XRT_FPAT_135_STRIPE,
    XRT_FPAT_DIAG_HATCHED,
    XRT_FPAT_CROSS_HATCHED
    ,
    XRT_WFPAT_BDIAGONAL,
    XRT_WFPAT_CROSS,
    XRT_WFPAT_DIAGCROSS,
    XRT_WFPAT_FDIAGONAL,
    XRT_WFPAT_HORIZONTAL,
    XRT_WFPAT_VERTICAL
} XrtFillPattern;

/**
 * Supported image layout and alignment options.
 */
typedef enum {
    XRT_IMAGELAYOUT_CENTERED = 1,
    XRT_IMAGELAYOUT_TILED
    ,
    XRT_IMAGELAYOUT_FITTED,
    XRT_IMAGELAYOUT_STRETCHED,
    XRT_IMAGELAYOUT_STRETCHED_TO_WIDTH,
    XRT_IMAGELAYOUT_STRETCHED_TO_HEIGHT,
    XRT_IMAGELAYOUT_CROP_FITTED
} XrtImageLayout;

/**
 * Possible line styles for the graph.
 */
typedef enum {
    XRT_LPAT_NONE = 1,
    XRT_LPAT_SOLID,
    XRT_LPAT_LONG_DASH,
    XRT_LPAT_DOTTED,
    XRT_LPAT_SHORT_DASH,
    XRT_LPAT_LSL_DASH,
    XRT_LPAT_DASH_DOT
} XrtLinePattern;

typedef enum {
    XRT_MARKER_SET_POINT = 1,
    XRT_MARKER_VALUE
} XrtMarkerMethod;

typedef enum {
    XRT_NUM_PRECISION = 0,
    XRT_NUM_ROUND
} XrtNumMethod;

/**
 * Possibble settinsgs for the origin placement.
 */
typedef enum {
    XRT_ORIGIN_AUTO = 0,
    XRT_ORIGIN_ZERO,
    XRT_ORIGIN_MIN,
    XRT_ORIGIN_MAX
} XrtOriginPlacement;

typedef enum {
    XRT_PIEORDER_ASCENDING = 1,
    XRT_PIEORDER_DESCENDING,
    XRT_PIEORDER_DATA_ORDER
} XrtPieOrder;

typedef enum {
    XRT_PIE_SLICE_CUTOFF = 1,
    XRT_PIE_PERCENTILE
} XrtPieThresholdMethod;

/**
 * List of supported point styles.
 */
typedef enum {
    XRT_POINT_NONE = 1,
    XRT_POINT_DOT,
    XRT_POINT_BOX,
    XRT_POINT_TRI,
    XRT_POINT_DIAMOND,
    XRT_POINT_STAR,
    XRT_POINT_VERT_LINE,
    XRT_POINT_HORIZ_LINE,
    XRT_POINT_CROSS,
    XRT_POINT_CIRCLE,
    XRT_POINT_SQUARE,
    XRT_POINT_INVERT_TRI,
    XRT_POINT_DIAG_CROSS,
    XRT_POINT_OPEN_TRI,
    XRT_POINT_OPEN_DIAMOND,
    XRT_POINT_OPEN_INVERT_TRI,
    XRT_POINT_IMAGE_FILE
} XrtPoint;

/* stay out of range of Motif reasons */

typedef enum {
    XRT_RGN_IN_GRAPH = -100,
    XRT_RGN_IN_LEGEND,
    XRT_RGN_IN_HEADER,
    XRT_RGN_IN_FOOTER,
#ifdef ENHANCED_PICK
    XRT_RGN_IN_AXIS,
    XRT_RGN_IN_LABEL,
    XRT_RGN_IN_TITLE,
#endif
    XRT_RGN_NOWHERE
} XrtRegion;

/**
 * List of supported rotation values.
 */
typedef enum {
    XRT_ROTATE_NONE = 1,
    XRT_ROTATE_90,
    XRT_ROTATE_270,
    XRT_ROTATE_45,
    XRT_ROTATE_315,
    XRT_ROTATE_OTHER
} XrtRotate;

typedef enum {
    XRT_SHADING_COLOR = 1,
    XRT_SHADING_DITHERED,
    XRT_SHADING_NONE
} XrtShading;

typedef enum {
    XRT_TMUNIT_SECONDS = 1,
    XRT_TMUNIT_MINUTES,
    XRT_TMUNIT_HOURS,
    XRT_TMUNIT_DAYS,
    XRT_TMUNIT_WEEKS,
    XRT_TMUNIT_MONTHS,
    XRT_TMUNIT_YEARS
} XrtTimeUnit;

/** Types of basic graphs supported */
typedef enum {
    XRT_TYPE_PLOT = 1,
    XRT_TYPE_BAR,
    XRT_TYPE_PIE,
    XRT_TYPE_STACKING_BAR,
    XRT_TYPE_AREA,
	XRT_TYPE_HILO,
	XRT_TYPE_HILO_OPEN_CLOSE,
	XRT_TYPE_CANDLE,
	XRT_TYPE_POLAR,
	XRT_TYPE_RADAR,
    XRT_TYPE_FILLED_RADAR,
    XRT_TYPE_BUBBLE
} XrtType;

typedef enum {
    XRT_XMETHOD_XVALUES = 0,
    XRT_XMETHOD_POINT_LABELS,
    XRT_XMETHOD_XLABELS,
    XRT_XMETHOD_TIME_LABELS
} XrtXMethod;


/** \page Data XRT/graph Data Definitions
 **
 *  Theses are the basic types of data that can be attached to
 *  an XRT graph.
 *
 *  \arg ARRAY        This defines a sequence of n sets, each having
 *                    common \c x values; each set consists of \c m points.
 *                    The \c x values are stored as an array of \c m points,
 *                    while the \c y values are stored as an array \c m vectors.
 *                    Each vector contains \c n values representing
 *                    the \c y values for each set.
 *
 *  \arg GENERAL      This defines a set of lines \c 1, \c 2, ..., \c n, with
 *                    each set containing \c m1, \c m2, ..., \c mn points
 *                    respectively.  The data is stored as an array of
 *                    \c n sets, with each set consisting of a structure
 *                    containing the number of points, and a pair of
 *                    vectors for the \c x and \c y values.
 *
 *  Each of the structures (XrtArray and XrtGeneral) define the structures which
 *  are used to define each type of graph.  The first element in each structure
 *  defines the type of data; this, combined with placing the structures
 *  in a union, allows us to extract the type of data from an opaque
 *  handle to the data.
 *
 */

/**
 * Returns the type of data contained in the structure. The structure  will be
 * either XrtArray or XrtGeneral. This corresponds to macro values of XRT_DATA_ARRAY
 * or XRT_DATA_GENERAL, respectively.
 */
#define XrtGetDataType(xrt_data)        (xrt_data->a.type)

/** The data values in an XrtArray struct */
typedef struct {
    double           *xp;
    double          **yp;
} XrtArrayData;

/**
 * Structure defining array data.
 * Refer to \ref Data for a description of the possible data types.
 */
typedef struct {
    XrtDataType       type;         /**< = XRT_DATA_ARRAY */
    double            hole;
    int               nsets;
    int               npoints;
    XrtArrayData      data;
} XrtArray;

/** This is one set of general data */
typedef struct {
    int               npoints;
    double           *xp;
    double           *yp;
} XrtGeneralData;

typedef struct {
    XrtDataType       type;         /**< = XRT_DATA_GENERAL */
    double            hole;			/**< value to use as the HOLE_VALUE */
    int               nsets;		/**< number of sets in the data. */
    XrtGeneralData   *data;			/**< Pointer to the list of XrtGeneralData structures */
} XrtGeneral;

typedef union {
    XrtArray          a;
    XrtGeneral        g;
} XrtData;

typedef struct {
    int         pix_x, pix_y;
    int         dataset;
    int         set, point;
    int         distance;
} XrtPickResult;

typedef struct {
    int         pix_x, pix_y;
    int         yaxis;
    double      x, y;
} XrtMapResult;

/** XLABELS structure */
typedef struct {
    double      xvalue;
    TCHAR       *string;
} XrtXLabel;

/** Value Label structure */
typedef struct {
    double      value;
    TCHAR       *string;
    XrtColor    color;
} XrtValueLabel;

typedef union {
    struct {
        XrtAttachType       type;
        int                 x, y;
    } pixel;
    struct {
        XrtAttachType       type;
        int                 dataset;
        double              x, y;
    } value;
    struct {
        XrtAttachType       type;
        int                 dataset;
        int                 set, point;
    } data;
    struct {
        XrtAttachType       type;
        int                 dataset;
        int                 set, point;
        double              y;
    } data_value;
} XrtTextPosition;

/* Callback structures */

typedef struct {
    HDC              hdc;           /* read only */
    RECT             rectDamaged;   /* read only */
} XrtCallbackStruct;

typedef struct  {
    XrtDimension     width;          /* read only */
    XrtDimension     height;         /* read only */
} XrtResizeCallbackStruct;

typedef struct  {
    XrtPosition      x;              /* read only */
    XrtPosition      y;              /* read only */
} XrtPropertiesCallbackStruct;

typedef struct {

    int              rotation;
    int              inclination;
    XrtBoolean       doit;
} XrtRotateCallbackStruct;

typedef struct {
    XrtBoolean       reset;         /* read only */
    XrtPosition      left_margin;
    XrtPosition      right_margin;
    XrtPosition      top_margin;
    XrtPosition      bottom_margin;
    XrtBoolean       doit;
} XrtTransformCallbackStruct;

typedef struct {
    XrtBoolean       reset;         /* read only */
    double           new_xaxis_min;
    double           new_xaxis_max;
    double           new_yaxis_min;
    double           new_yaxis_max;
    double           new_y2axis_min;
    double           new_y2axis_max;
    XrtBoolean       doit;
} XrtZoomAxisCallbackStruct;

typedef struct {
    XrtBoolean       doit;
} XrtModifyCallbackStruct;

typedef struct {
    XrtRegion        region;        /**< read only */
    XrtMapResult     map;           /**< read only */
} XrtMapCallbackStruct;

typedef struct {
    XrtRegion        region;        /**< read only */
    XrtPickResult    pick;          /**< read only */
} XrtPickCallbackStruct;


typedef enum {
    XRT_DRAW_BITMAP = 1,
    XRT_DRAW_METAFILE,
	XRT_DRAW_ENHMETAFILE,
	XRT_DRAW_STANDARDMETAFILE
} XrtDrawFormat;

typedef enum {
    XRT_DRAWSCALE_NONE = 1,
    XRT_DRAWSCALE_TOWIDTH,
    XRT_DRAWSCALE_TOHEIGHT,
    XRT_DRAWSCALE_TOFIT,
    XRT_DRAWSCALE_TOMAX
} XrtDrawScale;

typedef enum {
    XRT_DRAWTYPE_DC = 1,
    XRT_DRAWTYPE_PRINTER,
    XRT_DRAWTYPE_CLIPBOARD,
    XRT_DRAWTYPE_FILE,
} XrtDrawType;

typedef struct {
    XrtLinePattern    lpat;     /**< line pattern */
    XrtFillPattern    fpat;     /**< fill pattern */
    XrtColor          color;    /**< line color   */
    int               width;    /**< line width   */
    XrtPoint          point;    /**< point style  */
    XrtColor          pcolor;   /**< point color  */
    int               psize;    /**< point size - pixels */

    XrtColor          resh;     /* reserved */
    XrtColor          ress;     /* reserved */
} XrtDataStyle;

typedef struct {
	int           	  set;
	int           	  point;
	XrtDataStyle	 *datastyle;
} XrtAlternateDataStyle;

/** Attached text structure */
typedef struct {
    XrtTextPosition   position;
    TCHAR            **strings;
    XrtAnchor         anchor;
    int               offset;
    int               connected;
    XrtAdjust         adjust;
    XrtColor          fore_color;
    XrtColor          back_color;
    XrtBorder         border;
    int               border_width;
    HFONT             font;
    XrtRectangle      coords;       /* read-only */
} XrtTextDesc;

typedef struct tag_AlarmZoneCoord
{
    XrtFloat    upper_y;
    XrtFloat    lower_y;
} XrtAlarmZoneCoord;

typedef struct tag_AlarmZone
{
    XrtBoolean            is_showing;
    XrtAlarmZoneCoord     zonecoord;
    XrtColor              line_color;
    XrtColor              fill_color;
    XrtFillPattern        fill_pattern;
    TCHAR                 *name;
    struct tag_AlarmZone *next;
} XrtAlarmZone;

typedef struct {
	short v[4];
	TCHAR *ProductName;
	TCHAR *ProductVersion;
	TCHAR *Copyright;
	TCHAR *Copyright2;
} XrtVersionInfo;

/* Chart option constant */
#define XRT_CO_GRID_OVER_CHART 0x00000001

typedef struct {
    XrtValueLabel* label;
    XrtAxis axis;
} XrtFormatAxisAnnoCallbackStruct; 

# if defined(_WIN32)
#  include "poppack.h"
# else
#  if !defined(RC_INVOKED)
#   if defined(_MSC_VER)
#     pragma pack()
#   else
#    include "poppack.h"
#   endif
#  endif
# endif

#endif /* ~xrt_com_DEFINED */
