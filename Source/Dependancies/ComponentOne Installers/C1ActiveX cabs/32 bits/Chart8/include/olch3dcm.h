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

/** \file xrt3dcom.h
 *
 *  This include file is to contain all the definitions which are
 *  necessary for both the API level and the XRT implementation level;
 *  include files at both of these levels include this file.
 *
 */

#if !defined(LONG_PTR) && defined(_MSC_VER) && _MSC_VER < 1400
#define LONG_PTR long
#endif

#ifndef xrt3d_common_DEFINED
#define xrt3d_common_DEFINED

#define WIN_TOOLKIT
#define WIN_DRIVER

#ifndef RC_INVOKED
#include <math.h>
#endif

#ifdef __STDC__
# include <float.h>
# define XRT3D_HUGE_VAL						DBL_MAX
#else
# include <float.h>
/* DO NOT USE DBL_MAX!  VB can't handle it - this is big enough */
#  define XRT3D_HUGE_VAL					1.0E308
#endif

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

/*
 *  Version numbers and strings, in case developers need
 *  to have any version dependent code.  This scheme allows
 *  for 99 revisions, 9 point updates.  Also provide a
 *  string, for printing purposes.
*/

#define XRT3D_VERSION		3
#define XRT3D_REVISION		0
#define XRT3D_POINT			0
#define XRT3D_RELEASE		\
					(XRT3D_VERSION * 1000 + XRT3D_REVISION*10 + XRT3D_POINT)
#define XRT3D_RELEASE_STRING	_T("3.0.0")

#define XRT3D_DEFAULT_FLOAT				XRT3D_HUGE_VAL
#define XRT3D_DEFAULT_DISTN_LEVELS		10
#define XRT3D_DEFAULT_MESH_FILTER		1
#define XRT3D_DEFAULT_STROKE_SIZE       80      /* 8% of cube */
#define XRT3D_MAX_DISTN_LEVELS			100
#define XRT3D_MIN_PERSPECTIVE_DEPTH		1.0
#define XRT3D_MAX_VIEW_SCALE			30.0
#define XRT3D_MIN_VIEW_SCALE			(1.0 / XRT3D_MAX_VIEW_SCALE)

#define XRT3D_DEFAULT_SYMBOL_SIZE		7

#define XRT3D_MAX_BORDER_WIDTH			20

#define	XRT3D_XY_PLANE	1
#define XRT3D_XZ_PLANE	2
#define XRT3D_YZ_PLANE	4

#define	XRT3D_PROJECT_CONTOURS	1
#define	XRT3D_PROJECT_ZONES		2

/* Change statuses, for handling repaints */

#define XRT3D_GRAPH_CHANGED		0x01
#define XRT3D_HEADER_CHANGED	0x02
#define XRT3D_FOOTER_CHANGED	0x04
#define XRT3D_LEGEND_CHANGED	0x08
#define XRT3D_DATA_CHANGED		0x19		/**< data, graph, and legend */
#define XRT3D_ALL_CHANGED		0x1F		/**< all of the above        */


#ifndef MK_ALT
#define MK_ALT 0x20		/* Win32 defines this, Win16 doesn't */
#endif

 #ifdef _UNICODE
  #ifdef XRT_STATICLIB
   #define XRT3DWNDCLASS "C1Chart3DSU8"		/*NO UNICODE CONVERSION HERE*/
  #else
   #define XRT3DWNDCLASS "C1Chart3DU8"		/*NO UNICODE CONVERSION HERE*/
  #endif
 #else
  #ifdef XRT_STATICLIB
   #define XRT3DWNDCLASS "C1Chart3DS8"		/*NO UNICODE CONVERSION HERE*/
  #else
   #define XRT3DWNDCLASS "C1Chart3D8"		/*NO UNICODE CONVERSION HERE*/
  #endif
 #endif
 #define XRT3D _T(XRT3DWNDCLASS)

#if defined( STRICT )
typedef struct HXRT3D__ { int unused; } FAR* HXRT3D;
typedef struct HXRT3DTEXT__ { int unused; } FAR* HXRT3DTEXT;
#else
typedef void FAR* HXRT3D;
typedef void FAR* HXRT3DTEXT;
#endif

#define XRT3D_DEFAULT_COLOR	0xffffffff

# ifdef XRT_STATICLIB
#  define XRTEXPORT(type)   extern type
#  define XRTIMPORT(type)   extern type
#  define XRTEXP(type)      XRTEXPORT(type)
#  define XRTIMP(type)      XRTIMPORT(type)
# else /* !XRT_STATICLIB */
#  define XRTEXPORT(type)  extern __declspec( dllexport ) type
#  define XRTIMPORT(type)  extern __declspec( dllimport ) type
#  define XRTEXP(type)     XRTEXPORT(type)
#  define XRTIMP(type)     XRTIMPORT(type)
# endif /* !XRT_STATICLIB */



#ifndef XRT_TYPES_DEFINED
/* Portability section - use these typedefs on Windows or Motif */


typedef COLORREF 			XrtColor;			/* colors specified as COLORREF */
typedef BOOL				XrtBoolean;
typedef int		 			XrtDimension; 		/* 16-bit GDI in Win16, 32-bit GDI in Win32 */
typedef int					XrtPosition; 		/* 16-bit GDI in Win16, 32-bit GDI in Win32 */
typedef HFONT               XrtFont;
typedef double		XrtFloat;

#define XRT_TYPES_DEFINED
#endif /* XRT_TYPES_DEFINED */


#ifndef XRT3D_TYPES_DEFINED

typedef TCHAR *				Xrt3dImage;

#define XRT3D_TYPES_DEFINED
#endif /* XRT3D_TYPES_DEFINED */


/*
 *	All XRT3D_* properties are defined as a continuous sequence of numbers within each representation type
 */
#if ((defined(sun) && !defined(SOLARIS)) || defined(_DGUX_SOURCE)) && !defined(__STDC__) && !defined(__cplusplus)
#define XRT3D_PROPERTY(type, index)             ((unsigned int) (((XRT3DT_/**/type) << 8) | (index)))
#else
#define XRT3D_PROPERTY(type, index)             ((unsigned int) (((XRT3DT_##type) << 8) | (index)))
#endif
#define XRT3D_PROPERTY_TYPE(property)           ((((unsigned int) (property)) >> 8) & 0xFF)
#define XRT3D_PROPERTY_INDEX(property)          (((unsigned int) (property)) & 0xFF)

/**< \internal All types over ENUM_BASE are enums */
#define XRT3DT_ENUM_BASE      64      /* keep value positive : 128 introduces sign extension problems */
#define XRT3DT_IS_ENUM(type)  ((type) >= XRT3DT_ENUM_BASE)

typedef enum {
    XRT3DT_NONE = 0,    /* unused */
	XRT3DT_BOOLVAL,
	XRT3DT_COLOR,
	XRT3DT_CONTOUR_STYLES,
	XRT3DT_DIMENSION,
	XRT3DT_DISTN_TABLE,
	XRT3DT_DATA,
	XRT3DT_DOUBLE,
	XRT3DT_FONT,
	XRT3DT_FUNCTION,
    XRT3DT_GRIDLINES,
	XRT3DT_INT,
    XRT3DT_POSITION,
	XRT3DT_PROJECT,
	XRT3DT_STRING,
	XRT3DT_STRINGS,
	XRT3DT_VALUE_LABELS,
	XRT3DT_USE_DEFAULT,
	XRT3DT_XYCOLORS,
    XRT3DT_LINE_STYLE,
	XRT3DT_DATA_STYLES,
	XRT3DT_IMAGE,

    XRT3DT_ADJUST = XRT3DT_ENUM_BASE,
	XRT3DT_ALIGN,
	XRT3DT_ANCHOR,
	XRT3DT_ANNO_METHOD,
	XRT3DT_ATTACH_METHOD,
	XRT3DT_BAR_FORMAT,
	XRT3DT_BORDER,
	XRT3DT_DISTN_METHOD,
	XRT3DT_LEGEND_STYLE,
	XRT3DT_LINE_PATTERN,
	XRT3DT_PLANE,
	XRT3DT_PREVIEW_METHOD,
	XRT3DT_STROKE_FONT,
	XRT3DT_TYPE,
	XRT3DT_ZONE_METHOD,
	XRT3DT_TEXT_PLANE,
    XRT3DT_IMAGE_LAYOUT,
    XRT3DT_DISTN_RANGE,
	XRT3DT_SYMBOL_PATTERN,
	XRT3DT_FONT_ROTATION,
	XRT3DT_ANNO_POSITION,
} Xrt3dPropertyType;

/**< \internal last type in list */
#define XRT3DT_LAST_TYPE       XRT3DT_FONT_ROTATION

typedef enum {
    /* XrtBoolean */
    XRT3D_DEBUG                     = XRT3D_PROPERTY(BOOLVAL, 0),
    XRT3D_DOUBLE_BUFFER,
    XRT3D_DRAW_CONTOURS,
    XRT3D_DRAW_HIDDEN_LINES,
    XRT3D_DRAW_MESH,
    XRT3D_DRAW_SHADED,
    XRT3D_DRAW_ZONES,
    XRT3D_LEGEND_SHOW,
    XRT3D_REPAINT,
    XRT3D_SOLID_SURFACE,
    XRT3D_TEXT_LINE_SHOW,
    XRT3D_TEXT_SHOW,
    XRT3D_VIEW_NORMALIZED,
    XRT3D_XAXIS_SHOW,
    XRT3D_XMESH_SHOW,
    XRT3D_YAXIS_SHOW,
    XRT3D_YMESH_SHOW,
    XRT3D_ZAXIS_SHOW,
    XRT3D_TEXT_IMAGE_MINIMUM_SIZE,
    XRT3D_HEADER_IMAGE_MINIMUM_SIZE,
    XRT3D_FOOTER_IMAGE_MINIMUM_SIZE,
    XRT3D_USE_TRUETYPE,
	XRT3D_DRAW_DROP_LINES,

    XRT3D_IMAGE_TRANSPARENT,
    XRT3D_GRAPH_IMAGE_TRANSPARENT,
    XRT3D_HEADER_IMAGE_TRANSPARENT,
    XRT3D_FOOTER_IMAGE_TRANSPARENT,
    XRT3D_LEGEND_IMAGE_TRANSPARENT,
    XRT3D_TEXT_IMAGE_TRANSPARENT,

    /* XrtPosition */
    XRT3D_FOOTER_X                  = XRT3D_PROPERTY(POSITION, 0),
    XRT3D_FOOTER_Y,
    XRT3D_GRAPH_X,
    XRT3D_GRAPH_Y,
    XRT3D_HEADER_X,
    XRT3D_HEADER_Y,
    XRT3D_LEGEND_X,
    XRT3D_LEGEND_Y,
    XRT3D_TEXT_ATTACH_PIXEL_X,
    XRT3D_TEXT_ATTACH_PIXEL_Y,
    XRT3D_TEXT_OFFSET_X,
    XRT3D_TEXT_OFFSET_Y,

    /* XrtDimension */
    XRT3D_AXIS_STROKE_SIZE          = XRT3D_PROPERTY(DIMENSION, 0),
    XRT3D_AXIS_TITLE_STROKE_SIZE,
    XRT3D_BORDER_WIDTH,
    XRT3D_FOOTER_BORDER_WIDTH,
    XRT3D_FOOTER_HEIGHT,
    XRT3D_FOOTER_WIDTH,
    XRT3D_GRAPH_BORDER_WIDTH,
    XRT3D_GRAPH_HEIGHT,
    XRT3D_GRAPH_WIDTH,
    XRT3D_HEADER_BORDER_WIDTH,
    XRT3D_HEADER_HEIGHT,
    XRT3D_HEADER_WIDTH,
    XRT3D_LEGEND_BORDER_WIDTH,
    XRT3D_LEGEND_HEIGHT,
    XRT3D_LEGEND_WIDTH,
    XRT3D_TEXT_BORDER_WIDTH,
    XRT3D_TEXT_STROKE_SIZE,
    XRT3D_WIDTH,
    XRT3D_HEIGHT,

    /* int */
    XRT3D_NUM_DISTN_LEVELS          = XRT3D_PROPERTY(INT, 0),
    XRT3D_PROJECT_ZMAX,
    XRT3D_PROJECT_ZMIN,
    XRT3D_TEXT_ATTACH_INDEX_X,
    XRT3D_TEXT_ATTACH_INDEX_Y,
    XRT3D_XGRID_LINES,
    XRT3D_XMESH_FILTER,
    XRT3D_YGRID_LINES,
    XRT3D_YMESH_FILTER,
    XRT3D_ZGRID_LINES,
    XRT3D_LEGEND_LABEL_FILTER,
    XRT3D_TEXT_ATTACH_INDEX_POINT,
    XRT3D_TEXT_ATTACH_INDEX_SERIES,

    /* XrtColors */
    XRT3D_BACKGROUND_COLOR              = XRT3D_PROPERTY(COLOR, 0),
    XRT3D_DATA_AREA_BACKGROUND_COLOR,
    XRT3D_FOOTER_BACKGROUND_COLOR,
    XRT3D_FOOTER_FOREGROUND_COLOR,
    XRT3D_FOREGROUND_COLOR,
    XRT3D_GRAPH_BACKGROUND_COLOR,
    XRT3D_GRAPH_FOREGROUND_COLOR,
    XRT3D_HEADER_BACKGROUND_COLOR,
    XRT3D_HEADER_FOREGROUND_COLOR,
    XRT3D_LEGEND_BACKGROUND_COLOR,
    XRT3D_LEGEND_FOREGROUND_COLOR,
    XRT3D_MESH_BOTTOM_COLOR,
    XRT3D_MESH_TOP_COLOR,
    XRT3D_SURFACE_BOTTOM_COLOR,
    XRT3D_SURFACE_TOP_COLOR,
    XRT3D_TEXT_BACKGROUND_COLOR,
    XRT3D_TEXT_FOREGROUND_COLOR,

    /* enum Plane */
    XRT3D_TEXT_PLANE                    = XRT3D_PROPERTY(TEXT_PLANE, 0),

    /* enum StrokeFont */
    XRT3D_AXIS_STROKE_FONT              = XRT3D_PROPERTY(STROKE_FONT, 0),
    XRT3D_AXIS_TITLE_STROKE_FONT,
    XRT3D_TEXT_STROKE_FONT,

    /* enum BarFormat */
    XRT3D_XBAR_FORMAT                  = XRT3D_PROPERTY(BAR_FORMAT, 0),
    XRT3D_YBAR_FORMAT,

    /* enum DistnMethod */
    XRT3D_DISTN_METHOD                  = XRT3D_PROPERTY(DISTN_METHOD, 0),

    /* enum Adjust */
    XRT3D_FOOTER_ADJUST                 = XRT3D_PROPERTY(ADJUST, 0),
    XRT3D_HEADER_ADJUST,
    XRT3D_TEXT_ADJUST,

    /* enum Border */
    XRT3D_BORDER                        = XRT3D_PROPERTY(BORDER, 0),
    XRT3D_FOOTER_BORDER,
    XRT3D_GRAPH_BORDER,
    XRT3D_HEADER_BORDER,
    XRT3D_LEGEND_BORDER,
    XRT3D_TEXT_BORDER,

    /* enum Anchor */
    XRT3D_LEGEND_ANCHOR                 = XRT3D_PROPERTY(ANCHOR, 0),

    /* enum Align */
    XRT3D_LEGEND_ORIENTATION            = XRT3D_PROPERTY(ALIGN, 0),

    /* enum LegendStyle */
    XRT3D_LEGEND_STYLE                  = XRT3D_PROPERTY(LEGEND_STYLE, 0),

    /* enum AttachMethod */
    XRT3D_TEXT_ATTACH_METHOD            = XRT3D_PROPERTY(ATTACH_METHOD, 0),

    /* enum Type */
    XRT3D_TYPE                          = XRT3D_PROPERTY(TYPE, 0),

    /* enum AnnoMethod */
    XRT3D_XANNO_METHOD                 = XRT3D_PROPERTY(ANNO_METHOD, 0),
    XRT3D_YANNO_METHOD,
    XRT3D_ZANNO_METHOD,

    /* enum ZoneMethod */
    XRT3D_ZONE_METHOD                   = XRT3D_PROPERTY(ZONE_METHOD, 0),

    /* enum PreviewMethod */
    XRT3D_PREVIEW_METHOD                = XRT3D_PROPERTY(PREVIEW_METHOD, 0),

    /* enum IMAGE_LAYOUT */
    XRT3D_IMAGE_LAYOUT                  = XRT3D_PROPERTY(IMAGE_LAYOUT, 0),
    XRT3D_GRAPH_IMAGE_LAYOUT,
    XRT3D_HEADER_IMAGE_LAYOUT,
    XRT3D_FOOTER_IMAGE_LAYOUT,
    XRT3D_LEGEND_IMAGE_LAYOUT,
    XRT3D_TEXT_IMAGE_LAYOUT,

    /* enum DIST_RANGE */
    XRT3D_LEGEND_DISTN_RANGE           = XRT3D_PROPERTY(DISTN_RANGE, 0),

    XRT3D_FONT_ROTATION                = XRT3D_PROPERTY(FONT_ROTATION, 0),

    XRT3D_XANNO_POSITION                 = XRT3D_PROPERTY(ANNO_POSITION, 0),
    XRT3D_YANNO_POSITION,
    XRT3D_ZANNO_POSITION,

    /* string */
    XRT3D_XAXIS_TITLE                  = XRT3D_PROPERTY(STRING, 0),
    XRT3D_YAXIS_TITLE,
    XRT3D_ZAXIS_TITLE,
    XRT3D_TEXT_PRINT_FONT,
    XRT3D_NAME,
    XRT3D_IMAGE_OBSOLETE,
    XRT3D_GRAPH_IMAGE_OBSOLETE,
    XRT3D_HEADER_IMAGE_OBSOLETE,
    XRT3D_FOOTER_IMAGE_OBSOLETE,
    XRT3D_LEGEND_IMAGE_OBSOLETE,
    XRT3D_TEXT_IMAGE_OBSOLETE,
    XRT3D_LEGEND_TITLE,

    /* DATA */
    XRT3D_SURFACE_DATA                  = XRT3D_PROPERTY(DATA, 0),
    XRT3D_ZONE_DATA,

    /* FONT */
    XRT3D_FOOTER_FONT                   = XRT3D_PROPERTY(FONT, 0),
    XRT3D_HEADER_FONT,
    XRT3D_LEGEND_FONT,
    XRT3D_TEXT_FONT,
    XRT3D_AXIS_FONT,
    XRT3D_AXIS_TITLE_FONT,

    /* DISTN_TABLE */
    XRT3D_DISTN_TABLE                   = XRT3D_PROPERTY(DISTN_TABLE, 0),

    /* FUNCTION */
    XRT3D_LEGEND_LABEL_FUNC             = XRT3D_PROPERTY(FUNCTION, 0),

    /* Grid Line Styles */
    XRT3D_XGRID_LINE_STYLE              = XRT3D_PROPERTY(LINE_STYLE, 0),
    XRT3D_YGRID_LINE_STYLE,
    XRT3D_ZGRID_LINE_STYLE,

    /* double */
    XRT3D_XBAR_SPACING                  = XRT3D_PROPERTY(DOUBLE, 0),
    XRT3D_YBAR_SPACING,
    XRT3D_PERSPECTIVE_DEPTH,
    XRT3D_TEXT_ATTACH_POINT_X,
    XRT3D_TEXT_ATTACH_POINT_Y,
    XRT3D_TEXT_ATTACH_POINT_Z,
    XRT3D_VIEW_SCALE,
    XRT3D_VIEW_XTRANSLATE,
    XRT3D_VIEW_YTRANSLATE,
    XRT3D_XMAX,
    XRT3D_XMIN,
    XRT3D_XROTATION,
    XRT3D_XSCALE,
    XRT3D_YMAX,
    XRT3D_YMIN,
    XRT3D_YROTATION,
    XRT3D_YSCALE,
    XRT3D_ZMAX,
    XRT3D_ZMIN,
    XRT3D_ZORIGIN,
    XRT3D_ZROTATION,
    XRT3D_ZSCALE,

    /* ContourStyle array */
    XRT3D_CONTOUR_STYLES                = XRT3D_PROPERTY(CONTOUR_STYLES, 0),

    /* String array */
    XRT3D_FOOTER_STRINGS                = XRT3D_PROPERTY(STRINGS, 0),
    XRT3D_HEADER_STRINGS,
    XRT3D_LEGEND_STRINGS,
    XRT3D_TEXT_STRINGS,
    XRT3D_XDATA_LABELS,
    XRT3D_YDATA_LABELS,

    /* ValueLabel array */
    XRT3D_XVALUE_LABELS                 = XRT3D_PROPERTY(VALUE_LABELS, 0),
    XRT3D_YVALUE_LABELS,
    XRT3D_ZVALUE_LABELS,

    /* XYColor array */
    XRT3D_XY_COLORS                     = XRT3D_PROPERTY(XYCOLORS, 0),

	/* DataStyles */
	XRT3D_DATA_STYLES                   = XRT3D_PROPERTY(DATA_STYLES, 0),

	/* Xrt3dImage type */
    XRT3D_IMAGE                         = XRT3D_PROPERTY(IMAGE, 0),
    XRT3D_GRAPH_IMAGE,
    XRT3D_HEADER_IMAGE,
    XRT3D_FOOTER_IMAGE,
    XRT3D_LEGEND_IMAGE,
    XRT3D_TEXT_IMAGE,

    /* put use_default resources last */
    XRT3D_FOOTER_X_USE_DEFAULT          = XRT3D_PROPERTY(USE_DEFAULT, 0),
    XRT3D_FOOTER_Y_USE_DEFAULT,
    XRT3D_GRAPH_HEIGHT_USE_DEFAULT,
    XRT3D_GRAPH_WIDTH_USE_DEFAULT,
    XRT3D_GRAPH_X_USE_DEFAULT,
    XRT3D_GRAPH_Y_USE_DEFAULT,
    XRT3D_HEADER_X_USE_DEFAULT,
    XRT3D_HEADER_Y_USE_DEFAULT,
    XRT3D_LEGEND_X_USE_DEFAULT,
    XRT3D_LEGEND_Y_USE_DEFAULT,
    XRT3D_XMAX_USE_DEFAULT,
    XRT3D_XMIN_USE_DEFAULT,
    XRT3D_YMAX_USE_DEFAULT,
    XRT3D_YMIN_USE_DEFAULT,
    XRT3D_ZMAX_USE_DEFAULT,
    XRT3D_ZMIN_USE_DEFAULT,
	XRT3D_DATA_STYLES_USE_DEFAULT,
    XRT3D_XGRID_LINE_STYLE_USE_DEFAULT,
    XRT3D_YGRID_LINE_STYLE_USE_DEFAULT,
    XRT3D_ZGRID_LINE_STYLE_USE_DEFAULT
} Xrt3dProperty;

/* notifications sent back to parent */
#define XRT3DN_RESIZED						(WM_USER + 4000)
#define XRT3DN_REPAINTED					(WM_USER + 4001)
#define XRT3DN_PALETTECHANGED 				(WM_USER + 4002)
#define XRT3DN_MODIFY_START					(WM_USER + 4003)
#define XRT3DN_MODIFY_END					(WM_USER + 4004)
#define XRT3DN_ROTATE						(WM_USER + 4005)
#define XRT3DN_TRANSFORM					(WM_USER + 4006)
#define XRT3DN_PROPERTIES                   (WM_USER + 4007)

typedef enum {
	XRT3D_ACTION_NONE,
	XRT3D_ACTION_MODIFY_START,
	XRT3D_ACTION_MODIFY_END,
	XRT3D_ACTION_MODIFY_CANCEL,
	XRT3D_ACTION_ROTATE,
	XRT3D_ACTION_SCALE,
	XRT3D_ACTION_TRANSLATE,
	XRT3D_ACTION_ZOOM_START,
	XRT3D_ACTION_ZOOM_UPDATE,
	XRT3D_ACTION_ZOOM_END,
	XRT3D_ACTION_ZOOM_CANCEL,
	XRT3D_ACTION_RESET,
	XRT3D_ACTION_PROPERTIES,
	/* Axis selections to constrain rotation */
	XRT3D_ACTION_ROTATE_XAXIS = 100,
	XRT3D_ACTION_ROTATE_YAXIS,
	XRT3D_ACTION_ROTATE_ZAXIS,
	XRT3D_ACTION_ROTATE_EYE,
	XRT3D_ACTION_ROTATE_FREE,
} Xrt3dAction;

typedef struct tag_Xrt3dActionItem {
	UINT						 msg;
	UINT						 modifier;
	UINT						 keycode;
	Xrt3dAction					 action;
	struct tag_Xrt3dActionItem	*next;
} Xrt3dActionItem;

#ifdef NDEBUG
#define XRT3D_CHECK()
#else
#define XRT3D_CHECK() xrt3d_check(_T(__FILE__), __LINE__)
#endif

/**
 * Returns the type of data contained in the structure. The structure will be
 * Xrt3dGridData, Xrt3dIrGridData, or Xrt3dPointData. This corresponds to macro
 * values of XRT3D_DATA_GRID, XRT3D_DATA_IRGRID or XRT3D_DATA_POINT, respectively.
 */
#define Xrt3dGetDataType(xrt3d_data)	(xrt3d_data->g.type)

typedef enum {
	XRT3D_ADJUST_LEFT = 1,
	XRT3D_ADJUST_RIGHT,
	XRT3D_ADJUST_CENTER
} Xrt3dAdjust;

typedef enum {
	XRT3D_ALIGN_VERTICAL = 1,
	XRT3D_ALIGN_HORIZONTAL
} Xrt3dAlign;

typedef enum {
	XRT3D_ANCHOR_NORTH		= 0x10,
	XRT3D_ANCHOR_SOUTH		= 0x20,
	XRT3D_ANCHOR_EAST		= 0x01,
	XRT3D_ANCHOR_WEST		= 0x02,
	XRT3D_ANCHOR_NORTHEAST	= 0x11,
	XRT3D_ANCHOR_NORTHWEST	= 0x12,
	XRT3D_ANCHOR_SOUTHEAST	= 0x21,
	XRT3D_ANCHOR_SOUTHWEST	= 0x22
} Xrt3dAnchor;

typedef enum {
	XRT3D_ANNO_VALUES = 1,
	XRT3D_ANNO_DATA_LABELS,
	XRT3D_ANNO_VALUE_LABELS
} Xrt3dAnnoMethod;

typedef enum {
	XRT3D_AXIS_X = 0,
	XRT3D_AXIS_Y,
	XRT3D_AXIS_Z,
	XRT3D_AXIS_NONE,
	XRT3D_AXIS_EYE
} Xrt3dAxis;

typedef enum {
	XRT3D_BAR_FIXED,
	XRT3D_BAR_HISTOGRAM
} Xrt3dBarFormat;

typedef enum {
	XRT3D_BORDER_NONE = 0,
	XRT3D_BORDER_3D_OUT,
	XRT3D_BORDER_3D_IN,
	XRT3D_BORDER_SHADOW,
	XRT3D_BORDER_PLAIN,
	XRT3D_BORDER_ETCHED_IN,
	XRT3D_BORDER_ETCHED_OUT,
	XRT3D_BORDER_FRAME_IN,
	XRT3D_BORDER_FRAME_OUT,
	XRT3D_BORDER_BEVEL
} Xrt3dBorder;

typedef enum {
	XRT3D_DATA_GRID = 1,
	XRT3D_DATA_IRGRID,
	XRT3D_DATA_POINT
} Xrt3dDataType;

typedef enum {
	XRT3D_DISTN_LINEAR,
	XRT3D_DISTN_FROM_TABLE
} Xrt3dDistnMethod;

typedef enum {
    XRT3D_DISTN_RANGE_DATA = 1,
    XRT3D_DISTN_RANGE_ALL
} Xrt3dDistnRange;


typedef enum {
    XRT3D_DRAW_BITMAP = 1,
    XRT3D_DRAW_METAFILE,
    XRT3D_DRAW_ENHMETAFILE,
    XRT3D_DRAW_STANDARDMETAFILE
} Xrt3dDrawFormat;

typedef enum {
    XRT3D_DRAWSCALE_NONE = 1,
    XRT3D_DRAWSCALE_TOWIDTH,
    XRT3D_DRAWSCALE_TOHEIGHT,
    XRT3D_DRAWSCALE_TOFIT,
    XRT3D_DRAWSCALE_TOMAX
} Xrt3dDrawScale;

typedef enum {
    XRT3D_DRAWTYPE_DC = 1,
    XRT3D_DRAWTYPE_PRINTER,
    XRT3D_DRAWTYPE_CLIPBOARD,
    XRT3D_DRAWTYPE_FILE,
} Xrt3dDrawType;


typedef enum {
	XRT3D_INTERP_LINEAR_SPLINE = 1,
	XRT3D_INTERP_CUBIC_SPLINE
} Xrt3dInterpMethod;

typedef enum {
    XRT3D_IMAGELAYOUT_CENTERED = 1,
#ifdef WIN_TOOLKIT
    XRT3D_IMAGELAYOUT_TILED,
    XRT3D_IMAGELAYOUT_FITTED,
    XRT3D_IMAGELAYOUT_STRETCHED,
    XRT3D_IMAGELAYOUT_STRETCHED_TO_WIDTH,
    XRT3D_IMAGELAYOUT_STRETCHED_TO_HEIGHT,
    XRT3D_IMAGELAYOUT_CROP_FITTED
#endif
} Xrt3dImageLayout;

typedef enum	{
	XRT3D_LEGEND_STYLE_STEPPED = 1,
	XRT3D_LEGEND_STYLE_CONTINUOUS
} Xrt3dLegendStyle;

/**
 * List of the support line patterns.
 */
typedef enum    {
    XRT3D_LPAT_NONE = 1,
    XRT3D_LPAT_SOLID,
    XRT3D_LPAT_LONG_DASH,
    XRT3D_LPAT_DOTTED,
    XRT3D_LPAT_SHORT_DASH,
    XRT3D_LPAT_LSL_DASH,
    XRT3D_LPAT_DASH_DOT
} Xrt3dLinePattern;

typedef enum {
	XRT3D_PREVIEW_CUBE = 0,
	XRT3D_PREVIEW_FULL = 100	/**< always last - leave room for alternatives */
} Xrt3dPreviewMethod;

/* stay out of range of Motif reasons */
typedef enum	{
	XRT3D_REASON_MODIFY_START = 500,
	XRT3D_REASON_MODIFY_END,
	XRT3D_REASON_ROTATE,
	XRT3D_REASON_TRANSFORM,
	XRT3D_REASON_MAP,
	XRT3D_REASON_PICK
} Xrt3dReason;

typedef enum {
	XRT3D_RGN_IN_GRAPH = -100,
	XRT3D_RGN_IN_LEGEND,
	XRT3D_RGN_IN_HEADER,
	XRT3D_RGN_IN_FOOTER,
	XRT3D_RGN_NOWHERE
} Xrt3dRegion;

typedef enum {
	XRT3D_SF_CYRILLIC_COMPLEX = 0,
	XRT3D_SF_GOTHIC_ENGLISH,
	XRT3D_SF_GOTHIC_GERMAN,
	XRT3D_SF_GOTHIC_ITALIAN,
	XRT3D_SF_GREEK_COMPLEX,
	XRT3D_SF_GREEK_COMPLEX_SMALL,
	XRT3D_SF_GREEK_SIMPLEX,
	XRT3D_SF_ITALIC_COMPLEX,
	XRT3D_SF_ITALIC_COMPLEX_SMALL,
	XRT3D_SF_ITALIC_TRIPLEX,
	XRT3D_SF_ROMAN_COMPLEX,
	XRT3D_SF_ROMAN_COMPLEX_SMALL,
	XRT3D_SF_ROMAN_DUPLEX,
	XRT3D_SF_ROMAN_SIMPLEX,
	XRT3D_SF_ROMAN_TRIPLEX,
	XRT3D_SF_SCRIPT_COMPLEX,
	XRT3D_SF_SCRIPT_SIMPLEX
} Xrt3dStrokeFont;

typedef enum {
    XRT3D_SYMBOL_NONE = 1,
    XRT3D_SYMBOL_DOT,
    XRT3D_SYMBOL_BOX,
    XRT3D_SYMBOL_TRI,
    XRT3D_SYMBOL_DIAMOND,
    XRT3D_SYMBOL_STAR,
    XRT3D_SYMBOL_VERT_LINE,
    XRT3D_SYMBOL_HORIZ_LINE,
    XRT3D_SYMBOL_CROSS,
    XRT3D_SYMBOL_CIRCLE,
    XRT3D_SYMBOL_SQUARE,
    XRT3D_SYMBOL_INVERT_TRI,
    XRT3D_SYMBOL_DIAG_CROSS,
    XRT3D_SYMBOL_OPEN_TRI,
    XRT3D_SYMBOL_OPEN_DIAMOND,
    XRT3D_SYMBOL_OPEN_INVERT_TRI
} Xrt3dSymbolPattern;

typedef enum {
	XRT3D_ATTACH_INDEX = 0,
	XRT3D_ATTACH_PIXEL,
	XRT3D_ATTACH_POINT
} Xrt3dTextAttachMethod;

typedef enum {
	XRT3D_TYPE_SURFACE,
	XRT3D_TYPE_BAR,
	XRT3D_TYPE_SCATTER
} Xrt3dType;

typedef enum {
	XRT3D_ZONE_CONTOURS,
	XRT3D_ZONE_CELLS
} Xrt3dZoneMethod;


/**
 * Definition for the regualrly gridded data structure.
 * \sa Xrt3dIrGridData and  Xrt3dPointData
 */
typedef struct {
	Xrt3dDataType		type;
	int					numx, numy;
	double				noval;
	double				xstep, ystep;
	double				xorig, yorig;
	double				**values;
} Xrt3dGridData;

/**
 * Definition for irregularly gridded data.
 * \sa Xrt3dGridData and Xrt3dPointData
 */
typedef struct {
	Xrt3dDataType		type;
	int					numx, numy;
	double				noval;
	double				*xgrid, *ygrid;
	double				**values;	
} Xrt3dIrGridData;

/**
 * A single point in an Xrt3dPointSeries structure.
 */
typedef struct {
	double				x, y, z;
} Xrt3dPoint3D;

/**
 * A series of point in an Xrt3dPointData structure.
 */
typedef struct {
	int					npoints;
	Xrt3dPoint3D		*points;
} Xrt3dPointSeries;

/**
 * Definition for point (scatter) data.
 * \sa Xrt3dGridData and Xrt3dIrGridData.
 */
typedef struct {
	Xrt3dDataType		type;
	int					nseries;
	double				noval;
	Xrt3dPointSeries	*series;
} Xrt3dPointData;

/**
 * Opaque handle to the data used by XRT/3d. Use the macro Xrt3dGetDataType()
 * to determine the type of data stored in this union. Possible values are:
 * \arg Xrt3dGridData
 * \arg Xrt3dIrGridData
 * \arg Xrt3dPointData
 */
typedef union {
	Xrt3dGridData		g;
	Xrt3dIrGridData		ig;
	Xrt3dPointData		p;
} Xrt3dData;

typedef struct {
	int					xindex;
	int					yindex;
	XrtColor			color;
} Xrt3dXYColor;

/**
 * Structure used to describe a value label.
 */
typedef struct {
	double				value;		/**< The calue to draw the label at. */
	TCHAR				*label;		/**< The test of the label. */
} Xrt3dValueLabel;

/**
 * Structure defining the style used to draw a contour.
 */
typedef struct {
	XrtColor			fill_color;		/**< The fill color for this contour level */
	XrtColor			line_color;		/**< The line color for this contour line */
	int				 	line_width;		/**< Contour line width */
	Xrt3dLinePattern	lpat;			/**< Line pattern, 2D contours only */

} Xrt3dContourStyle;

typedef struct {
    Xrt3dLinePattern    pattern;
    XrtColor            color;
    int                 width;
	
} Xrt3dLineStyle;

typedef struct {
    Xrt3dSymbolPattern  pattern;
    XrtColor            color;
    int                 size;
	
} Xrt3dSymbolStyle;

typedef struct {
	Xrt3dLineStyle		line_style;
	Xrt3dSymbolStyle	symbol_style;
} Xrt3dDataStyle;

typedef struct {
	int					nentries;
	double				*entry;
} Xrt3dDistnTable;

typedef struct	{
	XrtPosition			pix_x, pix_y;
	double				x, y, z;
} Xrt3dMapResult;

typedef struct	{
	XrtPosition			pix_x, pix_y;
	int					xindex, yindex;
	int					distance;
} Xrt3dPickResult;

typedef struct	{
    HDC         hdc;                /* read only */
    RECT        rectDamaged;        /* read only */
} Xrt3dCallbackStruct;

typedef struct	{
	XrtDimension	width;          /* read only */
	XrtDimension	height;         /* read only */
} Xrt3dResizeCallbackStruct;

typedef struct	{
	XrtPosition     x;              /* read only */
	XrtPosition     y;              /* read only */
} Xrt3dPropertiesCallbackStruct;

typedef struct	{
	
	Xrt3dRegion		region;			/**< read only */
	Xrt3dMapResult	map;			/**< read only */
} Xrt3dMapCallbackStruct;

typedef struct	{

	XrtBoolean	doit;
} Xrt3dModifyCallbackStruct;

typedef struct	{

	double		xrotation;
	double		yrotation;
	double		zrotation;
	XrtBoolean	doit;
} Xrt3dRotateCallbackStruct;

typedef struct	{

	double		scale;
	double		xtranslate;
	double		ytranslate;
	XrtBoolean	doit;
} Xrt3dTransformCallbackStruct;

typedef struct	{
	
	Xrt3dRegion		region;			/**< read only */
	Xrt3dPickResult	pick;			/**< read only */
} Xrt3dPickCallbackStruct;

typedef struct {
	short v[4];
	TCHAR *ProductName;
	TCHAR *ProductVersion;
	TCHAR *Copyright;
	TCHAR *Copyright2;
} Xrt3dVersionInfo;

typedef enum {
	XRT3D_FONT_ROTATION_NONE = 0,
	XRT3D_FONT_ROTATION_2D,
	XRT3D_FONT_ROTATION_3D
} Xrt3dFontRotation;

typedef enum {
	XRT3D_ANNO_POSITION_BOTH = 0,
	XRT3D_ANNO_POSITION_NEAR,
	XRT3D_ANNO_POSITION_FAR,
	XRT3D_ANNO_POSITION_NONE
} Xrt3dAnnoPosition;

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

#endif /* ~xrt3d_common_DEFINED */
