#if !defined( OC_ERROR_CONSTANTS )
#define OC_ERROR_CONSTANTS

enum /* ErrorConstants */               /* Specifies the error detected by the chart. */
{
    ocErrorInvalidIndex = 0,            /* Invalid index found in variant */
    ocErrorReadOnly = 1,                /* Property is read-only and cannot be modified */
    ocErrorOutOfRange = 2,              /* Numeric value is out of range */
    ocErrorInvalidEnum = 3,             /* Invalid enumerated value for property */
    ocErrorInvalidValue = 4,            /* Invalid value for property */
    ocErrorNoSuchProperty = 5,          /* Property does not exist for this instance of the object */
    ocErrorDuplicateValue = 6,          /* Collection already contains a duplicate value */
    ocErrorNoSuchItem = 7,              /* Cannot find item in collection */
    ocErrorCannotCreateObject = 8,      /* Cannot create object */
    ocErrorTypeMismatch = 9,            /* Property set to wrong object type */
    ocErrorReadOnlyWhenBound = 10,      /* Property is read-only when the chart group is bound */
    ocErrorBadArrayDimensions = 11,     /* Wrong dimensions in SafeArray passed to a function */
    ocErrorNotConnected = 12,           /* Object does not have a parent */
    ocErrorInvalidFormat = 13,          /* Invalid custom format */
    ocErrorDllNotCompatible = 14,       /* DLL version is not compatible */
    ocErrorSeriesArrayNotOne = 15,      /* Series index must be 1 for Array data layout */
};

#endif

enum /* AdjustConstants */              /* Specifies the alignment of multiple lines of text. */
{
    oc2dAdjustLeft = 1,                 /* Aligned to left edge. */
    oc2dAdjustRight = 2,                /* Aligned to right edge. */
    oc2dAdjustCenter = 3,               /* Centered. */
};

enum /* OrientationConstants */         /* Specifies the orientation. */
{
    oc2dOrientVertical = 1,             /* Vertical orientation. */
    oc2dOrientHorizontal = 2,           /* Horizontal orientation. */
};

enum /* AnchorConstants */              /* Specifies the position of the object relative to the anchor position. */
{
    oc2dAnchorNorth = 16,               /* Above anchor position. */
    oc2dAnchorSouth = 32,               /* Below anchor position. */
    oc2dAnchorEast = 1,                 /* To right of anchor position. */
    oc2dAnchorWest = 2,                 /* To left of anchor position. */
    oc2dAnchorNorthEast = 17,           /* Above and to right of anchor position. */
    oc2dAnchorNorthWest = 18,           /* Above and to left of anchor position. */
    oc2dAnchorSouthEast = 33,           /* Below and to right of position. */
    oc2dAnchorSouthWest = 34,           /* Below and to left of position. */
    oc2dAnchorCenter = 0,               /* Center over anchor position (ChartLabel object only). */
    oc2dAnchorAuto = 256,               /* Automatically position the object (ChartLabel object only). */
};

enum /* AnnotationMethodConstants */    /* Specifies the annotation method. */
{
    oc2dAnnotateValues = 0,             /* Annotate using chart's X- or Y-values. */
    oc2dAnnotatePointLabels = 1,        /* Annotate using Point-labels (X-axis only). */
    oc2dAnnotateValueLabels = 2,        /* Annotate using ValueLabels property. */
    oc2dAnnotateTimeLabels = 3,         /* Annotate using Time-labels (X-axis only). */
};

enum /* AnnotationPlacementConstants */  /* Specifies the placement of the axis annotation and title. */
{
    oc2dAnnotateAuto = 0,               /* Annotation placement automatically determined. */
    oc2dAnnotateOrigin = 1,             /* Annotation along the origin. */
    oc2dAnnotateMinimum = 2,            /* Annotation along the axis minimum. */
    oc2dAnnotateMaximum = 3,            /* Annotation along the axis maximum. */
};

enum /* AttachMethodConstants */        /* Specifies how a chart label is attached to a chart. */
{
    oc2dAttachCoord = 0,                /* Attach at specified pixel coordinate. */
    oc2dAttachDataCoord = 1,            /* Attach to specified data coordinates. */
    oc2dAttachDataIndex = 2,            /* Attach to specified data indices. */
    oc2dAttachDataIndexDataY = 3,       /* Attach using data index for X coordinate, and data value for Y coordinate. */
};

enum /* AxisIDConstants */              /* Specifies a chart axis. */
{
    oc2dAxisX = 1,                      /* X-axis. */
    oc2dAxisY = 2,                      /* Y-axis. */
    oc2dAxisY2 = 3,                     /* Y2-axis (second Y-axis). */
};

enum /* BorderTypeConstants */          /* Valid border styles. */
{
    oc2dBorderNone = 0,                 /* Border None */
    oc2dBorder3DOut = 1,                /* Border 3D Out */
    oc2dBorder3DIn = 2,                 /* Border 3D In */
    oc2dBorderShadow = 3,               /* Border Shadow */
    oc2dBorderPlain = 4,                /* Border Plain */
    oc2dBorderEtchedIn = 5,             /* Border Etched In */
    oc2dBorderEtchedOut = 6,            /* Border Etched Out */
    oc2dBorderFrameIn = 7,              /* Border Frame In */
    oc2dBorderFrameOut = 8,             /* Border Frame Out */
    oc2dBorderBevel = 9,                /* Border Bevel */
};

enum /* DataLayoutConstants */          /* Specifies the layout of the data to be graphed. */
{
    oc2dDataArray = 1,                  /* Array data. */
    oc2dDataGeneral = 2,                /* General data. */
};

enum /* DisplayConstants */             /* Specifies the visibility of the item. */
{
    oc2dDisplayShow = 1,                /* Show */
    oc2dDisplayHide = 2,                /* Hide */
    oc2dDisplayExclude = 3,             /* Exclude */
};

enum /* FillPatternConstants */         /* Valid fill patterns. */
{
    oc2dFillNone = 1,                   /* None */
    oc2dFillSolid = 2,                  /* Solid */
    oc2dFill25Percent = 3,              /* 25 % */
    oc2dFill50Percent = 4,              /* 50 % */
    oc2dFill75Percent = 5,              /* 75 % */
    oc2dFillHorizontalStripe = 6,       /* Horizontal Stripe */
    oc2dFillVerticalStripe = 7,         /* Vertical Stripe */
    oc2dFill45DegreeStripe = 8,         /* 45 Degree Stripe */
    oc2dFill135DegreeStripe = 9,        /* 135 Degree Stripe */
    oc2dFillDiagHatched = 10,           /* Diagonal Hatched */
    oc2dFillCrossHatched = 11,          /* Cross Hatched */
    oc2dFillBackwardDiagonal = 12,      /* Backward Diagonal */
    oc2dFillCross = 13,                 /* Cross */
    oc2dFillDiagonalCross = 14,         /* Diagonal Cross */
    oc2dFillForwardDiagonal = 15,       /* Forward Diagonal */
    oc2dFillHorizontal = 16,            /* Horizontal */
    oc2dFillVertical = 17,              /* Vertical */
};

enum /* FocusConstants */               /* Specifies how the nearest series and point value are determined. */
{
    oc2dFocusX = 1,                     /* Use X-axis to determine closest value. */
    oc2dFocusY = 2,                     /* Use Y-axis to determine closest value. */
    oc2dFocusXY = 3,                    /* Use both axes to determine closest value. */
};

enum /* LinePatternConstants */         /* Valid line patterns. */
{
    oc2dLineNone = 1,                   /* None */
    oc2dLineSolid = 2,                  /* Solid */
    oc2dLineLongDash = 3,               /* Long Dash */
    oc2dLineDotted = 4,                 /* Dotted */
    oc2dLineShortDash = 5,              /* Short Dash */
    oc2dLineLongShortLongDash = 6,      /* Long Short Long Dash */
    oc2dLineDashDot = 7,                /* Dash Dot */
};

enum /* MarkerMethodConstants */        /* Specifies the marker positioning method used. */
{
    oc2dMarkerDataCoord = 2,            /* Specify position in data coordinates. */
    oc2dMarkerDataIndex = 1,            /* Specify position as an index into a data set. */
};

enum /* NumMethodConstants */           /* Specifies the method used to calculate axis numbering. */
{
    oc2dNumPrecision = 0,               /* Use current value of axis precision. */
    oc2dNumRound = 1,                   /* Use rounded numbers (most significant digit is 1, 2 or 5). */
};

enum /* OriginPlacementConstants */     /* Specifies the location of the origin. */
{
    oc2dOriginAuto = 0,                 /* Origin location automatically determined. */
    oc2dOriginZero = 1,                 /* Origin at zero. */
    oc2dOriginMinimum = 2,              /* Origin at axis minimum. */
    oc2dOriginMaximum = 3,              /* Origin at axis maximum. */
};

enum /* SortOrderConstants */           /* Specifies the order in which pie slices should be displayed. */
{
    oc2dSortOrderAscending = 1,         /* Ascending order. */
    oc2dSortOrderDescending = 2,        /* Descending order. */
    oc2dSortOrderNone = 3,              /* Data Order. */
};

enum /* ThresholdConstants */           /* Specifies the method used to determine which point values will be grouped into the "other" slice in pie charts. */
{
    oc2dThresholdSliceCutoff = 1,       /* Group all values whose individual percentage is less than ThresholdValue. */
    oc2dThresholdPercentile = 2,        /* Group values whose total percentage is less than ThresholdValue. */
};

enum /* ShapeConstants */               /* Valid symbol shapes. */
{
    oc2dShapeNone = 1,                  /* None */
    oc2dShapeDot = 2,                   /* Dot */
    oc2dShapeBox = 3,                   /* Box */
    oc2dShapeTriangle = 4,              /* Triangle */
    oc2dShapeDiamond = 5,               /* Diamond */
    oc2dShapeStar = 6,                  /* Star */
    oc2dShapeVerticalLine = 7,          /* Vertical Line */
    oc2dShapeHorizontalLine = 8,        /* Horizontal Line */
    oc2dShapeCross = 9,                 /* Cross */
    oc2dShapeCircle = 10,               /* Circle */
    oc2dShapeSquare = 11,               /* Square */
    oc2dShapeInvertTriangle = 12,       /* Invert Triangle */
    oc2dShapeDiagonalCross = 13,        /* Diagonal Cross */
    oc2dShapeOpenTriangle = 14,         /* Open Triangle */
    oc2dShapeOpenDiamond = 15,          /* Open Diamond */
    oc2dShapeOpenInvertTriangle = 16,   /* Open Invert Triangle */
    oc2dShapeImageFile = 17,			/* Image File */
};

enum /* RegionConstants */              /* Specifies a region of a chart. */
{
    oc2dRegionInChartArea = -100,       /* Located in chart area. */
    oc2dRegionInLegend = -99,           /* Located in legend. */
    oc2dRegionInHeader = -98,           /* Located in header. */
    oc2dRegionInFooter = -97,           /* Located in footer. */
    oc2dRegionNone = -96,               /* No particular region. */
};

enum /* RotationConstants */            /* Specifies an angle of rotation (in counterclockwise degrees). */
{
    oc2dRotateNone = 1,                 /* No rotation. */
    oc2dRotate90Degrees = 2,            /* Rotate 90 degrees. */
    oc2dRotate270Degrees = 3,           /* Rotate 270 degrees. */
    oc2dRotate45Degrees = 4,            /* Rotate 45 degrees. */
    oc2dRotate315Degrees = 5,           /* Rotate 315 degrees. */
};

enum /* TimeUnitConstants */            /* Specifies the time unit to use in a time-axis. */
{
    oc2dTimeSeconds = 1,                /* Seconds */
    oc2dTimeMinutes = 2,                /* Minutes */
    oc2dTimeHours = 3,                  /* Hours */
    oc2dTimeDays = 4,                   /* Days */
    oc2dTimeWeeks = 5,                  /* Weeks */
    oc2dTimeMonths = 6,                 /* Months */
    oc2dTimeYears = 7,                  /* Years */
};

enum /* ChartTypeConstants */           /* Specifies the chart display type. */
{
    oc2dTypePlot = 1,                   /* Line chart. */
    oc2dTypeBar = 2,                    /* Bar chart. */
    oc2dTypePie = 3,                    /* Pie chart. */
    oc2dTypeStackingBar = 4,            /* Stacking bar chart. */
    oc2dTypeArea = 5,                   /* Area chart. */
    oc2dTypeHiLo = 6,                   /* HiLo chart. */
    oc2dTypeHiLoOpenClose = 7,          /* HiLoOpenClose chart. */
    oc2dTypeCandle = 8,                 /* Candle chart. */
    oc2dTypePolar = 9,                  /* Polar chart. */
    oc2dTypeRadar = 10,                 /* Radar chart. */
    oc2dTypeFilledRadar = 11,           /* Filled radar chart. */
    oc2dTypeBubble = 12,                /* Bubble chart. */
};

enum /* ActionConstants */              /* Specifies a user action. */
{
    oc2dActionNone = 0,                 /* No action. */
    oc2dActionModifyStart = 1,          /* Start a user interaction. */
    oc2dActionModifyEnd = 2,            /* Notify application that user interaction has ended. */
    oc2dActionRotate = 3,               /* Rotate the chart. */
    oc2dActionScale = 4,                /* Scale the chart. */
    oc2dActionTranslate = 5,            /* Translate the chart. */
    oc2dActionZoomStart = 6,            /* Start zoom. */
    oc2dActionZoomUpdate = 7,           /* Update the zoom area selection rubberband. */
    oc2dActionZoomEnd = 8,              /* Zoom into chart at the area selected by the rubberband. */
    oc2dActionZoomCancel = 9,           /* Cancel zoom in progress. */
    oc2dActionReset = 10,               /* Restore all scaling and translations to default. */
    oc2dActionProperties = 11,          /* Display property page. */
    oc2dActionZoomAxisEnd = 12,         /* Zoom into chart at selected area, retaining the axes. */
};

enum /* FormatConstants */              /* Specifies the format to use when printing. */
{
    oc2dFormatBitmap = 1,               /* Bitmap (.BMP). */
    oc2dFormatMetafile = 2,             /* Metafile (.WMF). */
    oc2dFormatEnhMetafile = 3,          /* Enhanced metafile (.EMF). */
    oc2dFormatStandardMetafile = 4,     /* Standard metafile (.WMF). */
};

enum /* ScaleConstants */               /* Specifies the scaling to use when printing. */
{
    oc2dScaleNone = 1,                  /* Same size as actual chart. */
    oc2dScaleToWidth = 2,               /* Scale to specified width, preserving aspect ratio. */
    oc2dScaleToHeight = 3,              /* Scale to specified height, preserving aspect ratio. */
    oc2dScaleToFit = 4,                 /* Scale to existing window (minimum of height and width), preserving aspect ratio. */
    oc2dScaleToMax = 5,                 /* Enlarge to size of page. */
};

enum /* AngleUnitConstants */           /* Specifies the angle unit of polar and radar charts. */
{
    oc2dAngleUnitDegrees = 1,           /* Degree angles. */
    oc2dAngleUnitRadians = 2,           /* Radian angles. */
    oc2dAngleUnitGrads = 3,             /* Grad angles. */
};

enum /* ShadingConstants */             /* Specifies the shading method applied to 3D areas. */
{
    oc2dShadingColor = 1,               /* Colored 3D effects. */
    oc2dShadingDithered = 2,            /* Dithered 3D effects. */
    oc2dShadingNone = 3,                /* No special 3D effects. */
};

enum /* CategoryConstants */            /* Specifies the value format category. */
{
    oc2dCategoryStandard = 1,           /* . */
    oc2dCategoryNumber = 2,             /* 1,234.56. */
    oc2dCategoryCurrency = 3,           /* $1,234.56. */
    oc2dCategoryAccounting = 4,         /* $1,234.56. */
    oc2dCategoryPercentage = 7,         /* 12.34%. */
    oc2dCategoryFraction = 8,           /* 1,234 56/100. */
    oc2dCategoryScientific = 9,         /* 1.23E+03. */
    oc2dCategoryCustom = 12,            /* . */
};

enum /* OrderSignValueConstants */      /* Specifies the order of the sign and value. */
{
    oc2dosvValueSign = 1,               /* Value, sign:  1.1-. */
    oc2dosvSignValue = 2,               /* Sign, value:  -1.1. */
};

enum /* OrderCurrencySignValueConstants */  /* Specifies the order of the currency symbol, sign and value. */
{
    oc2docsvValueSignCurrency = 1,      /* Value, sign, currency:  1.1-$. */
    oc2docsvSignValueCurrency = 2,      /* Sign, value, currency:  -1.1$. */
    oc2docsvValueCurrencySign = 3,      /* Value, currency, sign:  1.1$-. */
    oc2docsvSignCurrencyValue = 4,      /* Sign, currency, value:  -$1.1. */
    oc2docsvCurrencyValueSign = 5,      /* Currency, value, sign:  $1.1-. */
    oc2docsvCurrencySignValue = 6,      /* Currency, sign, value:  $-1.1. */
};

enum /* OrderPercentSignValueConstants */  /* Specifies the order of the percent symbol, sign and value. */
{
    oc2dopsvValueSignPercent = 1,       /* Value, sign, percent:  1.1-%. */
    oc2dopsvSignValuePercent = 2,       /* Sign, value, percent:  -1.1%. */
    oc2dopsvValuePercentSign = 3,       /* Value, percent, sign:  1.1%-. */
    oc2dopsvSignPercentValue = 4,       /* Sign, percent, value:  -%1.1. */
    oc2dopsvPercentValueSign = 5,       /* Percent, value, sign:  %1.1-. */
    oc2dopsvPercentSignValue = 6,       /* Percent, sign, value:  %-1.1. */
};

enum /* ChartEditorElementConstants */  /* Chart editor elements. */
{
    oc2dceeControl = 1,                 /* . */
    oc2dceeAxes = 2,                    /* . */
    oc2dceeChartGroups = 3,             /* . */
    oc2dceeChartStyles = 4,             /* . */
    oc2dceeTitles = 5,                  /* . */
    oc2dceeView3D = 9,                  /* . */
    oc2dceeMarkers = 10,                /* . */
    oc2dceeLegend = 11,                 /* . */
    oc2dceeChartArea = 12,              /* . */
    oc2dceePlotArea = 13,               /* . */
    oc2dceeChartLabels = 14,            /* . */
    oc2dceeControlGeneral = 41,         /* . */
    oc2dceeControlBorder = 42,          /* . */
    oc2dceeControlInterior = 43,        /* . */
    oc2dceeControlImage = 44,           /* . */
    oc2dceeControlAbout = 45,           /* . */
    oc2dceeAxesGeneral = 61,            /* . */
    oc2dceeAxesAnnotation = 62,         /* . */
    oc2dceeAxesScale = 63,              /* . */
    oc2dceeAxesTitle = 64,              /* . */
    oc2dceeAxesGrid = 65,               /* . */
    oc2dceeAxesGridStyle = 66,          /* . */
    oc2dceeAxesFont = 67,               /* . */
    oc2dceeAxesPolarRadar = 68,         /* . */
    oc2dceeAxesAxisStyle = 69,          /* . */
    oc2dceeChartGroupsGeneral = 81,     /* . */
    oc2dceeChartGroupsData = 82,        /* . */
    oc2dceeChartGroupsLabels = 83,      /* . */
    oc2dceeChartGroupsBinding = 84,     /* . */
    oc2dceeChartGroupsBar = 85,         /* . */
    oc2dceeChartGroupsCandle = 86,      /* . */
    oc2dceeChartGroupsHiLoOpenClose = 87,  /* . */
    oc2dceeChartGroupsPie = 88,         /* . */
    oc2dceeChartGroupsOtherSlice = 89,  /* . */
    oc2dceeChartGroupsPolar = 90,       /* . */
    oc2dceeChartGroupsInternet = 91,    /* . */
    oc2dceeChartGroupsBubble = 92,      /* . */
    oc2dceeChartStylesFillStyle = 101,  /* . */
    oc2dceeChartStylesLineStyle = 102,  /* . */
    oc2dceeChartStylesSymbolStyle = 103,  /* . */
    oc2dceeChartStylesSliceStyle = 104,  /* . */
    oc2dceeChartStylesImage = 105,  /* . */
    oc2dceeTitlesGeneral = 121,         /* . */
    oc2dceeTitlesLabel = 122,           /* . */
    oc2dceeTitlesLocation = 123,        /* . */
    oc2dceeTitlesBorder = 124,          /* . */
    oc2dceeTitlesInterior = 125,        /* . */
    oc2dceeTitlesFont = 126,            /* . */
    oc2dceeTitlesImage = 127,           /* . */
    oc2dceeView3DGeneral = 201,         /* . */
    oc2dceeMarkersGeneral = 221,        /* . */
    oc2dceeMarkersAttach = 222,         /* . */
    oc2dceeMarkersLineStyle = 223,      /* . */
    oc2dceeLegendGeneral = 241,         /* . */
    oc2dceeLegendLocation = 242,        /* . */
    oc2dceeLegendBorder = 243,          /* . */
    oc2dceeLegendInterior = 244,        /* . */
    oc2dceeLegendFont = 245,            /* . */
    oc2dceeLegendImage = 246,           /* . */
    oc2dceeLegendTitle = 247,           /* . */
    oc2dceeChartAreaGeneral = 261,      /* . */
    oc2dceeChartAreaLocation = 262,     /* . */
    oc2dceeChartAreaBorder = 263,       /* . */
    oc2dceeChartAreaInterior = 264,     /* . */
    oc2dceeChartAreaImage = 265,        /* . */
    oc2dceePlotAreaGeneral = 281,       /* . */
    oc2dceePlotAreaInterior = 282,      /* . */
    oc2dceePlotAreaImage = 283,         /* . */
    oc2dceeChartLabelsGeneral = 301,    /* . */
    oc2dceeChartLabelsLabel = 302,      /* . */
    oc2dceeChartLabelsBorder = 303,     /* . */
    oc2dceeChartLabelsInterior = 304,   /* . */
    oc2dceeChartLabelsFont = 305,       /* . */
    oc2dceeChartLabelsLocation = 306,   /* . */
    oc2dceeChartLabelsImage = 307,      /* . */
    oc2dceeControlGeneralIsBatched = 359,    /* . */
    oc2dceeControlGeneralIsDoubleBuffered = 360,    /* . */
    oc2dceeControlGeneralLoad = 361,    /* . */
    oc2dceeControlGeneralSave = 362,    /* . */
    oc2dceeControlGeneralReset = 363,    /* . */
    oc2dceeChartStylesAdd = 401,        /* . */
    oc2dceeChartStylesRemove = 402,     /* . */
    oc2dceeChartGroupsDataLoad = 441,   /* . */
    oc2dceeChartGroupsDataSave = 442,   /* . */
    oc2dceeChartGroupsDataSort = 443,   /* . */
    oc2dceeChartGroupsDataEdit = 444,   /* . */
    oc2dceeChartGroupsDataSeries = 445,  /* . */
};

enum /* ImageLayoutConstants */         /* Specifies the image layout. */
{
    oc2dImageCentered = 1,              /* Centered image */
    oc2dImageTiled = 2,                 /* Tiled image */
    oc2dImageFitted = 3,                /* Fitted image */
    oc2dImageStretched = 4,             /* Stretched image */
    oc2dImageStretchedToWidth = 5,      /* Stretched to width image */
    oc2dImageStretchedToHeight = 6,     /* Stretched to height image */
    oc2dImageCropFitted = 7,            /* Crop fitted image */
};

enum /* BubbleMethodConstants */        /* Specifies the method of determining bubble size */
{
    oc2dBubbleDiameter = 1,             /* Bubble size based on diameter */
    oc2dBubbleArea = 2,                 /* Bubble size based on area */
};

