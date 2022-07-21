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

enum /* AdjustConstants */              /* Specifies the alignment of multiple lines of text in a header or footer. */
{
    oc3dAdjustLeft = 1,                 /* Aligned to left edge. */
    oc3dAdjustRight = 2,                /* Aligned to right edge. */
    oc3dAdjustCenter = 3,               /* Centered. */
};

enum /* OrientationConstants */         /* Specifies the orientation. */
{
    oc3dOrientVertical = 1,             /* Vertical orientation. */
    oc3dOrientHorizontal = 2,           /* Horizontal orientation. */
};

enum /* AnchorConstants */              /* Specifies the position of the object relative to the anchor position. */
{
    oc3dAnchorNorth = 16,               /* Above anchor position. */
    oc3dAnchorSouth = 32,               /* Below anchor position. */
    oc3dAnchorEast = 1,                 /* To right of anchor position. */
    oc3dAnchorWest = 2,                 /* To left of anchor position. */
    oc3dAnchorNorthEast = 17,           /* Above and to right of anchor position. */
    oc3dAnchorNorthWest = 18,           /* Above and to left of anchor position. */
    oc3dAnchorSouthEast = 33,           /* Below and to right of position. */
    oc3dAnchorSouthWest = 34,           /* Below and to left of position. */
};

enum /* AnnotationMethodConstants */    /* Specifies the annotation method. */
{
    oc3dAnnotateValues = 1,             /* Annotate using chart's data values. */
    oc3dAnnotateDataLabels = 2,         /* Annotate using data labels. */
    oc3dAnnotateValueLabels = 3,        /* Annotate using ValueLabels property. */
};

enum /* AttachMethodConstants */        /* Specifies how a chart label is attached to a chart. */
{
    oc3dAttachDataIndex = 0,            /* Attach to a grid index. */
    oc3dAttachCoord = 1,                /* Attach to pixel coordinates on the chart. */
    oc3dAttachDataCoord = 2,            /* Attach to a point in 3D space. */
};

enum /* AxisIDConstants */              /* Specifies a chart axis. */
{
    oc3dAxisX = 1,                      /* X-axis. */
    oc3dAxisY = 2,                      /* Y-axis. */
    oc3dAxisZ = 3,                      /* Z-axis. */
};

enum /* BarFormatConstants */           /* The format in which bars are displayed. */
{
    oc3dBarFixed = 0,                   /* Fixed format. */
    oc3dBarHistogram = 1,               /* Histogram format. */
};

enum /* BorderTypeConstants */          /* Valid border styles. */
{
    oc3dBorderNone = 0,                 /* Border None. */
    oc3dBorder3DOut = 1,                /* Border 3D Out. */
    oc3dBorder3DIn = 2,                 /* Border 3D In. */
    oc3dBorderShadow = 3,               /* Border Shadow. */
    oc3dBorderPlain = 4,                /* Border Plain. */
    oc3dBorderEtchedIn = 5,             /* Border Etched In. */
    oc3dBorderEtchedOut = 6,            /* Border Etched Out. */
    oc3dBorderFrameIn = 7,              /* Border Frame In */
    oc3dBorderFrameOut = 8,             /* Border Frame Out */
    oc3dBorderBevel = 9,                /* Border Bevel */
};

enum /* DataLayoutConstants */          /* Specifies the type of data to be charted. */
{
    oc3dDataGrid = 1,                   /* Regular grid data. */
    oc3dDataIrregularGrid = 2,          /* Irregular grid data. */
    oc3dDataPoint = 3,                  /* Point data. */
};

enum /* InterpolationConstants */       /* Specifies the spline interpolation method to use. */
{
    oc3dInterpolateLinearSpline = 1,    /* Linear spline interpolation. */
    oc3dInterpolateCubicSpline = 2,     /* Cubic spline interpolation. */
};

enum /* DistributionRangeConstants */   /* Specifies the distribution range to use. */
{
    oc3dDistributionRangeData = 1,      /* Range spanned by the data. */
    oc3dDistributionRangeAll = 2,       /* Entire range of distribution table. */
};

enum /* LegendTypeConstants */          /* Determines the legend style. */
{
    oc3dLegendTypeStepped = 1,          /* Range printed for each zone. */
    oc3dLegendTypeContinuous = 2,       /* Level printed at the division between two adjacent zones. */
};

enum /* LinePatternConstants */         /* Valid line patterns. */
{
    oc3dLineNone = 1,                   /* None */
    oc3dLineSolid = 2,                  /* Solid */
    oc3dLineLongDash = 3,               /* Long Dash */
    oc3dLineDotted = 4,                 /* Dotted */
    oc3dLineShortDash = 5,              /* Short Dash */
    oc3dLineLongShortLongDash = 6,      /* Long Short Long Dash */
    oc3dLineDashDot = 7,                /* Dash Dot */
};

enum /* PlaneConstants */               /* Specifies a plane of the chart. */
{
    oc3dXYPlane = 1,                    /* The XY-Plane */
    oc3dXZPlane = 2,                    /* The XZ-Plane */
    oc3dYZPlane = 4,                    /* The YZ-Plane */
};

enum /* RegionConstants */              /* Specifies a region of a chart. */
{
    oc3dRegionInChartArea = -100,       /* Located in chart area. */
    oc3dRegionInLegend = -99,           /* Located in legend. */
    oc3dRegionInHeader = -98,           /* Located in header. */
    oc3dRegionInFooter = -97,           /* Located in footer. */
    oc3dRegionNone = -96,               /* No particular region. */
};

enum /* StrokeFontConstants */          /* Specifies a stroke font. */
{
    oc3dCyrillicComplex = 0,            /* Cyrillic Complex */
    oc3dGothicEnglish = 1,              /* Gothic English */
    oc3dGothicGerman = 2,               /* Gothic German */
    oc3dGothicItalian = 3,              /* Gothic Italian */
    oc3dGreekComplex = 4,               /* Greek Complex */
    oc3dGreekComplexSmall = 5,          /* Greek Complex Small */
    oc3dGreekSimplex = 6,               /* Greek Simplex */
    oc3dItalicComplex = 7,              /* Italic Complex */
    oc3dItalicComplexSmall = 8,         /* Italic Complex Small */
    oc3dItalicTriplex = 9,              /* Italic Triplex */
    oc3dRomanComplex = 10,              /* Roman Complex */
    oc3dRomanComplexSmall = 11,         /* Roman Complex Small */
    oc3dRomanDuplex = 12,               /* Roman Duplex */
    oc3dRomanSimplex = 13,              /* Roman Simplex */
    oc3dRomanTriplex = 14,              /* Roman Triplex */
    oc3dScriptComplex = 15,             /* Script Complex */
    oc3dScriptSimplex = 16,             /* Script Simplex */
};

enum /* ChartTypeConstants */           /* Specifies the chart display type. */
{
    oc3dTypeSurface = 0,                /* Surface chart. */
    oc3dTypeBar = 1,                    /* Bar chart. */
    oc3dTypeScatter = 2,                /* Scatter chart. */
};

enum /* ZoneMethodConstants */          /* Specifies the method used to fill each zone region. */
{
    oc3dZoneContours = 0,               /* Contour zoning. */
    oc3dZoneCells = 1,                  /* Cell zoning. */
};

enum /* PreviewMethodConstants */       /* Specifies how interactive rotations are to be presented to the user. */
{
    oc3dPreviewCube = 0,                /* Draw outline of plot cube. */
    oc3dPreviewFull = 100,              /* Draw entire chart. */
};

enum /* ActionConstants */              /* Specifies a user action. */
{
    oc3dActionNone = 0,                 /* No action. */
    oc3dActionModifyStart = 1,          /* Start any user interaction. */
    oc3dActionModifyEnd = 2,            /* Notify application that user interaction has ended. */
    oc3dActionModifyCancel = 3,         /* Cancel a user interaction. */
    oc3dActionRotate = 4,               /* Rotate the chart. */
    oc3dActionScale = 5,                /* Scale the chart. */
    oc3dActionTranslate = 6,            /* Translate the chart. */
    oc3dActionZoomStart = 7,            /* Start zoom. */
    oc3dActionZoomUpdate = 8,           /* Update the zoom rubberband. */
    oc3dActionZoomEnd = 9,              /* End a zoom. */
    oc3dActionZoomCancel = 10,          /* Cancel a zoom in progress. */
    oc3dActionReset = 11,               /* Undo all scaling and translations. */
    oc3dActionProperties = 12,          /* Display property page. */
    oc3dActionRotateXAxis = 100,        /* Rotate about X-axis. */
    oc3dActionRotateYAxis = 101,        /* Rotate about Y-axis. */
    oc3dActionRotateZAxis = 102,        /* Rotate about Z-axis. */
    oc3dActionRotateEye = 103,          /* Rotate about the eyepoint. */
    oc3dActionRotateFree = 104,         /* Unconstrained Rotation. */
};

enum /* FormatConstants */              /* Specifies the format to use when printing. */
{
    oc3dFormatBitmap = 1,               /* Bitmap. */
    oc3dFormatMetafile = 2,             /* Metafile. */
    oc3dFormatEnhMetafile = 3,          /* Enhanced Metafile. */
    oc3dFormatStandardMetafile = 4,     /* Standard Metafile. */
};

enum /* ScaleConstants */               /* Specifies the scaling to use when printing. */
{
    oc3dScaleNone = 1,                  /* Same size as actual chart. */
    oc3dScaleToWidth = 2,               /* Scale to specified width, preserving aspect ratio. */
    oc3dScaleToHeight = 3,              /* Scale to specified height, preserving aspect ratio. */
    oc3dScaleToFit = 4,                 /* Scale to existing window (minimum of height and width), preserving aspect ratio. */
    oc3dScaleToMax = 5,                 /* Enlarge to size of page. */
};

enum /* ChartEditorElementConstants */  /* Chart editor elements constants. */
{
    oc3dceeControl = 1,                 /* . */
    oc3dceeAxes = 2,                    /* . */
    oc3dceeChartGroup = 3,              /* . */
    oc3dceeStyles = 4,                  /* . */
    oc3dceeTitles = 5,                  /* . */
    oc3dceeLegend = 6,                  /* . */
    oc3dceeChartArea = 7,               /* . */
    oc3dceePlotCube = 8,                /* . */
    oc3dceeChartLabels = 9,             /* . */
    oc3dceeBar = 10,                    /* . */
    oc3dceeSurface = 11,                /* . */
    oc3dceeView3D = 12,                 /* . */
    oc3dceeLevels = 13,                 /* . */
    oc3dceeChartStyles = 14,            /* . */
    oc3dceeControlGeneral = 41,         /* . */
    oc3dceeControlBorder = 42,          /* . */
    oc3dceeControlInterior = 43,        /* . */
    oc3dceeControlAbout = 44,           /* . */
    oc3dceeControlImage = 45,           /* . */
    oc3dceeAxesGeneral = 61,            /* . */
    oc3dceeAxesScale = 62,              /* . */
    oc3dceeAxesTitle = 63,              /* . */
    oc3dceeAxesGridLines = 64,          /* . */
    oc3dceeAxesValueLabels = 65,        /* . */
    oc3dceeAxesStrokeFont = 66,         /* . */
    oc3dceeAxesGridStyle = 67,          /* . */
    oc3dceeAxesFont = 68,               /* . */
    oc3dceeChartGroupGeneral = 81,      /* . */
    oc3dceeChartGroupData = 82,         /* . */
    oc3dceeChartGroupLabels = 83,       /* . */
    oc3dceeChartGroupElevation = 84,    /* . */
    oc3dceeChartGroupInternet = 85,     /* . */
	oc3dceeChartGroupScatter = 86,      /* . */
	oc3dceeChartGroupSurface = 87,      /* . */
	oc3dceeChartGroupBar = 88,          /* . */
	oc3dceeChartGroupBarColor = 89,     /* . */
    oc3dceeStylesGeneral = 101,         /* . */
    oc3dceeStylesFillStyle = 102,       /* . */
    oc3dceeStylesLineStyle = 103,       /* . */
    oc3dceeStylesAdd = 104,             /* . */
    oc3dceeStylesRemove = 105,          /* . */
    oc3dceeTitlesGeneral = 121,         /* . */
    oc3dceeTitlesLabel = 122,           /* . */
    oc3dceeTitlesLocation = 123,        /* . */
    oc3dceeTitlesBorder = 124,          /* . */
    oc3dceeTitlesInterior = 125,        /* . */
    oc3dceeTitlesFont = 126,            /* . */
    oc3dceeTitlesImage = 127,           /* . */
    oc3dceeLegendGeneral = 141,         /* . */
    oc3dceeLegendLabels = 142,          /* . */
    oc3dceeLegendLocation = 143,        /* . */
    oc3dceeLegendBorder = 144,          /* . */
    oc3dceeLegendInterior = 145,        /* . */
    oc3dceeLegendFont = 146,            /* . */
    oc3dceeLegendImage = 147,           /* . */
    oc3dceeLegendTitle = 148,           /* . */
    oc3dceeChartAreaLocation = 162,     /* . */
    oc3dceeChartAreaBorder = 163,       /* . */
    oc3dceeChartAreaInterior = 164,     /* . */
    oc3dceeChartAreaImage = 165,        /* . */
    oc3dceePlotCubeGeneral = 181,       /* . */
    oc3dceePlotCubeCeiling = 182,       /* . */
    oc3dceePlotCubeFloor = 183,         /* . */
    oc3dceePlotCubeViewport = 184,      /* . */
    oc3dceePlotCubeInterior = 185,      /* . */
    oc3dceeChartLabelsGeneral = 201,    /* . */
    oc3dceeChartLabelsAttach = 202,     /* . */
    oc3dceeChartLabelsLabel = 203,      /* . */
    oc3dceeChartLabelsBorder = 204,     /* . */
    oc3dceeChartLabelsInterior = 205,   /* . */
    oc3dceeChartLabelsFont = 206,       /* . */
    oc3dceeChartLabelsStrokeFont = 207,  /* . */
    oc3dceeChartLabelsImage = 208,      /* . */
    oc3dceeBarGeneral = 221,            /* . */
    oc3dceeBarColors = 222,             /* . */
    oc3dceeSurfaceGeneral = 241,        /* . */
	oc3dceeScatterGeneral = 242,      /* . */
    oc3dceeView3DGeneral = 261,         /* . */
    oc3dceeLevelsGeneral = 281,         /* . */
    oc3dceeLevelsAdd = 282,             /* . */
    oc3dceeLevelsRemove = 283,          /* . */
    oc3dceeControlGeneralIsBatched = 299,    /* . */
    oc3dceeControlGeneralIsDoubleBuffered = 300,    /* . */
    oc3dceeControlGeneralLoad = 301,    /* . */
    oc3dceeControlGeneralSave = 302,    /* . */
    oc3dceeControlGeneralReset = 303,    /* . */
    oc3dceeChartGroupDataLoad = 341,    /* . */
    oc3dceeChartGroupDataSave = 342,    /* . */
    oc3dceeChartGroupDataSmooth = 343,  /* . */
    oc3dceeChartGroupDataEdit = 344,    /* . */
    oc3dceeChartGroupDataRowLayout = 345,  /* . */
    oc3dceeChartGroupDataColLayout = 346,  /* . */
    oc3dceeStylesGeneralLoad = 381,     /* . */
    oc3dceeStylesGeneralSave = 382,     /* . */
    oc3dceeStylesGeneralReset = 383,    /* . */
    oc3dceeLevelsGeneralLoad = 421,     /* . */
    oc3dceeLevelsGeneralSave = 422,     /* . */
    oc3dceeChartStylesLineStyle = 431,  /* . */
    oc3dceeChartStylesSymbolStyle = 432,  /* . */
    oc3dceeChartStylesAdd = 433,        /* . */
    oc3dceeChartStylesRemove = 434,     /* . */
};

enum /* ImageLayoutConstants */         /* Specifies the image layout. */
{
    oc3dImageCentered = 1,              /* Centered image */
    oc3dImageTiled = 2,                 /* Tiled image */
    oc3dImageFitted = 3,                /* Fitted image */
    oc3dImageStretched = 4,             /* Stretched image */
    oc3dImageStretchedToWidth = 5,      /* Stretched to width image */
    oc3dImageStretchedToHeight = 6,     /* Stretched to height image */
    oc3dImageCropFitted = 7,            /* Crop fitted image */
};

enum /* ShapeConstants */               /* Valid symbol shapes. */
{
    oc3dShapeNone = 1,                  /* None */
    oc3dShapeDot = 2,                   /* Dot */
    oc3dShapeBox = 3,                   /* Box */
    oc3dShapeTriangle = 4,              /* Triangle */
    oc3dShapeDiamond = 5,               /* Diamond */
    oc3dShapeStar = 6,                  /* Star */
    oc3dShapeVerticalLine = 7,          /* Vertical Line */
    oc3dShapeHorizontalLine = 8,        /* Horizontal Line */
    oc3dShapeCross = 9,                 /* Cross */
    oc3dShapeCircle = 10,               /* Circle */
    oc3dShapeSquare = 11,               /* Square */
    oc3dShapeInvertTriangle = 12,       /* Invert Triangle */
    oc3dShapeDiagonalCross = 13,        /* Diagonal Cross */
    oc3dShapeOpenTriangle = 14,         /* Open Triangle */
    oc3dShapeOpenDiamond = 15,          /* Open Diamond */
    oc3dShapeOpenInvertTriangle = 16,   /* Open Invert Triangle */
};

