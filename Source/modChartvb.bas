Attribute VB_Name = "RMChartDeclares"
    Enum RMC_Colors
    AliceBlue = &HFFF0F8FF
    AntiqueWhite = &HFFFAEBD7
    Aquamarine = &HFF7FFFD4
    ArmyGreen = &HFF669966
    AutumnOrange = &HFFFF6633
    AvocadoGreen = &HFF669933
    Azure = &HFFF0FFFF
    BabyBlue = &HFF6699FF
    BananaYellow = &HFFCCCC33
    Beige = &HFFF5F5DC
    Bisque = &HFFFFE4C4
    Black = &HFF000000
    BlanchedAlmond = &HFFFFEBCD
    Blue = &HFF0000FF
    BlueViolet = &HFF8A2BE2
    Brown = &HFFA52A2A
    BurlyWood = &HFFDEB887
    CadetBlue = &HFF5F9EA0
    Chalk = &HFFFFFF99
    Chartreuse = &HFF7FFF00
    Chocolate = &HFFD2691E
    Coral = &HFFFF7F50
    CornflowerBlue = &HFF6495ED
    Cornsilk = &HFFFFF8DC
    Crimson = &HFFDC143C
    Cyan = &HFF00FFFF
    DarkBlue = &HFF00008B
    DarkBrown = &HFF663333
    DarkCrimson = &HFF993366
    DarkCyan = &HFF008B8B
    DarkGold = &HFFCC9933
    DarkGoldenrod = &HFFB8860B
    DarkGray = &HFFA9A9A9
    DarkGreen = &HFF006400
    DarkKhaki = &HFFBDB76B
    DarkMagenta = &HFF8B008B
    DarkOliveGreen = &HFF556B2F
    DarkOrange = &HFFFF8C00
    DarkOrchid = &HFF9932CC
    DarkRed = &HFF8B0000
    DarkSalmon = &HFFE9967A
    DarkSeaGreen = &HFF8FBC8B
    DarkSlateBlue = &HFF483D8B
    DarkSlateGray = &HFF2F4F4F
    DarkTurquoise = &HFF00CED1
    DarkViolet = &HFF9400D3
    DeepAzure = &HFF6633FF
    DeepPink = &HFFFF1493
    DeepPurple = &HFF330066
    DeepRiver = &HFF6600CC
    DeepRose = &HFFCC3399
    DeepSkyBlue = &HFF00BFFF
    DeepYellow = &HFFFFCC00
    DesertBlue = &HFF336699
    DimGray = &HFF696969
    DodgerBlue = &HFF1E90FF
    DullGreen = &HFF99CC66
    EasterPurple = &HFFCC99FF
    FadeGreen = &HFF99CC99
    Firebrick = &HFFB22222
    FloralWhite = &HFFFFFAF0
    ForestGreen = &HFF228B22
    Gainsboro = &HFFDCDCDC
    GhostWhite = &HFFF8F8FF
    GhostGreen = &HFFCCFFCC
    Gold = &HFFFFD700
    Goldenrod = &HFFDAA520
    Grape = &HFF663399
    GrassGreen = &HFF009933
    Gray = &HFF808080
    Green = &HFF008000
    GreenYellow = &HFFADFF2F
    Honeydew = &HFFF0FFF0
    HotPink = &HFFFF69B4
    IndianRed = &HFFCD5C5C
    Indigo = &HFF4B0082
    Ivory = &HFFFFFFF0
    Khaki = &HFFF0E68C
    KentuckyGreen = &HFF339966
    Lavender = &HFFE6E6FA
    LavenderBlush = &HFFFFF0F5
    LawnGreen = &HFF7CFC00
    LemonChiffon = &HFFFFFACD
    LightBlue = &HFFADD8E6
    LightCoral = &HFFF08080
    LightCyan = &HFFE0FFFF
    LightGoldenrod = &HFFEEDD82
    LightGoldenrodYellow = &HFFFAFAD2
    LightGray = &HFFD3D3D3
    LightGreen = &HFF90EE90
    LightOrange = &HFFFF9933
    LightPink = &HFFFFB6C1
    LightSalmon = &HFFFFA07A
    LightSeaGreen = &HFF20B2AA
    LightSkyBlue = &HFF87CEFA
    LightSlateGray = &HFF778899
    LightSteelBlue = &HFFB0C4DE
    LightViolet = &HFFFF99FF
    LightYellow = &HFFFFFFE0
    Lime = &HFF00FF00
    LimeGreen = &HFF32CD32
    Linen = &HFFFAF0E6
    Magenta = &HFFFF00FF
    Maroon = &HFF800000
    MartianGreen = &HFF99CC33
    MediumAquamarine = &HFF66CDAA
    MediumBlue = &HFF0000CD
    MediumOrchid = &HFFBA55D3
    MediumPurple = &HFF9370DB
    MediumSeaGreen = &HFF3CB371
    MediumSlateBlue = &HFF7B68EE
    MediumSpringGreen = &HFF00FA9A
    MediumTurquoise = &HFF48D1CC
    MediumVioletRed = &HFFC71585
    MidnightBlue = &HFF191970
    MintCream = &HFFF5FFFA
    MistyRose = &HFFFFE4E1
    Moccasin = &HFFFFE4B5
    MoonGreen = &HFFCCFF66
    MossGreen = &HFF336666
    NavajoWhite = &HFFFFDEAD
    Navy = &HFF000080
    OceanGreen = &HFF669999
    OldLace = &HFFFDF5E6
    Olive = &HFF808000
    OliveDrab = &HFF6B8E23
    Orange = &HFFFFA500
    OrangeRed = &HFFFF4500
    Orchid = &HFFDA70D6
    PaleGoldenrod = &HFFEEE8AA
    PaleGreen = &HFF98FB98
    PaleTurquoise = &HFFAFEEEE
    PaleVioletRed = &HFFDB7093
    PaleYellow = &HFFFFFFCC
    PapayaWhip = &HFFFFEFD5
    PeachPuff = &HFFFFDAB9
    Peru = &HFFCD853F
    Pink = &HFFFFC0CB
    Plum = &HFFDDA0DD
    PowderBlue = &HFFB0E0E6
    Purple = &HFF800080
    Red = &HFFFF0000
    RosyBrown = &HFFBC8F8F
    RoyalBlue = &HFF4169E1
    SaddleBrown = &HFF8B4513
    Salmon = &HFFFA8072
    Sand = &HFFFFCC99
    SandyBrown = &HFFF4A460
    SeaGreen = &HFF2E8B57
    SeaShell = &HFFFFF5EE
    Sienna = &HFFA0522D
    Silver = &HFFC0C0C0
    SkyBlue = &HFF87CEEB
    SlateBlue = &HFF6A5ACD
    SlateGray = &HFF708090
    Snow = &HFFFFFAFA
    SpringGreen = &HFF00FF7F
    SteelBlue = &HFF4682B4
    Tan = &HFFD2B48C
    Teal = &HFF008080
    Thistle = &HFFD8BFD8
    Tomato = &HFFFF6347
    Transparent = &HFFFFFFFE
    TropicalPink = &HFFFF6666
    Turquoise = &HFF40E0D0
    Violet = &HFFEE82EE
    VioletRed = &HFFD02090
    Walnut = &HFF663300
    Wheat = &HFFF5DEB3
    White = &HFFFFFFFF
    WhiteSmoke = &HFFF5F5F5
    Yellow = &HFFFFFF00
    YellowGreen = &HFF9ACD32
End Enum

' No data symbol
Public Const RMC_NO_DATA = &HC521974F

Enum CtrlStyle
    RMC_CTRLSTYLEFLAT = 0
    RMC_CTRLSTYLEFLATSHADOW = 1
    RMC_CTRLSTYLE3D = 2
    RMC_CTRLSTYLE3DLIGHT = 3
    RMC_CTRLSTYLEIMAGE = 4
    RMC_CTRLSTYLEIMAGETILED = 5
End Enum

Enum SeriesType
    RMC_BARSERIES = 1
    RMC_LINESERIES = 2
    RMC_GRIDLESSSERIES = 2
    RMC_VOLUMEBARSERIES = 4
    RMC_HIGHLOWSERIES = 5
    RMC_XYSERIES = 6
End Enum

Enum BarSeriesType
    RMC_BARSINGLE = 1
    RMC_BARGROUP = 2
    RMC_BARSTACKED = 3
    RMC_BARSTACKED100 = 4
    RMC_FLOATINGBAR = 5
    RMC_FLOATINGBARGROUP = 6
End Enum

Enum LineSeriesType
    RMC_LINE = 21
    RMC_AREA = 22
    RMC_LINE_INDEXED = 23
    RMC_AREA_INDEXED = 24
    RMC_AREA_STACKED = 25
    RMC_AREA_STACKED100 = 26
End Enum

Enum BarSeriesStyle
    RMC_BAR_FLAT = 1
    RMC_BAR_FLAT_GRADIENT1 = 2
    RMC_BAR_FLAT_GRADIENT2 = 3
    RMC_BAR_HOVER = 4
    RMC_COLUMN_FLAT = 5
    RMC_BAR_3D = 6
    RMC_BAR_3D_GRADIENT = 7
    RMC_COLUMN_3D = 8
    RMC_COLUMN_3D_GRADIENT = 9
    RMC_COLUMN_FLUTED = 10
End Enum

Enum CTypes                       ' only for tRMC_INFO
    RMC_VOLUMEBAR = 31
    RMC_HIGHLOW = 41
    RMC_GRIDLESS = 51
    RMC_XYCHART = 70
    RMC_GRIDBASED = 10
End Enum

Enum LineSeriesStyle
    RMC_LINE_FLAT = 21
    RMC_LINE_FLAT_DOT = 19
    RMC_LINE_FLAT_DASH = 18
    RMC_LINE_FLAT_DASHDOT = 17
    RMC_LINE_FLAT_DASHDOTDOT = 16
    RMC_LINE_FASTLINE = 15
    RMC_LINE_CABLE = 22
    RMC_LINE_3D = 23
    RMC_LINE_3D_GRADIENT = 24
    RMC_AREA_FLAT = 25
    RMC_AREA_FLAT_GRADIENT_V = 26
    RMC_AREA_FLAT_GRADIENT_H = 27
    RMC_AREA_FLAT_GRADIENT_C = 28
    RMC_AREA_3D = 29
    RMC_AREA_3D_GRADIENT_V = 30
    RMC_AREA_3D_GRADIENT_H = 31
    RMC_AREA_3D_GRADIENT_C = 32
    RMC_LINE_FLAT_SHADOW = 33
    RMC_LINE_CABLE_SHADOW = 34
    RMC_LINE_SYMBOLONLY = 35
End Enum

Enum LineSeriesLineStyle
    RMC_LSTYLE_LINE = 1
    RMC_LSTYLE_SPLINE = 2
    RMC_LSTYLE_STAIR = 3
    RMC_LSTYLE_LINE_AREA = 4       ' Draws a line and a transparent area
    RMC_LSTYLE_SPLINE_AREA = 5     ' Draws a spline and a transparent area
    RMC_LSTYLE_STAIR_AREA = 6      ' Draws a stair and a transparent area
End Enum

Enum LineSeriesSymbol
    RMC_SYMBOL_NONE = 0
    RMC_SYMBOL_BULLET = 21
    RMC_SYMBOL_ROUND = 1
    RMC_SYMBOL_DIAMOND = 2
    RMC_SYMBOL_SQUARE = 3
    RMC_SYMBOL_STAR = 4
    RMC_SYMBOL_ARROW_DOWN = 5
    RMC_SYMBOL_ARROW_UP = 6
    RMC_SYMBOL_POINT = 7
    RMC_SYMBOL_CIRCLE = 8
    RMC_SYMBOL_RECTANGLE = 9
    RMC_SYMBOL_CROSS = 10
    RMC_SYMBOL_BULLET_SMALL = 22
    RMC_SYMBOL_ROUND_SMALL = 11
    RMC_SYMBOL_DIAMOND_SMALL = 12
    RMC_SYMBOL_SQUARE_SMALL = 13
    RMC_SYMBOL_STAR_SMALL = 14
    RMC_SYMBOL_ARROW_DOWN_SMALL = 15
    RMC_SYMBOL_ARROW_UP_SMALL = 16
    RMC_SYMBOL_POINT_SMALL = 17
    RMC_SYMBOL_CIRCLE_SMALL = 18
    RMC_SYMBOL_RECTANGLE_SMALL = 19
    RMC_SYMBOL_CROSS_SMALL = 20
End Enum

Enum HighLowSeriesStyle
    RMC_OHLC = 1
    RMC_CANDLESTICK = 2
End Enum

Enum GridlessSeriesStyle
    RMC_PIE_FLAT = 51
    RMC_PIE_GRADIENT = 52
    RMC_PIE_3D = 53
    RMC_PIE_3D_GRADIENT = 54
    RMC_DONUT_FLAT = 55
    RMC_DONUT_GRADIENT = 56
    RMC_DONUT_3D = 57
    RMC_DONUT_3D_GRADIENT = 58
    RMC_PYRAMIDE = 59
    RMC_PYRAMIDE3 = 60
End Enum

Enum PieDonutAlignment
    RMC_FULL = 1
    RMC_HALF_TOP = 2
    RMC_HALF_RIGHT = 3
    RMC_HALF_BOTTOM = 4
    RMC_HALF_LEFT = 5
End Enum

Enum XYSeriesStyle
    RMC_XY_LINE = 70
    RMC_XY_LINE_DOT = 69
    RMC_XY_LINE_DASH = 68
    RMC_XY_LINE_DASHDOT = 67
    RMC_XY_LINE_DASHDOTDOT = 66
    RMC_XY_SYMBOL = 71
    RMC_XY_CABLE = 73
End Enum

Enum Hatchmodes
    RMC_HATCHBRUSH_OFF = 0
    RMC_HATCHBRUSH_ON = 1
    RMC_HATCHBRUSH_ONPRINTING = 2
End Enum

Enum DAxisAlignment
    RMC_DATAAXISLEFT = 1
    RMC_DATAAXISRIGHT = 2
    RMC_DATAAXISTOP = 3
    RMC_DATAAXISBOTTOM = 4
End Enum

Enum LAxisAlignment
    RMC_LABELAXISLEFT = 5
    RMC_LABELAXISRIGHT = 6
    RMC_LABELAXISTOP = 7
    RMC_LABELAXISBOTTOM = 8
End Enum

Enum XAxisAlignment
    RMC_XAXISTOP = 11
    RMC_XAXISBOTTOM = 12
End Enum

Enum YAxisAlignment
    RMC_YAXISLEFT = 9
    RMC_YAXISRIGHT = 10
End Enum

Enum AxisType
    RMC_DATAAXIS = 1
    RMC_LABELAXIS = 2
End Enum

Enum AxisLineStyle
    RMC_LINESTYLESOLID = 0
    RMC_LINESTYLEDASH = 1
    RMC_LINESTYLEDOT = 2
    RMC_LINESTYLEDASHDOT = 3
    RMC_LINESTYLENONE = 6
End Enum

Enum LabelTextAlignment
    RMC_TEXTCENTER = 0
    RMC_TEXTLEFT = 1
    RMC_TEXTRIGHT = 2
    RMC_TEXTDOWNWARD = 3
    RMC_TEXTUPWARD = 4
End Enum

Enum LegendAlignment
    RMC_LEGEND_NONE = -1
    RMC_LEGEND_TOP = 1
    RMC_LEGEND_LEFT = 2
    RMC_LEGEND_RIGHT = 3
    RMC_LEGEND_BOTTOM = 4
    RMC_LEGEND_UL = 5
    RMC_LEGEND_UR = 6
    RMC_LEGEND_LL = 7
    RMC_LEGEND_LR = 8
    RMC_LEGEND_ONVLABELS = 9
    RMC_LEGEND_CUSTOM_TOP = 11
    RMC_LEGEND_CUSTOM_LEFT = 12
    RMC_LEGEND_CUSTOM_RIGHT = 13
    RMC_LEGEND_CUSTOM_BOTTOM = 14
    RMC_LEGEND_CUSTOM_UL = 15
    RMC_LEGEND_CUSTOM_UR = 16
    RMC_LEGEND_CUSTOM_LL = 17
    RMC_LEGEND_CUSTOM_LR = 18
    RMC_LEGEND_CUSTOM_CENTER = 19
    RMC_LEGEND_CUSTOM_CR = 20
    RMC_LEGEND_CUSTOM_CL = 21
End Enum

Enum LegendStyle
    RMC_LEGENDNORECT = 1
    RMC_LEGENDRECT = 2
    RMC_LEGENDRECTSHADOW = 3
    RMC_LEGENDROUNDRECT = 4
    RMC_LEGENDROUNDRECTSHADOW = 5
End Enum



Enum ValueLabels
    RMC_VLABEL_NONE = 0
    RMC_VLABEL_DEFAULT = 1
    RMC_VLABEL_PERCENT = 5
    RMC_VLABEL_ABSOLUTE = 6
    RMC_VLABEL_TWIN = 7
    RMC_VLABEL_LEGENDONLY = 8
    RMC_VLABEL_DEFAULT_NOZERO = 11
    RMC_VLABEL_PERCENT_NOZERO = 15
    RMC_VLABEL_ABSOLUTE_NOZERO = 16
    RMC_VLABEL_TWIN_NOZERO = 17
End Enum

Enum BicolorMode
    RMC_BICOLOR_NONE = 0
    RMC_BICOLOR_DATAAXIS = 1
    RMC_BICOLOR_LABELAXIS = 2
    RMC_BICOLOR_BOTH = 3
End Enum

Enum RMCError
    RMC_ERROR_MAXINST = -1
    RMC_ERROR_MAXREGION = -2
    RMC_ERROR_MAXSERIES = -3
    RMC_ERROR_ALLOC = -4
    RMC_ERROR_NODATA = -5
    RMC_ERROR_CTRLID = -6
    RMC_ERROR_SERIESINDEX = -7
    RMC_ERROR_CREATEBITMAP = -8
    RMC_ERROR_WRONGREGION = -9
    RMC_ERROR_PARENTHANDLE = -10
    RMC_ERROR_CREATEWINDOW = -11
    RMC_ERROR_INIGDIP = -12
    RMC_ERROR_PRINT = -13
    RMC_ERROR_NOGDIP = -14
    RMC_ERROR_RMCFILE = -15
    RMC_ERROR_FILEFOUND = -16
    RMC_ERROR_READLINES = -17
    RMC_ERROR_XYAXIS = -18
    RMC_ERROR_LEGENDTEXT = -19
    RMC_ERROR_EMF = -20
    RMC_ERROR_NODATA_COUNT = -21
    RMC_ERROR_NODATA_ZERO = -22
    RMC_ERROR_NOCOLOR = -23
    RMC_ERROR_CLIPBOARD = -24
    RMC_ERROR_CBINFO = -25
    RMC_ERROR_FILECREATE = -26
    RMC_ERROR_DATAINDEX = -28
    RMC_ERROR_AXISALIGNMENT = -29
    RMC_ERROR_RANGE = -30
    RMC_ERROR_WRONGSERIESTYPE = -31
    RMC_ERROR_MAXCUSTOM = -50
    RMC_ERROR_CUSTOMINDEX = -51
    RMC_ERROR_LEGENDSIZE = 1
End Enum

Enum RMCFileType
    RMC_EMF = 1
    RMC_EMFPLUS = 2
    RMC_BMP = 3
End Enum

' Custom Objects
Enum COType
    RMC_CO_TEXT = 1
    RMC_CO_BOX = 2
    RMC_CO_CIRCLE = 3
    RMC_CO_LINE = 4
    RMC_CO_IMAGE = 5
    RMC_CO_SYMBOL = 6
    RMC_CO_POLYGON = 7
End Enum

' Line alignment for custom text
Enum COLineAlignment
    RMC_LINE_HORIZONTAL = 0
    RMC_LINE_UPWARD = 1
    RMC_LINE_DOWNWARD = 3
End Enum

'Line style for Custom lines
Enum COLineStyle
    RMC_FLAT_LINE = 21
    RMC_DOT_LINE = 19
    RMC_DASH_LINE = 18
    RMC_DASHDOT_LINE = 17
    RMC_DASHDOTDOT_LINE = 16
End Enum

' Anchors for custom lines
Enum COAnchor
    RMC_ANCHOR_NONE = 0
    RMC_ANCHOR_ROUND = 1
    RMC_ANCHOR_BULLET = 2
    RMC_ANCHOR_ARROW_CLOSED = 3
    RMC_ANCHOR_ARROW_OPEN = 4
End Enum

' Styles for custom box/text
Enum COBoxStyle
    RMC_BOX_NONE = 0
    RMC_BOX_FLAT = 1
    RMC_BOX_ROUNDEDGE = 2
    RMC_BOX_RHOMBUS = 3
    RMC_BOX_GRADIENTH = 4
    RMC_BOX_GRADIENTV = 5
    RMC_BOX_3D = 6
    RMC_BOX_FLAT_SHADOW = 7
    RMC_BOX_GRADIENTH_SHADOW = 8
    RMC_BOX_GRADIENTV_SHADOW = 9
    RMC_BOX_3D_SHADOW = 10
End Enum

' Styles for custom Circle
Enum COCircleStyle
    RMC_CIRCLE_FLAT = 1
    RMC_CIRCLE_BULLET = 2
End Enum

' Zoom mode
Enum ZoomMode
    RMC_ZOOM_DISABLE = 0
    RMC_ZOOM_EXTERNAL = 1
    RMC_ZOOM_INTERNAL = 2
End Enum

' nChartType in tRMC_INFO holds one of these when in zoom- or magnifier-mode
Enum ZoomState
    RMC_ZOOM_MODE = -99
    RMC_MAGNIFIER_MODE = -98
End Enum

Enum RMCMouseButton
    RMC_MOUSEMOVE = &H200
    RMC_LBUTTONDOWN = &H201
    RMC_LBUTTONUP = &H202
    RMC_LBUTTONDBLCLK = &H203
    RMC_RBUTTONDOWN = &H204
    RMC_RBUTTONUP = &H205
    RMC_RBUTTONDBLCLK = &H206
    RMC_MBUTTONDOWN = &H207
    RMC_MBUTTONUP = &H208
    RMC_MBUTTONDBLCLK = &H209
    RMC_SHIFTLBUTTONDOWN = &H20A
    RMC_SHIFTLBUTTONUP = &H20B
    RMC_SHIFTLBUTTONDBLCLK = &H20C
    RMC_SHIFTRBUTTONDOWN = &H20D
    RMC_SHIFTRBUTTONUP = &H20E
    RMC_SHIFTRBUTTONDBLCLK = &H20F
    RMC_SHIFTMBUTTONDOWN = &H210
    RMC_SHIFTMBUTTONUP = &H211
    RMC_SHIFTMBUTTONDBLCLK = &H212
    RMC_CTRLLBUTTONDOWN = &H213
    RMC_CTRLLBUTTONUP = &H214
    RMC_CTRLLBUTTONDBLCLK = &H215
    RMC_CTRLRBUTTONDOWN = &H216
    RMC_CTRLRBUTTONUP = &H217
    RMC_CTRLRBUTTONDBLCLK = &H218
    RMC_CTRLMBUTTONDOWN = &H219
    RMC_CTRLMBUTTONUP = &H21A
    RMC_CTRLMBUTTONDBLCLK = &H21B
End Enum

Type tRMC_INFO
    nXPos        As Long
    nYPos        As Long
    nXMove       As Long
    nYMove       As Long
    nRegionIndex As Long
    nRLeft       As Long
    nRTop        As Long
    nRRight      As Long
    nRBottom     As Long
    nSeriesIndex As Long
    nDataIndex   As Long
    nChartType   As Long
    nSLeft       As Long
    nSTop        As Long
    nSRight      As Long
    nSBottom     As Long
    nSTop2       As Long
    nSBottom2    As Long
    nGLeft       As Long
    nGTop        As Long
    nGRight      As Long
    nGBottom     As Long
    nGCol        As Long
    nGRow        As Long
    nData1       As Double
    nData2       As Double
    nData3       As Double
    nData4       As Double
    nVirtData1   As Double
    nVirtData2   As Double
    nVirtData3   As Double
    nVirtData4   As Double
End Type


Type tRMC_BARSERIES
    nType As BarSeriesType
    nStyle As BarSeriesStyle
    nIsLucent As Long
    nColor As RMC_Colors
    nIsHorizontal As Long
    nWhichDataAxis As Long
    nValueLabelOn As ValueLabels
    nPointsPerColumn As Long
    nHatchMode As Hatchmodes
End Type

Type tRMC_CAPTION
    nBackColor As RMC_Colors
    nTextColor As RMC_Colors
    nFontSize As Long
    nIsBold As Long
    sText As String * 200
End Type

Type tRMC_CHART
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    nBackColor As RMC_Colors
    nCtrlStyle As CtrlStyle
    nExportOnly As Long
    sBgImage As String * 100
    sFontName As String * 50
    nToolTipWidth   As Long
    nBitmapBKColor  As Long
End Type

Type tRMC_DATAAXIS
    nAlignment As DAxisAlignment
    nMinValue As Double
    nMaxValue As Double
    nTickCount As Long
    nFontSize As Long
    nTextColor As RMC_Colors
    nLineColor As RMC_Colors
    nLinestyle As AxisLineStyle
    nDecimalDigits As Long
    nLabelAlignment As Long
    sUnit As String * 16
    sText As String * 100
    sLabels As String * 500
End Type

Type tRMC_GRID
    nGridBackColor As RMC_Colors
    nAsGradient As Long
    nBiColor As BicolorMode
    nLeft As Long
    nTop As Long
    nWidth As Long
    nHeight As Long
End Type

Type tRMC_GRIDLESSSERIES
    nStyle As GridlessSeriesStyle
    nPieAlignment As PieDonutAlignment
    nExplodeMode As Long
    nIsLucent As Long
    nValueLabelOn As ValueLabels
    nHatchMode As Hatchmodes
    nStartAngle As Long
End Type

Type tRMC_LABELAXIS
    nCount As Long
    nTickCount As Long
    nAlignment As LAxisAlignment
    nFontSize As Long
    nTextColor As RMC_Colors
    nTextAlignment As LAxisAlignment
    nLineColor As RMC_Colors
    nLinestyle As AxisLineStyle
    sText As String * 100
End Type

Type tRMC_LEGEND
    nLegendAlign As LegendAlignment
    nLegendBackColor As RMC_Colors
    nLegendStyle As LegendStyle
    nLegendTextColor As RMC_Colors
    nLegendFontSize As Long
    nLegendIsBold As Long
End Type

Type tRMC_LINESERIES
    nType As LineSeriesType
    nStyle As LineSeriesStyle
    nLinestyle As LineSeriesLineStyle
    nIsLucent As Long
    nColor As RMC_Colors
    nSeriesSymbol As LineSeriesSymbol
    nWhichDataAxis As Long
    nValueLabelOn As ValueLabels
    nHatchMode As Hatchmodes
End Type

Type tRMC_REGION
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    sFooter As String * 200
    nShowBorder As Long
End Type

Type tRMC_XYAXIS
    nAlignment      As Long
    nMinValue       As Double
    nMaxValue       As Double
    nTickCount      As Long
    nFontSize       As Long
    nTextColor      As RMC_Colors
    nLineColor      As RMC_Colors
    nLinestyle      As AxisLineStyle
    nDecimalDigits  As Long
    nLabelAlignment As Long
    sUnit           As String * 16
    sText           As String * 100
    sLabels         As String * 500
End Type

Type tRMC_XYSERIES
    nColor              As RMC_Colors
    nStyle              As XYSeriesStyle
    nLinestyle          As LineSeriesLineStyle
    nSeriesSymbol       As LineSeriesSymbol
    nWhichXAxis         As Long
    nWhichYAxis         As Long
    nValueLabelOn       As ValueLabels
        
End Type

Declare Function RMC_AddBarSeries Lib "RMCHART.DLL" Alias "RMC_ADDBARSERIES" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByRef nFirstDataValue As Double, _
                            Optional ByVal nDataValuesCount As Long, _
                            Optional ByVal nType As BarSeriesType, _
                            Optional ByVal nStyle As BarSeriesStyle, _
                            Optional ByVal nIsLucent As Long, _
                            Optional ByVal nColor As RMC_Colors, _
                            Optional ByVal nIsHorizontal As Long, _
                            Optional ByVal nWhichDataAxis As Long, _
                            Optional ByVal nValueLabelOn As ValueLabels, _
                            Optional ByVal nPointsPerColumn As Long, _
                            Optional ByVal nHatchMode As Hatchmodes _
                            ) As RMCError
Declare Function RMC_AddBarSeriesI Lib "RMCHART.DLL" Alias "RMC_ADDBARSERIESI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByRef nFirstDataValue As Double, _
                            ByVal nDataValuesCount As Long, _
                            T As tRMC_BARSERIES _
                            ) As RMCError

Declare Function RMC_AddCaption Lib "RMCHART.DLL" Alias "RMC_ADDCAPTION" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByVal sCaption As String, _
                            Optional ByVal nTitelBackColor As RMC_Colors, _
                            Optional ByVal nTitelTextColor As RMC_Colors, _
                            Optional ByVal nTitelFontSize As Long, _
                            Optional ByVal nTitelIsBold As Long _
                            ) As RMCError
Declare Function RMC_AddCaptionI Lib "RMCHART.DLL" Alias "RMC_ADDCAPTIONI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            T As tRMC_CAPTION _
                            ) As RMCError

Declare Function RMC_AddDataAxis Lib "RMCHART.DLL" Alias "RMC_ADDDATAAXIS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByVal nAlignment As DAxisAlignment, _
                            Optional ByVal nMinValue As Double, _
                            Optional ByVal nMaxValue As Double, _
                            Optional ByVal nTickCount As Long, _
                            Optional ByVal nFontSize As Long, _
                            Optional ByVal nTextColor As RMC_Colors, _
                            Optional ByVal nLineColor As RMC_Colors, _
                            Optional ByVal nLinestyle As AxisLineStyle, _
                            Optional ByVal nDecimalDigits As Long, _
                            Optional ByVal sUnit As String, _
                            Optional ByVal sText As String, _
                            Optional ByVal sLabels As String, _
                            Optional ByRef nLabelAlignment As Long _
                            ) As RMCError
Declare Function RMC_AddDataAxisI Lib "RMCHART.DLL" Alias "RMC_ADDDATAAXISI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            T As tRMC_DATAAXIS _
                            ) As RMCError

Declare Function RMC_AddGrid Lib "RMCHART.DLL" Alias "RMC_ADDGRID" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByVal nBackColor As RMC_Colors, _
                            Optional ByVal nAsGradient As Long, _
                            Optional ByVal nLeft As Long, _
                            Optional ByVal nTop As Long, _
                            Optional ByVal nWidth As Long, _
                            Optional ByVal nHeight As Long, _
                            Optional ByVal nBiColor As BicolorMode _
                            ) As RMCError
Declare Function RMC_AddGridI Lib "RMCHART.DLL" Alias "RMC_ADDGRIDI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            T As tRMC_GRID _
                            ) As RMCError
                            
Declare Function RMC_AddGridlessSeries Lib "RMCHART.DLL" Alias "RMC_ADDGRIDLESSSERIES" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByRef nFirstDataValue As Double, _
                            Optional ByVal nDataValuesCount As Long, _
                            Optional ByRef nFirstColorElement As Long, _
                            Optional ByVal nColorElementsCount As Long, _
                            Optional ByVal nStyle As GridlessSeriesStyle, _
                            Optional ByVal nAlignment As PieDonutAlignment, _
                            Optional ByVal nExplodeMode As Long, _
                            Optional ByVal nIsLucent As Long, _
                            Optional ByVal nValueLabelOn As ValueLabels, _
                            Optional ByVal nHatchMode As Hatchmodes, _
                            Optional ByVal nStartAngle As Long _
                            ) As RMCError
Declare Function RMC_AddGridlessSeriesI Lib "RMCHART.DLL" Alias "RMC_ADDGRIDLESSSERIESI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByRef nFirstDataValue As Double, _
                            ByVal nDataValuesCount As Long, _
                            ByRef nFirstColorElement As Long, _
                            ByVal nColorElementsCount As Long, _
                            ByRef T As tRMC_GRIDLESSSERIES _
                            ) As RMCError
                            
Declare Function RMC_AddHighLowSeries Lib "RMCHART.DLL" Alias "RMC_ADDHIGHLOWSERIES" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByRef nFirstDataValue As Double, _
                            Optional ByVal nDataValuesCount As Long, _
                            Optional ByRef nFirstPPCValue As Long, _
                            Optional ByVal nPPCValuesCount As Long, _
                            Optional ByVal nStyle As HighLowSeriesStyle, _
                            Optional ByVal nWhichDataAxis As Long, _
                            Optional ByVal nColorLow As Long, _
                            Optional ByVal nColorHigh As Long _
                             ) As RMCError

Declare Function RMC_AddLabelAxis Lib "RMCHART.DLL" Alias "RMC_ADDLABELAXIS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sLabels As String, _
                            Optional ByVal nCount As Long, _
                            Optional ByVal nTickCount As Long, _
                            Optional ByVal nAlignment As LAxisAlignment, _
                            Optional ByVal nFontSize As Long, _
                            Optional ByVal nTextColor As RMC_Colors, _
                            Optional ByVal nTextAlignment As LabelTextAlignment, _
                            Optional ByVal nLineColor As RMC_Colors, _
                            Optional ByVal nLinestyle As AxisLineStyle, _
                            Optional ByVal sText As String _
                            ) As RMCError
Declare Function RMC_AddLabelAxisI Lib "RMCHART.DLL" Alias "RMC_ADDLABELAXISI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sLabels As String, _
                            T As tRMC_LABELAXIS _
                            ) As RMCError
                            
Declare Function RMC_AddLegend Lib "RMCHART.DLL" Alias "RMC_ADDLEGEND" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sLegend As String, _
                            Optional ByVal nLegendAlign As LegendAlignment, _
                            Optional ByVal nLegendBackColor As RMC_Colors, _
                            Optional ByVal nLegendStyle As LegendStyle, _
                            Optional ByVal nLegendTextColor As RMC_Colors, _
                            Optional ByVal nLegendFontSize As Long, _
                            Optional ByVal nLegendIsBold As Long _
                            ) As RMCError
Declare Function RMC_AddLegendI Lib "RMCHART.DLL" Alias "RMC_ADDLEGENDI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sLegend As String, _
                            T As tRMC_LEGEND _
                            ) As RMCError
                            
Declare Function RMC_AddLineSeries Lib "RMCHART.DLL" Alias "RMC_ADDLINESERIES" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByRef nFirstDataValue As Double, _
                            Optional ByVal nDataValuesCount As Long, _
                            Optional ByRef nFirstPPCValue As Long, _
                            Optional ByVal nPPCValuesCount As Long, _
                            Optional ByVal nChartType As LineSeriesType, _
                            Optional ByVal nStyle As LineSeriesStyle, _
                            Optional ByVal nLinestyle As LineSeriesLineStyle, _
                            Optional ByVal nIsLucent As Long, _
                            Optional ByVal nColor As RMC_Colors, _
                            Optional ByVal nSeriesSymbol As LineSeriesSymbol, _
                            Optional ByVal nWhichDataAxis As Long, _
                            Optional ByVal nValueLabelOn As ValueLabels, _
                            Optional ByVal nHatchMode As Hatchmodes _
                            ) As RMCError
Declare Function RMC_AddLineSeriesI Lib "RMCHART.DLL" Alias "RMC_ADDLINESERIESI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByRef nFirstDataValue As Double, _
                            ByVal nDataValuesCount As Long, _
                            ByRef nFirstPPCValue As Long, _
                            ByVal nPPCValuesCount As Long, _
                            T As tRMC_LINESERIES _
                            ) As RMCError
                            
Declare Function RMC_AddRegion Lib "RMCHART.DLL" Alias "RMC_ADDREGION" ( _
                            ByVal nCtrlId As Long, _
                            Optional ByVal nLeft As Long, _
                            Optional ByVal nTop As Long, _
                            Optional ByVal nWidth As Long, _
                            Optional ByVal nHeight As Long, _
                            Optional ByVal sFooter As String, _
                            Optional ByVal nShowBorder As Long _
                            ) As RMCError
Declare Function RMC_AddRegionI Lib "RMCHART.DLL" Alias "RMC_ADDREGIONI" ( _
                            ByVal nCtrlId As Long, _
                            T As tRMC_REGION _
                            ) As RMCError
                            
Declare Function RMC_AddToolTips Lib "RMCHART.DLL" Alias "RMC_ADDTOOLTIPS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal hWnd As Long, _
                            Optional ByVal nToolTipWidth As Long _
                            ) As RMCError
                            
Declare Function RMC_AddVolumeBarSeries Lib "RMCHART.DLL" Alias "RMC_ADDVOLUMEBARSERIES" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByRef nFirstDataValue As Double, _
                            Optional ByVal nDataValuesCount As Long, _
                            Optional ByRef nFirstPPCValue As Long, _
                            Optional ByVal nPPCValuesCount As Long, _
                            Optional ByVal nColor As RMC_Colors, _
                            Optional ByVal nWhichDataAxis As Long _
                            ) As RMCError

Declare Function RMC_AddXAxis Lib "RMCHART.DLL" Alias "RMC_ADDXAXIS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nAlignment As XAxisAlignment, _
                            Optional ByVal nMinValue As Double, _
                            Optional ByVal nMaxValue As Double, _
                            Optional ByVal nTickCount As Long, _
                            Optional ByVal nFontSize As Long, _
                            Optional ByVal nTextColor As RMC_Colors, _
                            Optional ByVal nLineColor As RMC_Colors, _
                            Optional ByVal nLinestyle As AxisLineStyle, _
                            Optional ByVal nDecimalDigits As Long, _
                            Optional ByVal sUnit As String, _
                            Optional ByVal sText As String, _
                            Optional ByVal sLables As String, _
                            Optional ByVal nLabelAlignment As Long _
                            ) As RMCError
Declare Function RMC_AddXAxisI Lib "RMCHART.DLL" Alias "RMC_ADDXAXISI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            T As tRMC_XYAXIS _
                            ) As RMCError

Declare Function RMC_AddYAxis Lib "RMCHART.DLL" Alias "RMC_ADDYAXIS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nAlignment As YAxisAlignment, _
                            Optional ByVal nMinValue As Double, _
                            Optional ByVal nMaxValue As Double, _
                            Optional ByVal nTickCount As Long, _
                            Optional ByVal nFontSize As Long, _
                            Optional ByVal nTextColor As RMC_Colors, _
                            Optional ByVal nLineColor As RMC_Colors, _
                            Optional ByVal nLinestyle As AxisLineStyle, _
                            Optional ByVal nDecimalDigits As Long, _
                            Optional ByVal sUnit As String, _
                            Optional ByVal sText As String, _
                            Optional ByVal sLables As String, _
                            Optional ByVal nLabelAlignment As Long _
                            ) As RMCError
Declare Function RMC_AddYAxisI Lib "RMCHART.DLL" Alias "RMC_ADDYAXISI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            T As tRMC_XYAXIS _
                            ) As RMCError

Declare Function RMC_AddXYSeries Lib "RMCHART.DLL" Alias "RMC_ADDXYSERIES" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            Optional ByRef nFirstXDataValue As Double, _
                            Optional ByVal nDataXValuesCount As Long, _
                            Optional ByRef nFirstYDataValue As Double, _
                            Optional ByVal nDataYValuesCount As Long, _
                            Optional ByVal nColor As RMC_Colors, _
                            Optional ByVal nStyle As XYSeriesStyle, _
                            Optional ByVal nLinestyle As LineSeriesLineStyle, _
                            Optional ByVal nSymbolStyle As LineSeriesSymbol, _
                            Optional ByVal nWhichXAxis As Long, _
                            Optional ByVal nWhichYAxis As Long, _
                            Optional ByVal nValueLabelOn As ValueLabels _
                            ) As RMCError
Declare Function RMC_AddXYSeriesI Lib "RMCHART.DLL" Alias "RMC_ADDXYSERIESI" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByRef nDataX As Double, _
                            ByVal nDataXCount As Long, _
                            ByRef nDataY As Double, _
                            ByVal nDataYCount As Long, _
                            T As tRMC_XYSERIES _
                            ) As RMCError

Declare Function RMC_CalcAverage Lib "RMCHART.DLL" Alias "RMC_CALCAVERAGE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeriesIndex As Long, _
                            ByRef nAverage As Double, _
                            ByRef nXStart As Long, _
                            ByRef nYStart As Long, _
                            ByRef nXEnd As Long, _
                            ByRef nYEnd As Long, _
                            Optional ByVal sHighLowIndex As String _
                            ) As RMCError
                            
Declare Function RMC_CalcTrend Lib "RMCHART.DLL" Alias "RMC_CALCTREND" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeriesIndex As Long, _
                            ByRef nFirstValue As Double, _
                            ByRef nLastValue As Double, _
                            ByRef nXStart As Long, _
                            ByRef nYStart As Long, _
                            ByRef nXEnd As Long, _
                            ByRef nYEnd As Long, _
                            Optional ByVal sHighLowIndex As String _
                            ) As RMCError
                            
Declare Function RMC_COBox Lib "RMCHART.DLL" Alias "RMC_COBOX" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByVal nLeft As Long, _
                            ByVal nTop As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            Optional ByVal nStyle As COBoxStyle, _
                            Optional ByVal nBGColor As RMC_Colors, _
                            Optional ByVal nLineColor As RMC_Colors, _
                            Optional ByVal nTransparency As Long _
                            ) As RMCError

Declare Function RMC_COCircle Lib "RMCHART.DLL" Alias "RMC_COCIRCLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByVal nXCenter As Long, _
                            ByVal nYCenter As Long, _
                            ByVal nWidth As Long, _
                            Optional ByVal nStyle As COCircleStyle, _
                            Optional ByVal nBGColor As RMC_Colors, _
                            Optional ByVal nLineColor As RMC_Colors, _
                            Optional ByVal nTransparency As Long _
                            ) As RMCError
                            
Declare Function RMC_CODash Lib "RMCHART.DLL" Alias "RMC_CODASH" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nCOIndex As Long, _
                            ByVal nXStart As Long, _
                            ByVal nYStart As Long, _
                            ByVal nXEnd As Long, _
                            ByVal nYEnd As Long, _
                            Optional ByVal nStyle As COLineStyle, _
                            Optional ByVal nColor As RMC_Colors, _
                            Optional ByVal nAsSpline As Boolean, _
                            Optional ByVal nLineWidth As Long, _
                            Optional ByVal nStartCap As COAnchor, _
                            Optional ByVal nEndCap As COAnchor _
                            ) As RMCError

Declare Function RMC_CODelete Lib "RMCHART.DLL" Alias "RMC_CODELETE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long _
                            ) As RMCError

Declare Function RMC_COGetTextWH Lib "RMCHART.DLL" Alias "RMC_COGETTEXTWH" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByRef nWH As Long _
                            ) As Long

Declare Function RMC_COImage Lib "RMCHART.DLL" Alias "RMC_COIMAGE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByVal sImagePath As String, _
                            ByVal nLeft As Long, _
                            ByVal nTop As Long, _
                            Optional ByVal nWidth As Long, _
                            Optional ByVal nHeight As Long _
                            ) As RMCError

Declare Function RMC_COLine Lib "RMCHART.DLL" Alias "RMC_COLINE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByRef nXPoints As Long, _
                            ByRef nYPoints As Long, _
                            ByVal nPointsCount As Long, _
                            Optional ByVal nStyle As COLineStyle, _
                            Optional ByVal nColor As RMC_Colors, _
                            Optional ByVal nAsSpline As Boolean, _
                            Optional ByVal nLineWidth As Long, _
                            Optional ByVal nStartCap As COAnchor, _
                            Optional ByVal nEndCap As COAnchor _
                            ) As RMCError

Declare Function RMC_COPolygon Lib "RMCHART.DLL" Alias "RMC_COPOLYGON" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByRef nXPoints As Long, _
                            ByRef nYPoints As Long, _
                            ByVal nPointsCount As Long, _
                            Optional ByVal nBGColor As RMC_Colors, _
                            Optional ByVal nLineColor As RMC_Colors, _
                            Optional ByVal nAsSpline As Boolean, _
                            Optional ByVal nTransparency As Long _
                            ) As RMCError

Declare Function RMC_COSymbol Lib "RMCHART.DLL" Alias "RMC_COSYMBOL" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByVal nXCenter As Long, _
                            ByVal nYCenter As Long, _
                            ByVal nStyle As LineSeriesSymbol, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_COText Lib "RMCHART.DLL" Alias "RMC_COTEXT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByVal sText As String, _
                            ByVal nLeft As Long, _
                            ByVal nTop As Long, _
                            Optional ByVal nWidth As Long, _
                            Optional ByVal nHeight As Long, _
                            Optional ByVal nStyle As COBoxStyle, _
                            Optional ByVal nBGColor As RMC_Colors, _
                            Optional ByVal nLineColor As RMC_Colors, _
                            Optional ByVal nTransparency As Long, _
                            Optional ByVal nLineAlignment As COLineAlignment, _
                            Optional ByVal nTextColor As RMC_Colors, _
                            Optional ByVal sTextProperties As String _
                            ) As RMCError

Declare Function RMC_COVisible Lib "RMCHART.DLL" Alias "RMC_COVISIBLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nIndex As Long, _
                            ByVal nHide As Long _
                            ) As RMCError

Declare Function RMC_CreateChartOnDC Lib "RMCHART.DLL" Alias "RMC_CREATECHARTONDC" ( _
                            ByVal nParentDC As Long, _
                            ByVal nCtrlId As Long, _
                            ByVal nX As Long, _
                            ByVal nY As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            Optional ByVal nBackColor As RMC_Colors, _
                            Optional ByVal nCtrlStyle As Long, _
                            Optional ByVal nExportOnly As Boolean, _
                            Optional ByVal sBgImage As String, _
                            Optional ByVal sFontName As String, _
                            Optional ByVal nBitmapBKColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_CreateChartOnDCI Lib "RMCHART.DLL" Alias "RMC_CREATECHARTONDCI" ( _
                            ByVal nParentDC As Long, _
                            ByVal nCtrlId As Long, _
                            T As tRMC_CHART _
                            ) As RMCError

Declare Function RMC_CreateChartFromFileOnDC Lib "RMCHART.DLL" Alias "RMC_CREATECHARTFROMFILEONDC" ( _
                            ByVal nParentDC As Long, _
                            ByVal nCtrlId As Long, _
                            ByVal nX As Long, _
                            ByVal nY As Long, _
                            ByVal nExportOnly As Long, _
                            ByVal sRMCFile As String _
                             ) As RMCError

Declare Function RMC_DeleteChart Lib "RMCHART.DLL" Alias "RMC_DELETECHART" ( _
                            ByVal nCtrlId As Long _
                            ) As RMCError

Declare Function RMC_Draw Lib "RMCHART.DLL" Alias "RMC_DRAW" ( _
                            ByVal nCtrlId As Long _
                            ) As RMCError
                            
Declare Function RMC_Draw2Clipboard Lib "RMCHART.DLL" Alias "RMC_DRAW2CLIPBOARD" ( _
                            ByVal nCtrlId As Long, _
                            Optional ByVal nType As RMCFileType _
                            ) As RMCError
                 
Declare Function RMC_Draw2File Lib "RMCHART.DLL" Alias "RMC_DRAW2FILE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal sFileName As String, _
                            Optional ByVal nWidth As Long, _
                            Optional ByVal nHeight As Long, _
                            Optional ByVal nJPGQualityLevel As Long = 0 _
                            ) As RMCError

Declare Function RMC_Draw2Printer Lib "RMCHART.DLL" Alias "RMC_DRAW2PRINTER" ( _
                            ByVal nCtrlId As Long, _
                            Optional ByVal nPrinterDC As Long, _
                            Optional ByVal nLeft As Long, _
                            Optional ByVal nTop As Long, _
                            Optional ByVal nWidth As Long, _
                            Optional ByVal nHeight As Long, _
                            Optional ByVal nType As RMCFileType _
                            ) As RMCError
                            
Declare Function RMC_GetChartSizeFromFile Lib "RMCHART.DLL" Alias "RMC_GETCHARTSIZEFROMFILE" ( _
                            ByVal sRMCFile As String, _
                            ByRef nWidth As Long, _
                            ByRef nHeight As Long _
                            ) As RMCError

Declare Function RMC_GetImageSizeFromFile Lib "RMCHART.DLL" Alias "RMC_GETIMAGESIZEFROMFILE" ( _
                            ByVal sImagePath As String, _
                            ByRef nWidth As Long, _
                            ByRef nHeight As Long _
                            ) As RMCError
                            
Declare Function RMC_GetCtrlWidth Lib "RMCHART.DLL" Alias "RMC_GETCTRLWIDTH" ( _
                            ByVal nCtrlId As Long _
                            ) As RMCError

Declare Function RMC_GetCtrlHeight Lib "RMCHART.DLL" Alias "RMC_GETCTRLHEIGHT" ( _
                            ByVal nCtrlId As Long _
                            ) As RMCError
                            
Declare Function RMC_GetData Lib "RMCHART.DLL" Alias "RMC_GETDATA" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeriesIndex As Long, _
                            ByVal nDataIndex As Long, _
                            ByRef nData As Double, _
                            Optional ByVal nYData As Long _
                            ) As Long

Declare Function RMC_GetDataCount Lib "RMCHART.DLL" Alias "RMC_GETDATACOUNT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeriesIndex As Long, _
                            ByRef nDataCount As Long _
                            ) As Long

Declare Function RMC_GetDataLocation Lib "RMCHART.DLL" Alias "RMC_GETDATALOCATION" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeriesIndex As Long, _
                            ByVal nData As Double, _
                            ByRef nXYPos As Long _
                            ) As Long

Declare Function RMC_GetDataLocationXY Lib "RMCHART.DLL" Alias "RMC_GETDATALOCATIONXY" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeriesIndex As Long, _
                            ByVal nDataX As Double, _
                            ByVal nDataY As Double, _
                            ByRef nXPos As Long, _
                            ByRef nYPos As Long _
                            ) As Long


Declare Function RMC_GetGridLocation Lib "RMCHART.DLL" Alias "RMC_GETGRIDLOCATION" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByRef nLeft As Long, _
                            ByRef nTop As Long, _
                            ByRef nRight As Long, _
                            ByRef nBottom As Long _
                            ) As Long
                            
Declare Function RMC_GetHighPart Lib "RMCHART.DLL" Alias "RMC_GETHIGHPART" ( _
                           ByVal nParam As Double _
                           ) As Long
                            
Declare Function RMC_GetINFO Lib "RMCHART.DLL" Alias "RMC_GETINFO" ( _
                            ByVal nCtrlId As Long, _
                            ByRef tINFO As tRMC_INFO, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nIndex As Long _
                            ) As RMCError
            
Declare Function RMC_GetINFOXY Lib "RMCHART.DLL" Alias "RMC_GETINFOXY" ( _
                            ByVal nCtrlId As Long, _
                            ByRef tINFO As tRMC_INFO, _
                            ByVal nX As Long, _
                            ByVal nY As Long _
                            ) As RMCError

Declare Function RMC_GetLowPart Lib "RMCHART.DLL" Alias "RMC_GETLOWPART" ( _
                           ByVal nParam As Double _
                           ) As Long
                            
Declare Function RMC_GetSeriesDataRange Lib "RMCHART.DLL" Alias "RMC_GETSERIESDATARANGE" ( _
                           ByVal nCtrlId As Long, _
                           ByVal nRegion As Long, _
                           ByVal nSeries As Long, _
                           ByRef nFirst As Long, _
                           ByRef nLast As Long _
                           ) As Long

Declare Function RMC_GetVersion Lib "RMCHART.DLL" Alias "RMC_GETVERSION" ( _
                            ) As Double

Declare Function RMC_Magnifier Lib "RMCHART.DLL" Alias "RMC_MAGNIFIER" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nEnable As Long, _
                            Optional ByVal nSize As Long, _
                            Optional ByVal nColor As Long, _
                            Optional ByVal nLineColor As Long, _
                            Optional ByVal nTransparency As Long _
                            ) As RMCError
                            
Declare Function RMC_Paint Lib "RMCHART.DLL" Alias "RMC_PAINT" ( _
                            ByVal nCtrlId As Long _
                            ) As RMCError

Declare Function RMCvb_ReadDataFromFile Lib "RMCHART.DLL" Alias "RMCVB_READDATAFROMFILE" ( _
                            ByRef aData() As Double, _
                            ByRef sFileName As String, _
                            ByRef sLines As String, _
                            ByRef sFields As String, _
                            ByRef sFieldDelimiter As String, _
                            Optional ByVal nReverse As Long _
                            ) As RMCError

Declare Function RMCvb_ReadStringFromFile Lib "RMCHART.DLL" Alias "RMCVB_READSTRINGFROMFILE" ( _
                            ByRef aValue() As String, _
                            ByRef sFileName As String, _
                            ByRef sLines As String, _
                            ByRef sFields As String, _
                            ByRef sFieldDelimiter As String, _
                            Optional ByVal nReverse As Long _
                            ) As RMCError

Declare Function RMC_Reset Lib "RMCHART.DLL" Alias "RMC_RESET" ( _
                            ByVal nCtrlId As Long _
                             ) As RMCError

Declare Function RMC_RND Lib "RMCHART.DLL" ( _
                            ByVal n1 As Long, _
                            ByVal n2 As Long _
                            ) As Long

Declare Function RMCvb_Split2Double Lib "RMCHART.DLL" Alias "RMCVB_SPLIT2DOUBLE" ( _
                            ByRef sData As String, _
                            ByRef aData() As Double _
                            ) As RMCError

Declare Function RMCvb_Split2Long Lib "RMCHART.DLL" Alias "RMCVB_SPLIT2LONG" ( _
                            ByRef sData As String, _
                            ByRef aData() As Long _
                            ) As RMCError

Declare Function RMC_SetCaptionText Lib "RMCHART.DLL" Alias "RMC_SETCAPTIONTEXT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sText As String _
                            ) As RMCError
                            
Declare Function RMC_SetCaptionBGColor Lib "RMCHART.DLL" Alias "RMC_SETCAPTIONBGCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetCaptionTextColor Lib "RMCHART.DLL" Alias "RMC_SETCAPTIONTEXTCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetCaptionFontSize Lib "RMCHART.DLL" Alias "RMC_SETCAPTIONFONTSIZE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFontSize As Long _
                            ) As RMCError

Declare Function RMC_SetCaptionFontBold Lib "RMCHART.DLL" Alias "RMC_SETCAPTIONFONTBOLD" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFontBold As Long _
                            ) As RMCError
                 
Declare Function RMC_SetCtrlBGColor Lib "RMCHART.DLL" Alias "RMC_SETCTRLBGCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetCtrlBGImage Lib "RMCHART.DLL" Alias "RMC_SETCTRLBGIMAGE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal sBgImage As String _
                            ) As RMCError

Declare Function RMC_SetCtrlFont Lib "RMCHART.DLL" Alias "RMC_SETCTRLFONT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal sFontName As String _
                            ) As RMCError

Declare Function RMC_SetCtrlSize Lib "RMCHART.DLL" Alias "RMC_SETCTRLSIZE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            Optional ByVal nRelative As Long, _
                            Optional ByVal nRecalcMode As Long _
                            ) As RMCError

Declare Function RMC_SetCtrlStyle Lib "RMCHART.DLL" Alias "RMC_SETCTRLSTYLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nStyle As CtrlStyle _
                            ) As RMCError

Declare Function RMC_SetCustomToolTipText Lib "RMCHART.DLL" Alias "RMC_SETCUSTOMTOOLTIPTEXT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nDataIndex As Long, _
                            ByVal sText As String _
                            ) As Long

Declare Function RMC_SetDAXAlignment Lib "RMCHART.DLL" Alias "RMC_SETDAXALIGNMENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nAlignment As DAxisAlignment _
                            ) As RMCError

Declare Function RMC_SetDAXDecimalDigits Lib "RMCHART.DLL" Alias "RMC_SETDAXDECIMALDIGITS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nDecimalDigits As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetDAXFontSize Lib "RMCHART.DLL" Alias "RMC_SETDAXFONTSIZE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFontSize As Long _
                            ) As RMCError

Declare Function RMC_SetDAXLineColor Lib "RMCHART.DLL" Alias "RMC_SETDAXLINECOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetDAXLineStyle Lib "RMCHART.DLL" Alias "RMC_SETDAXLINESTYLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nStyle As AxisLineStyle _
                            ) As RMCError
                  
Declare Function RMC_SetDAXMaxValue Lib "RMCHART.DLL" Alias "RMC_SETDAXMAXVALUE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nMaxValue As Double, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetDAXMinValue Lib "RMCHART.DLL" Alias "RMC_SETDAXMINVALUE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nMinValue As Double, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetDAXText Lib "RMCHART.DLL" Alias "RMC_SETDAXTEXT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sText As String, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetDAXTextColor Lib "RMCHART.DLL" Alias "RMC_SETDAXTEXTCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError
                  
Declare Function RMC_SetDAXTickcount Lib "RMCHART.DLL" Alias "RMC_SETDAXTICKCOUNT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nTickCount As Long _
                            ) As RMCError
                  
Declare Function RMC_SetDAXUnit Lib "RMCHART.DLL" Alias "RMC_SETDAXUNIT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sUnit As String, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetGridBGColor Lib "RMCHART.DLL" Alias "RMC_SETGRIDBGCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetGridBiColor Lib "RMCHART.DLL" Alias "RMC_SETGRIDBICOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nBiColor As Long _
                            ) As RMCError

Declare Function RMC_SetGridGradient Lib "RMCHART.DLL" Alias "RMC_SETGRIDGRADIENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nHasGradient As Long _
                            ) As RMCError

Declare Function RMC_SetGridMargin Lib "RMCHART.DLL" Alias "RMC_SETGRIDMARGIN" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nLeft As Long, _
                            ByVal nTop As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long _
                            ) As RMCError
                            
Declare Function RMC_SetHelpingGrid Lib "RMCHART.DLL" Alias "RMC_SETHELPINGGRID" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nSize As Long, _
                            Optional ByVal nGridColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetLAXAlignment Lib "RMCHART.DLL" Alias "RMC_SETLAXALIGNMENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nAlignment As LAxisAlignment _
                            ) As RMCError

Declare Function RMC_SetLAXCount Lib "RMCHART.DLL" Alias "RMC_SETLAXCOUNT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nLabelAxisCount As Long _
                            ) As RMCError

Declare Function RMC_SetLAXFontSize Lib "RMCHART.DLL" Alias "RMC_SETLAXFONTSIZE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFontSize As Long _
                            ) As RMCError

Declare Function RMC_SetLAXLabelAlignment Lib "RMCHART.DLL" Alias "RMC_SETLAXLABELALIGNMENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nAlignment As LabelTextAlignment _
                            ) As RMCError

Declare Function RMC_SetLAXLabels Lib "RMCHART.DLL" Alias "RMC_SETLAXLABELS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sLabels As String _
                            ) As RMCError

Declare Function RMC_SetLAXLabelsFile Lib "RMCHART.DLL" Alias "RMC_SETLAXLABELSFILE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sFileName As String, _
                            Optional ByVal sLines As String, _
                            Optional ByVal sFields As String, _
                            Optional ByVal sFieldDelimiter As String _
                            ) As RMCError

Declare Function RMC_SetLAXLabelsRange Lib "RMCHART.DLL" Alias "RMC_SETLAXLABELSRANGE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFirst As Long, _
                            ByVal nLast As Long _
                            ) As RMCError

Declare Function RMC_SetLAXLineColor Lib "RMCHART.DLL" Alias "RMC_SETLAXLINECOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetLAXLineStyle Lib "RMCHART.DLL" Alias "RMC_SETLAXLINESTYLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nStyle As AxisLineStyle _
                            ) As RMCError

Declare Function RMC_SetLAXText Lib "RMCHART.DLL" Alias "RMC_SETLAXTEXT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sText As String _
                            ) As RMCError
                  
Declare Function RMC_SetLAXTextColor Lib "RMCHART.DLL" Alias "RMC_SETLAXTEXTCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetLAXTickCount Lib "RMCHART.DLL" Alias "RMC_SETLAXTICKCOUNT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nTickCount As Long _
                            ) As RMCError

Declare Function RMC_SetLegendAlignment Lib "RMCHART.DLL" Alias "RMC_SETLEGENDALIGNMENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nAlignment As LegendAlignment _
                            ) As RMCError

Declare Function RMC_SetLegendBGColor Lib "RMCHART.DLL" Alias "RMC_SETLEGENDBGCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetLegendFontBold Lib "RMCHART.DLL" Alias "RMC_SETLEGENDFONTBOLD" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFontBold As Long _
                            ) As RMCError

Declare Function RMC_SetLegendFontSize Lib "RMCHART.DLL" Alias "RMC_SETLEGENDFONTSIZE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFontSize As Long _
                            ) As RMCError
                  
Declare Function RMC_SetLegendHide Lib "RMCHART.DLL" Alias "RMC_SETLEGENDHIDE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nHide As Long _
                            ) As Long

Declare Function RMC_SetLegendStyle Lib "RMCHART.DLL" Alias "RMC_SETLEGENDSTYLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nStyle As LegendStyle _
                            ) As RMCError

Declare Function RMC_SetLegendText Lib "RMCHART.DLL" Alias "RMC_SETLEGENDTEXT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sLegendText As String _
                            ) As RMCError

Declare Function RMC_SetLegendTextColor Lib "RMCHART.DLL" Alias "RMC_SETLEGENDTEXTCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors _
                            ) As RMCError

Declare Function RMC_SetMouseClick Lib "RMCHART.DLL" Alias "RMC_SETMOUSECLICK" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nButton As Long, _
                            ByVal nX As Long, _
                            ByVal nY As Long, _
                            ByRef tINFO As tRMC_INFO _
                            ) As RMCError

Declare Function RMC_SetRegionFooter Lib "RMCHART.DLL" Alias "RMC_SETREGIONFOOTER" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sFooter As String _
                            ) As RMCError

Declare Function RMC_SetRegionMargin Lib "RMCHART.DLL" Alias "RMC_SETREGIONMARGIN" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nLeft As Long, _
                            ByVal nTop As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long _
                            ) As RMCError
                  
Declare Function RMC_SetRegionBorder Lib "RMCHART.DLL" Alias "RMC_SETREGIONBORDER" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nShowBorder As Long _
                            ) As RMCError

Declare Function RMC_SetRMCFile Lib "RMCHART.DLL" Alias "RMC_SETRMCFILE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal sRMCFile As String _
                            ) As RMCError

Declare Function RMC_SetSeriesColor Lib "RMCHART.DLL" Alias "RMC_SETSERIESCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nColor As RMC_Colors, _
                            Optional ByVal nIndex As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesExplodeMode Lib "RMCHART.DLL" Alias "RMC_SETSERIESEXPLODEMODE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nExplodeMode As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesStartAngle Lib "RMCHART.DLL" Alias "RMC_SETSERIESSTARTANGLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nStartAngle As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesData Lib "RMCHART.DLL" Alias "RMC_SETSERIESDATA" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByRef nData As Double, _
                            ByVal nDataCount As Long, _
                            Optional ByVal nYData As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesDataFile Lib "RMCHART.DLL" Alias "RMC_SETSERIESDATAFILE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal sFileName As String, _
                            Optional ByVal sLines As String, _
                            Optional ByVal sFields As String, _
                            Optional ByVal sFieldDelimiter As String, _
                            Optional ByVal nYData As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesDataRange Lib "RMCHART.DLL" Alias "RMC_SETSERIESDATARANGE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nFirst As Long, _
                            ByVal nLast As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesSingleData Lib "RMCHART.DLL" Alias "RMC_SETSERIESSINGLEDATA" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nData As Double, _
                            ByVal nDataIndex As Long, _
                            Optional ByVal nYData As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesDataAxis Lib "RMCHART.DLL" Alias "RMC_SETSERIESDATAAXIS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nWhichAxis As Long _
                            ) As RMCError
                 
Declare Function RMC_SetSeriesHatchMode Lib "RMCHART.DLL" Alias "RMC_SETSERIESHATCHMODE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nHatchMode As Hatchmodes _
                            ) As RMCError
                  
Declare Function RMC_SetSeriesHide Lib "RMCHART.DLL" Alias "RMC_SETSERIESHIDE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nHide As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesHighLowColor Lib "RMCHART.DLL" Alias "RMC_SETSERIESHIGHLOWCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nColorLow As RMC_Colors, _
                            ByVal nColorHigh As RMC_Colors _
                            ) As RMCError
                  
Declare Function RMC_SetSeriesLinestyle Lib "RMCHART.DLL" Alias "RMC_SETSERIESLINESTYLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nLinestyle As LineSeriesLineStyle _
                            ) As RMCError
                  
Declare Function RMC_SetSeriesLucent Lib "RMCHART.DLL" Alias "RMC_SETSERIESLUCENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nLucent As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesPPColumn Lib "RMCHART.DLL" Alias "RMC_SETSERIESPPCOLUMN" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nPointsPerColumn As Long _
                            ) As Long
                 
Declare Function RMC_SetSeriesPPColumnArray Lib "RMCHART.DLL" Alias "RMC_SETSERIESPPCOLUMNARRAY" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByRef nFirstPPCValue As Long, _
                            ByVal nPPCValuesCount As Long _
                            ) As Long

Declare Function RMC_SetSeriesVertical Lib "RMCHART.DLL" Alias "RMC_SETSERIESVERTICAL" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nVertical As Long _
                            ) As RMCError
                  
Declare Function RMC_SetSeriesStyle Lib "RMCHART.DLL" Alias "RMC_SETSERIESSTYLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nStyle As Long _
                            ) As RMCError

Declare Function RMC_SetSeriesSymbol Lib "RMCHART.DLL" Alias "RMC_SETSERIESSYMBOL" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nSymbol As LineSeriesSymbol _
                            ) As RMCError

Declare Function RMC_SetSeriesValuelabel Lib "RMCHART.DLL" Alias "RMC_SETSERIESVALUELABEL" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nValuelabel As ValueLabels _
                            ) As RMCError

Declare Function RMC_SetSeriesXAxis Lib "RMCHART.DLL" Alias "RMC_SETSERIESXAXIS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nWhichXAxis As Long _
                            ) As RMCError
                     
Declare Function RMC_SetSeriesYAxis Lib "RMCHART.DLL" Alias "RMC_SETSERIESYAXIS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nSeries As Long, _
                            ByVal nWhichYAxis As Long _
                            ) As RMCError

Declare Function RMC_SetSingleBarColors Lib "RMCHART.DLL" Alias "RMC_SETSINGLEBARCOLORS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByRef nColor As RMC_Colors, _
                            ByVal nColorCount As Long _
                            ) As RMCError

Declare Function RMC_SetToolTipWidth Lib "RMCHART.DLL" Alias "RMC_SETTOOLTIPWIDTH" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nWidth As Long _
                            ) As Long

Declare Function RMC_SetWatermark Lib "RMCHART.DLL" Alias "RMC_SETWATERMARK" ( _
                            ByVal sWaterMark As String, _
                            Optional ByVal nWMColor As RMC_Colors, _
                            Optional ByVal nWMLucentValue As Long, _
                            Optional ByVal nAlignment As Long, _
                            Optional ByVal nFontSize As Long _
                            ) As RMCError
Declare Function RMC_SetXAXAlignment Lib "RMCHART.DLL" Alias "RMC_SETXAXALIGNMENT" ( _
                          ByVal nCtrlId As Long, _
                          ByVal nRegion As Long, _
                          ByVal nAlignment As XAxisAlignment, _
                          Optional ByVal nAxisIndex As Long _
                          ) As RMCError

Declare Function RMC_SetYAXAlignment Lib "RMCHART.DLL" Alias "RMC_SETYAXALIGNMENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nAlignment As YAxisAlignment, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXDecimalDigits Lib "RMCHART.DLL" Alias "RMC_SETXAXDECIMALDIGITS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nDecimalDigits As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetYAXDecimalDigits Lib "RMCHART.DLL" Alias "RMC_SETYAXDECIMALDIGITS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nDecimalDigits As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetXAXFontSize Lib "RMCHART.DLL" Alias "RMC_SETXAXFONTSIZE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFontSize As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetYAXFontSize Lib "RMCHART.DLL" Alias "RMC_SETYAXFONTSIZE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nFontSize As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXLabels Lib "RMCHART.DLL" Alias "RMC_SETXAXLABELS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sLabels As String, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetYAXLabels Lib "RMCHART.DLL" Alias "RMC_SETYAXLABELS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sLabels As String, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXLabelAlignment Lib "RMCHART.DLL" Alias "RMC_SETXAXLABELALIGNMENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nLabelAlignment As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetYAXLabelAlignment Lib "RMCHART.DLL" Alias "RMC_SETYAXLABELALIGNMENT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nLabelAlignment As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXLineColor Lib "RMCHART.DLL" Alias "RMC_SETXAXLINECOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetYAXLineColor Lib "RMCHART.DLL" Alias "RMC_SETYAXLINECOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                 
Declare Function RMC_SetXAXLineStyle Lib "RMCHART.DLL" Alias "RMC_SETXAXLINESTYLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nStyle As AxisLineStyle, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetYAXLineStyle Lib "RMCHART.DLL" Alias "RMC_SETYAXLINESTYLE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nStyle As AxisLineStyle, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXMaxValue Lib "RMCHART.DLL" Alias "RMC_SETXAXMAXVALUE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nMaxValue As Double, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetYAXMaxValue Lib "RMCHART.DLL" Alias "RMC_SETYAXMAXVALUE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nMaxValue As Double, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXMinValue Lib "RMCHART.DLL" Alias "RMC_SETXAXMINVALUE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nMinValue As Double, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetYAXMinValue Lib "RMCHART.DLL" Alias "RMC_SETYAXMINVALUE" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nMinValue As Double, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXText Lib "RMCHART.DLL" Alias "RMC_SETXAXTEXT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sText As String, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetYAXText Lib "RMCHART.DLL" Alias "RMC_SETYAXTEXT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sText As String, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXTextColor Lib "RMCHART.DLL" Alias "RMC_SETXAXTEXTCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetYAXTextColor Lib "RMCHART.DLL" Alias "RMC_SETYAXTEXTCOLOR" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nColor As RMC_Colors, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetXAXTickcount Lib "RMCHART.DLL" Alias "RMC_SETXAXTICKCOUNT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nTickCount As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetYAXTickcount Lib "RMCHART.DLL" Alias "RMC_SETYAXTICKCOUNT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal nTickCount As Long, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError
                  
Declare Function RMC_SetXAXUnit Lib "RMCHART.DLL" Alias "RMC_SETXAXUNIT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sUnit As String, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_SetYAXUnit Lib "RMCHART.DLL" Alias "RMC_SETYAXUNIT" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nRegion As Long, _
                            ByVal sUnit As String, _
                            Optional ByVal nAxisIndex As Long _
                            ) As RMCError

Declare Function RMC_ShowToolTips Lib "RMCHART.DLL" Alias "RMC_SHOWTOOLTIPS" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nX As Long, _
                            ByVal nY As Long _
                            ) As RMCError

Declare Function RMCvb_WriteRMCFile Lib "RMCHART.DLL" Alias "RMCVB_WRITERMCFILE" ( _
                            ByVal nCtrlId As Long, _
                            ByRef sFileName As String _
                            ) As RMCError

Declare Function RMC_Zoom Lib "RMCHART.DLL" Alias "RMC_ZOOM" ( _
                            ByVal nCtrlId As Long, _
                            ByVal nMode As Long, _
                            Optional ByVal nColor As Long, _
                            Optional ByVal nLineColor As Long, _
                            Optional ByVal nTransparency As Long _
                            ) As RMCError


Public Const RMC_USERWM = ""                      ' Your watermark
Public Const RMC_USERWMCOLOR = Black              ' Color for the watermark
Public Const RMC_USERWMLUCENT = 30                ' Lucent factor between 1(=not visible) and 255(=opaque)
Public Const RMC_USERWMALIGN = RMC_TEXTCENTER     ' Alignment for the watermark
Public Const RMC_USERFONTSIZE = 0                 ' Fontsize; if 0: maximal size is used

' ******************************************************************
