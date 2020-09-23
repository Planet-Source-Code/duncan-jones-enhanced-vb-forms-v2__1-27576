VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiDeviceContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'\\ API DEclarations ---------------------------------------------------
Private Declare Function GetDeviceCapsApi Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long

'\\ Drawing related...
Private Declare Function DrawIconApi Lib "user32" Alias "DrawIcon" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long


Public Enum enDeviceCapsIndexes
    cDRIVERVERSION = 0      '  Device driver version
    cTECHNOLOGY = 2         '  Device classification
    cHORZSIZE = 4           '  Horizontal size in millimeters
    cVERTSIZE = 6           '  Vertical size in millimeters
    cHORZRES = 8            '  Horizontal width in pixels
    cVERTRES = 10           '  Vertical width in pixels
    cLOGPIXELSX = 88        '  Logical pixels/inch in X
    cLOGPIXELSY = 90        '  Logical pixels/inch in Y
    cBITSPIXEL = 12         '  Number of bits per pixel
    cPLANES = 14            '  Number of planes
    cNUMBRUSHES = 16        '  Number of brushes the device has
    cNUMCOLORS = 24         '  Number of colors the device supports
    cNUMFONTS = 22          '  Number of fonts the device has
    cNUMMARKERS = 20        '  Number of markers the device has
    cNUMPENS = 18           '  Number of pens the device has
    cASPECTX = 40           '  Length of the X leg
    cASPECTXY = 44          '  Length of the hypotenuse
    cASPECTY = 42           '  Length of the Y leg
    cPDEVICESIZE = 26       '  Size required for device descriptor
    cCLIPCAPS = 36          '  Clipping capabilities
    cSIZEPALETTE = 104      '  Number of entries in physical palette
    cNUMRESERVED = 106      '  Number of reserved entries in palette
    cCOLORRES = 108         '  Actual color resolution
    cPHYSICALOFFSETX = 112 '  Physical Printable Area x margin
    cPHYSICALOFFSETY = 113 '  Physical Printable Area y margin
    cPHYSICALHEIGHT = 111 '  Physical Height in device units
    cPHYSICALWIDTH = 110 '  Physical Width in device units
    cSCALINGFACTORX = 114 '  Scaling factor x
    cSCALINGFACTORY = 115 '  Scaling factor y
    cRASTERCAPS = 38        '  Bitblt capabilities
    cCURVECAPS = 28         '  Curve capabilities
    cLINECAPS = 30          '  Line capabilities
    cPOLYGONALCAPS = 32     '  Polygonal capabilities
    cTEXTCAPS = 34          '  Text capabilities
End Enum

Public Enum enDisplayTypes
    DT_PLOTTER = 0             '  Vector plotter
    DT_RASCAMERA = 3           '  Raster camera
    DT_RASDISPLAY = 1          '  Raster display
    DT_RASPRINTER = 2          '  Raster printer
    DT_CHARSTREAM = 4          '  Character-stream, PLP
    DT_METAFILE = 5            '  Metafile, VDM
    DT_DISPFILE = 6            '  Display-file
End Enum

Public Enum enDeviceMesaurementScale
    DMS_Millimeters = 1
    DMS_Pixels = 2
End Enum

Public Enum enClipCapabilities
    CP_NONE = 0                '  No clipping of output
    CP_RECTANGLE = 1           '  Output clipped to rects
    CP_REGION = 2              '
End Enum

Public Enum enRasterCapabilities
    RC_BANDING = 2                 '  Device requires banding support
    RC_BIGFONT = &H400                 '  supports >64K fonts
    RC_BITBLT = 1                  '  Can do standard BLT.
    RC_BITMAP64 = 8                '  Device can support >64K bitmap
    RC_DEVBITS = &H8000
    RC_DI_BITMAP = &H80                '  supports DIB to memory
    RC_DIBTODEV = &H200                '  supports DIBitsToDevice
    RC_FLOODFILL = &H1000              '  supports FloodFill
    RC_NONE = 0
    RC_PALETTE = &H100                 '  supports a palette
    RC_SAVEBITMAP = &H40
    RC_SCALING = 4                 '  Device requires scaling support
    RC_STRETCHBLT = &H800              '  supports StretchBlt
    RC_STRETCHDIB = &H2000             '  supports StretchDIBits
End Enum

Public Enum enCurvecapabilities
    CC_CIRCLES = 1             '  Can do circles
    CC_CHORD = 4               '  Can do chord arcs
    CC_ELLIPSES = 8            '  Can do ellipese
    CC_PIE = 2                 '  Can do pie wedges
    CC_NONE = 0                '  Curves not supported
    CC_WIDE = 16               '  Can do wide lines
    CC_STYLED = 32             '  Can do styled lines
    CC_WIDESTYLED = 64         '  Can do wide styled lines
    CC_INTERIORS = 128 '  Can do interiors
End Enum

Public Enum enLineCapabilities
    LC_NONE = 0                '  Lines not supported
    LC_POLYLINE = 2            '  Can do polylines
    LC_MARKER = 4              '  Can do markers
    LC_POLYMARKER = 8          '  Can do polymarkers
    LC_WIDE = 16               '  Can do wide lines
    LC_STYLED = 32             '  Can do styled lines
    LC_WIDESTYLED = 64         '  Can do wide styled lines
    LC_INTERIORS = 128 '  Can do interiors
End Enum

Public Enum enPolygonCapabilities
    PC_NONE = 0                '  Polygonals not supported
    PC_POLYGON = 1             '  Can do polygons
    PC_RECTANGLE = 2           '  Can do rectangles
    PC_TRAPEZOID = 4           '  Can do trapezoids
    PC_SCANLINE = 8            '  Can do scanlines
    PC_WIDE = 16               '  Can do wide borders
    PC_STYLED = 32             '  Can do styled borders
    PC_WIDESTYLED = 64         '  Can do wide styled borders
    PC_INTERIORS = 128         '  Can do interiors
End Enum

Public Enum enTextCapabilities
    TC_OP_CHARACTER = &H1              '  Can do OutputPrecision   CHARACTER
    TC_OP_STROKE = &H2                 '  Can do OutputPrecision   STROKE
    TC_CP_STROKE = &H4                 '  Can do ClipPrecision     STROKE
    TC_CR_90 = &H8                     '  Can do CharRotAbility    90
    TC_CR_ANY = &H10                   '  Can do CharRotAbility    ANY
    TC_SF_X_YINDEP = &H20              '  Can do ScaleFreedom      X_YINDEPENDENT
    TC_SA_DOUBLE = &H40                '  Can do ScaleAbility      DOUBLE
    TC_SA_INTEGER = &H80               '  Can do ScaleAbility      INTEGER
    TC_SA_CONTIN = &H100               '  Can do ScaleAbility      CONTINUOUS
    TC_IA_ABLE = &H400                 '  Can do ItalisizeAbility  ABLE
    TC_UA_ABLE = &H800                 '  Can do UnderlineAbility  ABLE
    TC_RA_ABLE = &H2000                '  Can do RasterFontAble    ABLE
    TC_SCROLLBLT = &H10000             '  do text scroll with blt
    TC_VA_ABLE = &H4000                '  Can do VectorFontAble    ABLE
End Enum

Public Enum enMappingModes
    MM_ANISOTROPIC = 8
    MM_HIENGLISH = 5
    MM_HIMETRIC = 3
    MM_ISOTROPIC = 7
    MM_LOENGLISH = 4
    MM_LOMETRIC = 2
    MM_TEXT = 1
    MM_TWIPS = 6
End Enum


'\\ DrawEdge Related
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Public Enum enEdgeBorderStyles
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum

Public Enum enBorderFlags
    BF_LEFT = &H1
    BF_TOP = &H2
    BF_RIGHT = &H4
    BF_BOTTOM = &H8
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_DIAGONAL = &H10
'// For diagonal lines, the BF_RECT flags specify the end point of the
'// vector bounded by the rectangle parameter.
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_MIDDLE = &H800     '// Fill in the middle
    BF_SOFT = &H1000      '// For softer buttons
    BF_ADJUST = &H2000    '// Calculate the space left over
    BF_FLAT = &H4000      '// For flat rather than 3D borders
    BF_MONO = &H8000      '// For monochrome borders
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawEdgeApi Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long

'\\ DrawFrameControl related
Private Enum enDrawFrameControlTypes
    DFC_CAPTION = 1
    DFC_MENU = 2
    DFC_SCROLL = 3
    DFC_BUTTON = 4
    DFC_POPUPMENU = 5
End Enum

Public Enum enDrawFrameCaption
    DFCS_CAPTIONCLOSE = &H0
    DFCS_CAPTIONMIN = &H1
    DFCS_CAPTIONMAX = &H2
    DFCS_CAPTIONRESTORE = &H3
    DFCS_CAPTIONHELP = &H4
End Enum

Public Enum enDrawFrameMenu
    DFCS_MENUARROW = &H0
    DFCS_MENUCHECK = &H1
    DFCS_MENUBULLET = &H2
    DFCS_MENUARROWRIGHT = &H4
End Enum

Public Enum enDrawFrameScroll
    DFCS_SCROLLUP = &H0
    DFCS_SCROLLDOWN = &H1
    DFCS_SCROLLLEFT = &H2
    DFCS_SCROLLRIGHT = &H3
    DFCS_SCROLLCOMBOBOX = &H5
    DFCS_SCROLLSIZEGRIP = &H8
    DFCS_SCROLLSIZEGRIPRIGHT = &H10
End Enum

Public Enum enDrawFrameButton
    DFCS_BUTTONCHECK = &H0
    DFCS_BUTTONRADIOIMAGE = &H1
    DFCS_BUTTONRADIOMASK = &H2
    DFCS_BUTTONRADIO = &H4
    DFCS_BUTTON3STATE = &H8
    DFCS_BUTTONPUSH = &H10
End Enum
    
Public Enum enDrawFrameButtonState
    DFCS_NORMAL = &H0 '\\ DEJ Added for completeness...
    DFCS_INACTIVE = &H100
    DFCS_PUSHED = &H200
    DFCS_CHECKED = &H400
End Enum

Public Enum enDrawFrameMenuPopupMenu
    DFCS_TRANSPARENT = &H800
    DFCS_HOT = &H1000
    DFCS_ADJUSTRECT = &H2000
    DFCS_FLAT = &H4000
    DFCS_MONO = &H8000
End Enum

Private Declare Function DrawFrameControlApi Lib "user32" Alias "DrawFrameControl" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

'\\ The graphics mode...
Private Declare Function SetGraphicsModeApi Lib "gdi32" Alias "SetGraphicsMode" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function GetGraphicsModeApi Lib "gdi32" Alias "GetGraphicsMode" (ByVal hdc As Long) As Long
Public Enum enGraphicMode
    GM_COMPATIBLE = 1
    GM_ADVANCED = 2
End Enum

'\\ Member variables...
Private mHDC As Long

Public Property Get BitsPerPixel() As Long

    BitsPerPixel = GetDeviceCaps(cBITSPIXEL)
    
End Property


Public Property Get ClipingCapabilities() As enClipCapabilities

    ClipingCapabilities = GetDeviceCaps(cCLIPCAPS)

End Property


Public Property Get ColourAdjustment() As ApiColourAdjustment

Dim mCol As ApiColourAdjustment

Set mCol = New ApiColourAdjustment
Set mCol.ParentDC = Me

Set ColourAdjustment = mCol

End Property

Public Property Get ColourPlanes() As Long

    ColourPlanes = GetDeviceCaps(cPLANES)
    
End Property

Public Property Get ColourResolution() As Long

    ColourResolution = GetDeviceCaps(cCOLORRES)

End Property

Public Function CreateCompatibleBitmap(ByVal bmWidth As Long, ByVal bmHeight As Long) As ApiBitmap

Dim bitmapThis As ApiBitmap

Set bitmapThis = New ApiBitmap

bitmapThis.CreateCompatibleBitmap mHDC, bmWidth, bmHeight

Set CreateCompatibleBitmap = bitmapThis

End Function

Public Property Get CurveCapability(ByVal Capability As enCurvecapabilities) As Boolean

Dim lRet As Long

lRet = GetDeviceCaps(cCURVECAPS)
CurveCapability = (lRet And Capability)

End Property

Public Property Get DeviceType() As enDisplayTypes

DeviceType = GetDeviceCaps(cTECHNOLOGY)

End Property


Public Sub DrawFrameButton(ByVal RectIn As APIRect, ByVal ButtonStyle As enDrawFrameButton, ByVal ButtonState As enDrawFrameButtonState)

Dim lStyle As Long

lStyle = (ButtonStyle Or ButtonState)

Call DrawFrameControl(RectIn, DFC_BUTTON, lStyle)

End Sub

'\\ --[DrawFrameCaption]------------------------------------------------------------------
'\\ Draws a caption rectangle in the given RECT
'\\ ---------------------------------------------------------------------------------
Public Sub DrawFrameCaption(ByVal RectIn As APIRect, ByVal CaptionStyle As enDrawFrameCaption)

Call DrawFrameControl(RectIn, DFC_CAPTION, CaptionStyle)

End Sub


'\\ --[DrawEdge]----------------------------------------------------------
'\\ Draws the edge of the rectangle in RectIn as per the parameters passed
'\\ in.
'\\ Parameters:
'\\     Edge - The type of edge to draw (sunken,raised etc)
'\\     Flags - A combination of enBorderFlags to set which edges to draw
'\\ ----------------------------------------------------------------------
Public Sub DrawEdge(ByVal RectIn As APIRect, ByVal Edge As enEdgeBorderStyles, ByVal Flags As Long)

Dim lRet As Long
Dim rcThis As RECT

With rcThis
    .Bottom = RectIn.Bottom
    .Left = RectIn.Left
    .Right = RectIn.Right
    .Top = RectIn.Top
End With

lRet = DrawEdgeApi(mHDC, rcThis, Edge, Flags)
If (Err.LastDllError <> 0) Or (lRet = 0) Then
    '\\ An error occured
    Call ReportError(Err.LastDllError, "ApiDeviceContext:DrawEdge", APIDispenser.LastSystemError)
End If

End Sub

Private Sub DrawFrameControl(ByVal RectIn As APIRect, ByVal FrameType As enDrawFrameControlTypes, ByVal style As Long)

Dim rcThis As RECT
Dim lRet  As Long

With rcThis
    .Bottom = RectIn.Bottom
    .Left = RectIn.Left
    .Right = RectIn.Right
    .Top = RectIn.Top
End With

lRet = DrawFrameControlApi(mHDC, rcThis, FrameType, style)
If (Err.LastDllError > 0) Or (lRet = 0) Then
    Call ReportError(Err.LastDllError, "ApiDeviceContext:DrawFrameControl", APIDispenser.LastSystemError)
End If

End Sub

Public Sub DrawFrameMenu(ByVal RectIn As APIRect, ByVal MenuStyle As enDrawFrameMenu)

Call DrawFrameControl(RectIn, DFC_MENU, MenuStyle)

End Sub

Public Sub DrawFramePopup(ByVal RectIn As APIRect, ByVal style As enDrawFrameMenuPopupMenu)

Call DrawFrameControl(RectIn, DFC_POPUPMENU, style)

End Sub

Public Sub DrawFrameScroll(ByVal RectIn As APIRect, ByVal ScrollStyle As enDrawFrameScroll)

Call DrawFrameControl(RectIn, DFC_SCROLL, ScrollStyle)

End Sub

Public Sub DrawIcon(ByVal xPos As Long, ByVal yPos As Long, ByVal Icon As ApiIcon)

Dim lRet As Long

lRet = DrawIconApi(mHDC, xPos, yPos, Icon.hIcon)
If (Err.LastDllError > 0) Or (lRet = 0) Then
    Call ReportError(Err.LastDllError, "ApiDeviceContext:DrawIcon", APIDispenser.LastSystemError)
End If

End Sub

Public Property Get DriverVersion() As Long

    DriverVersion = GetDeviceCaps(cDRIVERVERSION)
    
End Property


Private Function GetDeviceCaps(ByVal dcIndex As enDeviceCapsIndexes) As Long

GetDeviceCaps = GetDeviceCapsApi(mHDC, dcIndex)

End Function


Public Property Let GraphicsMode(ByVal newMode As enGraphicMode)

Dim lRet As Long

lRet = SetGraphicsModeApi(mHDC, newMode)
If Err.LastDllError > 0 Or lRet = 0 Then
   Call ReportError(Err.LastDllError, "ApiDeviceContext:GraphicsMode", APIDispenser.LastSystemError)
End If

End Property

Public Property Get GraphicsMode() As enGraphicMode

Dim lRet As Long

lRet = GetGraphicsModeApi(mHDC)
GraphicsMode = lRet

End Property

Public Property Get hdc() As Long

    hdc = mHDC
    
End Property

Public Property Let hdc(ByVal newHDC As Long)

    mHDC = newHDC
    
End Property
Public Property Get Height(ByVal MeasurementScale As enDeviceMesaurementScale) As Long

If MeasurementScale = DMS_Millimeters Then
    Height = GetDeviceCaps(cVERTSIZE)
Else
    Height = GetDeviceCaps(cVERTRES)
End If

End Property

Public Property Get LineCapability(ByVal Capability As enLineCapabilities) As Boolean

Dim lRet As Long

lRet = GetDeviceCaps(cLINECAPS)
LineCapability = (lRet And Capability)

End Property

Public Property Let MappingMode(ByVal newMode As enMappingModes)

Static OldMappingMode As enMappingModes

If newMode <> OldMappingMode Then
    OldMappingMode = SetMapMode(mHDC, newMode)
End If

End Property

Public Property Get MappingMode() As enMappingModes

MappingMode = GetMapMode(mHDC)

End Property


Public Property Get PixelAspectDiagonal() As Long

    PixelAspectDiagonal = GetDeviceCaps(cASPECTXY)

End Property

Public Property Get PixelAspectHeight() As Long

    PixelAspectHeight = GetDeviceCaps(cASPECTY)
    
End Property

Public Property Get PixelAspectWidth() As Long

    PixelAspectWidth = GetDeviceCaps(cASPECTX)
    
End Property

Public Property Get PixelsPerInch_Horizontal() As Long

PixelsPerInch_Horizontal = GetDeviceCaps(cLOGPIXELSX)

End Property
Public Property Get PixelsPerInch_Vertical() As Long

PixelsPerInch_Vertical = GetDeviceCaps(cLOGPIXELSY)

End Property



Public Property Get PolygonCapability(ByVal Capability As enPolygonCapabilities) As Boolean

Dim lRet As Long

lRet = GetDeviceCaps(cPOLYGONALCAPS)
PolygonCapability = (lRet And Capability)

End Property

Public Property Get RasterCapability(ByVal Capability As enRasterCapabilities) As Boolean

Dim lRet As Long

lRet = GetDeviceCaps(cRASTERCAPS)
RasterCapability = (lRet And Capability)


End Property

Public Property Get TextCapability(ByVal Capability As enTextCapabilities) As Boolean

Dim lRet As Long

lRet = GetDeviceCaps(cTEXTCAPS)
TextCapability = (lRet And Capability)

End Property

Public Property Get width(ByVal MeasurementScale As enDeviceMesaurementScale) As Long

If MeasurementScale = DMS_Millimeters Then
    width = GetDeviceCaps(cHORZSIZE)
Else
    width = GetDeviceCaps(cHORZRES)
End If

End Property


