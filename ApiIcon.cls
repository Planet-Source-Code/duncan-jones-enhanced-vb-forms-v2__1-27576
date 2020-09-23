VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function LoadIconApi Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function GetIconInfoApi Lib "user32" Alias "GetIconInfo" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private mhIcon As Long

Private mStockIcon As Boolean

'\\ Stock icons
Public Enum enStockIcons
    IDI_WINLOGO = 32517&
    IDI_APPLICATION = 32512&
    IDI_ASTERISK = 32516&
    IDI_EXCLAMATION = 32515&
    IDI_HAND = 32513&
    IDI_QUESTION = 32514&
    IDI_UNKNOWN = 32518&
End Enum



Public Property Get ColourBitmap() As ApiBitmap

Dim info As ICONINFO

Dim bmThis As ApiBitmap
Dim lRet As Long

Set bmThis = New ApiBitmap

lRet = GetIconInfoApi(mhIcon, info)
If lRet = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiIcon:ColourBitmap", APIDispenser.LastSystemError)
End If

With info
    bmThis.hBitmap = .hbmColor
End With

Set ColourBitmap = bmThis

End Property

Public Property Let hIcon(ByVal newHandle As Long)

If newHandle <> mhIcon Then
    mhIcon = newHandle
End If

End Property

Public Property Get hIcon() As Long
    
    hIcon = mhIcon
    
End Property

Public Property Get Hotspot() As APIPoint

Dim info As ICONINFO

Dim pointThis As APIPoint
Dim lRet As Long

Set pointThis = New APIPoint

lRet = GetIconInfoApi(mhIcon, info)
If lRet = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiIcon:HotSpot", APIDispenser.LastSystemError)
End If
With info
    pointThis.x = .xHotspot
    pointThis.y = .yHotspot
End With

Set Hotspot = pointThis

End Property

Public Sub LoadSystemIcon(ByVal StockIcon As enStockIcons)

Dim lRet As Long

lRet = LoadIconApi(0, StockIcon)
If Err.LastDllError = 0 Then
    Me.hIcon = lRet
    '\\ Flag that this icon was loded from a stock icon...
    mStockIcon = True
End If

End Sub


Public Property Get MaskBitmap() As ApiBitmap

Dim info As ICONINFO

Dim bmThis As ApiBitmap
Dim lRet As Long

Set bmThis = New ApiBitmap

lRet = GetIconInfoApi(mhIcon, info)
If lRet = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiIcon:MaskBitmap", APIDispenser.LastSystemError)
End If

With info
    bmThis.hBitmap = .hbmMask
End With

Set MaskBitmap = bmThis

End Property


Private Sub Class_Terminate()
Dim lRet As Long

If mhIcon > 0 And Not (mStockIcon) Then
    lRet = DestroyIcon(mhIcon)
    If lRet = 0 Or Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiIcon:Terminate", APIDispenser.LastSystemError)
    End If
End If

End Sub


