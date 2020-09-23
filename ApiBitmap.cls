VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mhBitmap As Long

Private Declare Function GetBitmapDimensionExApi Lib "gdi32" Alias "GetBitmapDimensionEx" (ByVal hBitmap As Long, lpDimension As Size) As Long
Private Type Size
    cx As Long
    cy As Long
End Type


Private mStockBitmap As Boolean

Private Declare Function DeleteObjectApi Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long

Private Declare Function CreateCompatibleBitmapApi Lib "gdi32" Alias "CreateCompatibleBitmap" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Friend Sub CreateCompatibleBitmap(ByVal hdc As Long, ByVal width As Long, ByVal Height As Long)

Dim lRet As Long

lRet = CreateCompatibleBitmapApi(hdc, width, Height)
If lRet = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "CreateCompatibleBitmap", APIDispenser.LastSystemError)
Else
    mhBitmap = lRet
End If

End Sub


Public Property Get hBitmap() As Long

    hBitmap = mhBitmap
    
End Property

Public Property Let hBitmap(ByVal newHandle As Long)

    If mhBitmap <> newHandle Then
        mhBitmap = newHandle
    End If
    
End Property

'\\ --[Height]------------------------------------------------------------------------------
'\\ Returns the height of a bitmap in logical pixels
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Property Get Height() As Long

Dim nSize As Size
Dim lRet As Long

lRet = GetBitmapDimensionExApi(mhBitmap, nSize)
If lRet = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiBitmap:Height", APIDispenser.LastSystemError)
Else
    Height = nSize.cy
End If

End Property

'\\ --[Width]------------------------------------------------------------------------------
'\\ Returns the width of a bitmap in logical pixels
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Property Get width() As Long

Dim nSize As Size
Dim lRet As Long

lRet = GetBitmapDimensionExApi(mhBitmap, nSize)
If lRet = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiBitmap:Width", APIDispenser.LastSystemError)
Else
    width = nSize.cx
End If

End Property

Private Sub Class_Terminate()

Dim lRet As Long

If mhBitmap > 0 And Not (mStockBitmap) Then
    lRet = DeleteObjectApi(mhBitmap)
    If lRet = 0 Or Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiBitmap:Terminate", APIDispenser.LastSystemError)
    End If
End If

End Sub


