VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APIRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public CreatedOK As Boolean
Public Left As Long
Public Top As Long
Public Right As Long
Public Bottom As Long

'\\ Private memory handling functions
Private Declare Sub CopyMemoryRect Lib "kernel32" Alias "RtlMoveMemory" (Destination As RECT, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromRect Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As RECT, ByVal Length As Long)
Private Declare Function IsBadReadPtrRect Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrRect Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ RECT manipulation....
Private Declare Function InvalidateRectAPI Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetMenuItemRectAPI Lib "user32" Alias "GetMenuItemRect" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this Rect object from the location poiunted to by
'\\ lpRect
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpRect As Long) As Boolean

Dim ftThis As RECT

CreatedOK = False

If Not IsBadReadPtrRect(lpRect, Len(ftThis)) Then
    Call CopyMemoryRect(ftThis, lpRect, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            Left = .Left
            Right = .Right
            Top = .Top
            Bottom = .Bottom
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function

Public Function CreateFromWindow(ByVal hwnd As Long) As Boolean

Dim lRet As Long
Dim lpRect As RECT

CreatedOK = False
lRet = GetWindowRect(hwnd, lpRect)
If Err.LastDllError = 0 Then
    With lpRect
        Left = .Left
        Right = .Right
        Top = .Top
        Bottom = .Bottom
    End With
    If Err.LastDllError = 0 Then
        CreatedOK = True
    End If
End If

CreateFromWindow = CreatedOK

End Function


Public Function GetMenuItemRect(ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long) As APIRect

Dim rcThis As RECT
Dim rcRet As APIRect
Dim lRet As Long

lRet = GetMenuItemRectAPI(hwnd, hMenu, uItem, rcThis)
If Err.LastDllError = 0 Then
    Set rcRet = New APIRect
    If rcRet.CreateFromPointer(VarPtr(rcThis)) Then
        Set GetMenuItemRect = rcRet
    End If
End If

End Function

Public Function InvalidateRect(ByVal hwnd As Long, ByVal uErase As Boolean) As Long

Dim lRet As Long
Dim rcThis As RECT

With rcThis
    .Left = Left
    .Right = Right
    .Top = Top
    .Bottom = Bottom
End With

lRet = InvalidateRectAPI(hwnd, rcThis, uErase)
If Err.LastDllError = 0 Then
    InvalidateRect = lRet
End If

End Function


Public Sub SaveToPointer(lpRect As Long)

Dim ftThis As RECT

With ftThis
    .Left = Left
    .Right = Right
    .Top = Top
    .Bottom = Bottom
End With

If Not IsBadReadPtrRect(lpRect, Len(ftThis)) Then
    Call CopyMemoryFromRect(lpRect, ftThis, Len(ftThis))
    If Err.LastDllError <> 0 Then
        ReportError Err.LastDllError, "ApiRect:SaveToPointer", APIDispenser.LastSystemError
    End If
End If


End Sub


