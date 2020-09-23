VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As Long 'MSG
End Type

Public hwnd As Long
Public Message As Long
Public wParam As Long
Public lParam As Long
Public time As Long
Public pt As APIPoint

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryMSG Lib "kernel32" Alias "RtlMoveMemory" (Destination As msg, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrMSG Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrMSG Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this MSG object from the location poiunted to by
'\\ lpMSG
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for corMSGness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpMSG As Long) As Boolean

Dim ftThis As msg


CreatedOK = False

If Not IsBadReadPtrMSG(lpMSG, Len(ftThis)) Then
    Call CopyMemoryMSG(ftThis, lpMSG, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            .hwnd = .hwnd
            .lParam = .lParam
            .Message = .Message
            .time = .time
            .wParam = .wParam
            Set pt = New APIPoint
            If pt.CreateFromPointer(.pt) Then
                If Err.LastDllError = 0 Then
                    CreatedOK = True
                End If
            Else
                Set pt = Nothing
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function
