VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiCBT_CREATEWND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type CBT_CREATEWND
    csThis As Long '\\ Pointer to CREATESTRUCT
    hWndInsertAfter As Long
End Type


Public csThis As ApiCreateStruct
Public hWndInsertAfter As Long

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryCBT_CreateWnd Lib "kernel32" Alias "RtlMoveMemory" (Destination As CBT_CREATEWND, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrCBT_CreateWnd Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrCBT_CreateWnd Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this CBT_CreateWnd object from the location poiunted to by
'\\ lpCBT_CreateWnd
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpCBT_CreateWnd As Long) As Boolean

Dim ftThis As CBT_CREATEWND

CreatedOK = False

If Not IsBadReadPtrCBT_CreateWnd(lpCBT_CreateWnd, Len(ftThis)) Then
    Call CopyMemoryCBT_CreateWnd(ftThis, lpCBT_CreateWnd, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            .hWndInsertAfter = hWndInsertAfter
            Set csThis = New ApiCreateStruct
            If Not csThis.CreateFromPointer(.csThis) Then
                Set csThis = Nothing
            Else
                If Err.LastDllError = 0 Then
                    CreatedOK = True
                End If
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function
