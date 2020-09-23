VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiCreateStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hwndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As Long
    lpszClass As Long
    ExStyle As Long
End Type

Public lpCreateParams As Long
Public hInstance As Long
Public hMenu As Long
Public hwndParent As Long
Public cy As Long
Public cx As Long
Public y As Long
Public x As Long
Public style As Long
Public lpszName As String
Public lpszClass As String
Public ExStyle As Long

Public CreatedOK As Boolean

'\\ Private memory handling functions
Private Declare Sub CopyMemoryCreateStruct Lib "kernel32" Alias "RtlMoveMemory" (Destination As CREATESTRUCT, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrCreateStruct Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrCreateStruct Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long


'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this CreateStructobject from the location poiunted to by
'\\ lpCreateStruct
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for correctness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpCreateStruct As Long) As Boolean

Dim ftThis As CREATESTRUCT

CreatedOK = False

If Not IsBadReadPtrCreateStruct(lpCreateStruct, Len(ftThis)) Then
    Call CopyMemoryCreateStruct(ftThis, lpCreateStruct, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            cx = .cx
            cy = .cy
            ExStyle = .ExStyle
            hInstance = .hInstance
            hMenu = .hMenu
            hwndParent = .hwndParent
            lpCreateParams = .lpCreateParams
            lpszClass = .lpszClass
            lpszName = .lpszName
            style = .style
            x = .x
            y = .y
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function
