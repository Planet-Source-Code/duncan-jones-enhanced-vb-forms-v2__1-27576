VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APIMenuItemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type MenuItemInfo
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long '\\ Converted to string after API calls
    cch As Long
End Type

Public cbSize As Long
Public fMask As Long
Public fType As Long
Public fState As Long
Public wID As Long
Public hSubMenu As Long
Public hbmpChecked As Long
Public hbmpUnchecked As Long
Public dwItemData As Long
Public dwTypeData As String
Public cch As Long

Public CreatedOK As Boolean
'\\ Private memory handling functions
Private Declare Sub CopyMemoryMenuItemInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As MenuItemInfo, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadReadPtrMenuItemInfo Lib "kernel32" Alias "IsBadReadPtr" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtrMenuItemInfo Lib "kernel32" Alias "IsBadWritePtr" (ByVal lp As Long, ByVal ucb As Long) As Long

Private Declare Function GetMenuItemInfoApi Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MenuItemInfo) As Long
Private Declare Function SetMenuItemInfoApi Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MenuItemInfo) As Long

'\\ --[CreateFromPointer]---------------------------------------------
'\\ Fills this MenuItemInfo object from the location poiunted to by
'\\ lpMenuItemInfo
'\\ VB.NET Porting note: This function should be replaced with an override
'\\ of the New() for corMenuItemInfoness
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function CreateFromPointer(lpMenuItemInfo As Long) As Boolean

Dim ftThis As MenuItemInfo

CreatedOK = False

If Not IsBadReadPtrMenuItemInfo(lpMenuItemInfo, Len(ftThis)) Then
    Call CopyMemoryMenuItemInfo(ftThis, lpMenuItemInfo, Len(ftThis))
    If Err.LastDllError = 0 Then
        With ftThis
            cbSize = .cbSize
            cch = .cch
            dwItemData = .dwItemData
            dwTypeData = .dwTypeData
            fMask = .fMask
            fState = .fState
            fType = .fType
            hbmpChecked = .hbmpChecked
            hbmpUnchecked = .hbmpUnchecked
            hSubMenu = .hSubMenu
            wID = .wID
            If Err.LastDllError = 0 Then
                CreatedOK = True
            End If
        End With
    End If
End If

CreateFromPointer = CreatedOK

End Function



Public Function GetMenuItemInfo(ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean) As APIMenuItemInfo

Dim lRet As Long
Dim miiThis As MenuItemInfo
Dim miiRet As APIMenuItemInfo

lRet = GetMenuItemInfoApi(hMenu, un, b, miiThis)
If Err.LastDllError Then
    Set miiRet = New APIMenuItemInfo
    If miiRet.CreateFromPointer(VarPtr(miiThis)) Then
        Set GetMenuItemInfo = miiRet
    End If
End If

End Function

Public Function SetMenuItemInfo(ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean) As Long

Dim lRet As Long
Dim miiThis As MenuItemInfo

With miiThis
    .cbSize = cbSize
    .cch = cch
    .dwItemData = dwItemData
    .dwTypeData = dwTypeData
    .fMask = fMask
    .fState = fState
    .fType = fType
    .hbmpChecked = hbmpChecked
    .hbmpUnchecked = hbmpUnchecked
    .hSubMenu = hSubMenu
    .wID = wID
End With

lRet = SetMenuItemInfoApi(hMenu, un, b, miiThis)
If Err.LastDllError = 0 Then
    SetMenuItemInfo = lRet
End If

End Function


