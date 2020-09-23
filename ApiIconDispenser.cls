VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiIconDispenser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

'\\ --[IconsFromFilename]-------------------------------------------------------------------
'\\ Returns a collection of ApiIcon objects from the filename given.
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Property Get IconsFromExeFilename(ByVal Filename As String) As Collection

Dim lIndex As Long
Dim lIconCount As Long
Dim lRet As Long

Dim colIcons As Collection
Dim thisIcon As ApiIcon

'\\ Initialise the collection
Set colIcons = New Collection

'\\ Get the number of items
lIconCount = ExtractIcon(App.hInstance, Filename, -1)
If lIconCount > 0 Then
    For lIndex = 0 To lIconCount - 1
        lRet = ExtractIcon(App.hInstance, Filename, lIndex)
        If lRet > 0 Then
            Set thisIcon = New ApiIcon
            thisIcon.hIcon = lRet
            colIcons.Add thisIcon
        End If
    Next lIndex
End If

'\\ Return the collection
Set IconsFromExeFilename = colIcons

End Property
