VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiWindowDispenser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIsCommonContolsInitialsed As Boolean

'\\ Creating new windows.....
Private Declare Function CreateWindowExApi Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Public Function CreateWindow(ByVal ExtendedWindowStyle As Long, ByVal TheClass As APIWndClass, ByVal WindowName As String, ByVal WindowStyle As Long, ByVal xPos As Long, ByVal yPos As Long, ByVal width As Long, ByVal Height As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInst As Long, ByVal lpCreateStruct As Long) As ApiWindow

Dim wndThis As ApiWindow
Dim lRet As Long

lRet = CreateWindowExApi(ExtendedWindowStyle, TheClass.lpszClassName, WindowName, WindowStyle, xPos, yPos, width, Height, hwndParent, hMenu, hInst, lpCreateStruct)
If lRet > 0 And Err.LastDllError = 0 Then
    Set wndThis = New ApiWindow
    wndThis.hwnd = lRet
Else
    Call ReportError(Err.LastDllError, "ApiWindowDispenser:CreateWindow", APIDispenser.LastSystemError)
End If

End Function
