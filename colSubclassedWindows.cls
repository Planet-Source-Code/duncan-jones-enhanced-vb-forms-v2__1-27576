VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colSubclassedWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection


Public Sub Add(ByVal wnd As ApiWindow)

On Local Error Resume Next

If mCol Is Nothing Then
    Set mCol = New Collection
End If

mCol.Add wnd, wnd.Key
'\\ Start subclassing it.....
Call wnd.SetWindowLong(GWL_WNDPROC, AddressOf VB_WindowProc)

End Sub

Public Function Item(ByVal hwnd As Long) As ApiWindow

On Error Resume Next
Set Item = mCol.Item("HWND:" & hwnd)

End Function


Public Sub Remove(wnd As ApiWindow)

If Not (mCol.Item(wnd.Key) Is Nothing) Then
    '\\ Stop subclassing it.....
    
    '\\ and remove it from our collection
    mCol.Remove wnd.Key
    
End If

End Sub

Private Sub Class_Initialize()

    Set mCol = New Collection

End Sub

Private Sub Class_Terminate()


    Set mCol = Nothing

End Sub

