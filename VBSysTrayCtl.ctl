VERSION 5.00
Begin VB.UserControl VBSysTrayCtl 
   BackStyle       =   0  'Transparent
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VBSysTrayCtl.ctx":0000
   ScaleHeight     =   240
   ScaleWidth      =   240
   ToolboxBitmap   =   "VBSysTrayCtl.ctx":0312
End
Attribute VB_Name = "VBSysTrayCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNotify As ApiNotifyIcon
Private WithEvents mWnd As ApiWindow
Attribute mWnd.VB_VarHelpID = -1
Private WithEvents mAPI As APIFunctions
Attribute mAPI.VB_VarHelpID = -1

Public Event MouseMove()
Public Event MouseDown(ByVal Button As Integer)
Public Event MouseUp(ByVal Button As Integer)
Public Event MouseDblClick(ByVal Button As Integer)


Public Sub Hideicon()

    mNotify.UnsetNotifyIcon
    
End Sub

Public Property Get Icon() As ApiIcon

    Set Icon = mNotify.Icon
    
End Property


Public Property Let IconHandle(ByVal newHandle As Long)

    Dim newIco As ApiIcon
    
    Set newIco = New ApiIcon
    
    newIco.hIcon = newHandle
    Set mNotify.Icon = newIco
    
End Property

Public Property Get IconHandle() As Long

    IconHandle = mNotify.Icon.hIcon
    
End Property

Public Sub Refresh()

    mNotify.RefreshNotifyIcon
    
End Sub


Public Sub ShowIcon()

    mNotify.SetNotifyIcon
    
End Sub


Public Property Let Tooltip(ByVal newTip As String)

If mNotify.Tooltip <> newTip Then
    mNotify.Tooltip = newTip
    mNotify.RefreshNotifyIcon
End If

End Property

Public Property Get Tooltip() As String

    Tooltip = mNotify.Tooltip
    
End Property


Private Sub mWnd_WindowMessageFired(ByVal msg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean, ProcRet As Long)

Dim msgDecode As WindowMessages

If msg = mNotify.NotifyWindowMessage Then
    msgDecode = lParam
    Select Case msgDecode
    Case WM_MOUSEMOVE
        RaiseEvent MouseMove
    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(vbLeftButton)
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(vbRightButton)
    
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(vbLeftButton)
    Case WM_RBUTTONUP
        RaiseEvent MouseUp(vbRightButton)
    
    Case WM_LBUTTONDBLCLK
        RaiseEvent MouseDblClick(vbLeftButton)
    Case WM_RBUTTONDBLCLK
        RaiseEvent MouseDblClick(vbRightButton)
        
    Case Else
        '\\ UHANDLED MESSAGE
    End Select
End If

End Sub

Private Sub UserControl_Initialize()

Set mAPI = New APIFunctions

Set mWnd = New ApiWindow
mWnd.hwnd = UserControl.hwnd
mAPI.SubclassedWindows.Add mWnd

Set mNotify = New ApiNotifyIcon
Set mNotify.NotifyWindow = mWnd

End Sub

Private Sub UserControl_Terminate()

Set mNotify = Nothing
Set mWnd = Nothing
Set mAPI = Nothing

End Sub


