VERSION 5.00
Begin VB.UserControl VBEventWindow 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   225
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VBEventWindow.ctx":0000
   ScaleHeight     =   225
   ScaleWidth      =   225
   ToolboxBitmap   =   "VBEventWindow.ctx":0312
End
Attribute VB_Name = "VBEventWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'\\ Control Member Variables....
Private mMaxHeight As Long, mMaxWidth As Long
Private mMaxPositionTop As Long, mMaxPositionLeft As Long
Private mMinTrackWidth As Long, mMinTrackheight As Long
Private mMaxTrackWidth As Long, mMaxTrackHeight As Long


'\\ From ApiWindows
Public Event WindowMessageFired(ByVal msg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean, ProcRet As Long)
Public Event ActiveApplicationChanged(ByVal ActivatingThisApp As Boolean, ByVal hThread As Long, Cancel As Boolean)
Public Event LostCapture(ByVal hwndNewCapture As Long, Cancel As Boolean)
Public Event KeyPressed(ByVal VKey As Long, ByVal Repetition As Long, ByVal ScanCode As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As Boolean, ByVal AlreadyDown As Boolean, ByVal BeingPressed As Boolean, Cancel As Boolean)
Public Event LowMemory(ByVal TimeSpentCompacting As Long)
Public Event Move(ByVal x As Long, ByVal y As Long, Cancel As Boolean)

Public Event VerticalScroll(ByVal Message As enScrollMessages, ByVal Position As Long, Cancel As Boolean)
Public Event HorizontalScroll(ByVal Message As enScrollMessages, ByVal Position As Long, Cancel As Boolean)

Public Event WindowsSettingsChanged(ByVal NewSetting As enSystemParametersInfo)
Public Event WindowsINIChanged(ByVal Section As String)

Public Event NonClientMouseMove(ByVal Location As enHitTestResult, ByVal x As Single, ByVal y As Single)
Public Event NonClientMouseDown(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
Public Event NonClientMouseUp(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
Public Event NonClientDblClick(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)

Public Event MinMaxSize(MaxHeight As Long, MaxWidth As Long, MaxPositionTop As Long, MaxPositionLeft As Long, MinTrackWidth As Long, MinTrackHeight As Long, MaxTrackWidth As Long, MaxTrackHeight As Long)

Public Event MouseOverMenu(ByVal Caption As String)

'\\ From Apifunctions
Public Event ApiError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)

Private WithEvents mWnd As ApiWindow
Attribute mWnd.VB_VarHelpID = -1
Private WithEvents mAPI As APIFunctions
Attribute mAPI.VB_VarHelpID = -1



Public Property Get ClassName() As String

    ClassName = mWnd.ClassName

End Property


Public Property Get DeviceContext() As ApiDeviceContext

    Set DeviceContext = mWnd.DeviceContext
    
End Property

Public Property Let HorizontalScrollbar(ByVal bSetting As Boolean)

If bSetting Then
    mWnd.SetWindowStyle WS_HSCROLL
Else
    mWnd.UnSetWindowStyle WS_HSCROLL
End If
mWnd.Refresh

End Property

Public Property Let MaxHeight(ByVal mHeight As Integer)

    If mHeight > 0 Then
        mMaxHeight = mHeight
    End If
    
End Property

Public Property Get MaxHeight() As Integer

    MaxHeight = mMaxHeight
    
End Property

Public Property Get MaxPositionLeft() As Integer

    MaxPositionLeft = mMaxPositionLeft
    
End Property

Public Property Let MaxPositionTop(ByVal nTop As Integer)

    mMaxPositionTop = nTop
    
End Property

Public Property Get MaxPositionTop() As Integer

    MaxPositionTop = mMaxPositionTop

End Property

Public Property Let MaxTrackHeight(ByVal mHeight As Integer)

    If mHeight > mMinTrackheight Then
        mMaxTrackHeight = mHeight
    End If
    
End Property

Public Property Get MaxTrackHeight() As Integer

    MaxTrackHeight = mMaxTrackHeight
    
End Property

Public Property Let MaxTrackWidth(ByVal mWidth As Integer)

    If mWidth >= mMinTrackWidth Then
        mMaxTrackWidth = mWidth
    End If
    
End Property


Public Property Get MaxTrackWidth() As Integer

    MaxTrackWidth = mMaxTrackWidth
    
End Property

Public Property Let MaxWidth(ByVal mWidth As Integer)

    If mWidth > 0 Then
        mMaxWidth = mWidth
    End If
    
End Property

Public Property Get MaxWidth() As Integer

    MaxWidth = mMaxWidth
    
End Property


Public Property Let MinTrackHeight(ByVal mHeight As Integer)

    If mHeight > 0 Then
        mMinTrackheight = mHeight
    End If
    
End Property

Public Property Get MinTrackHeight() As Integer

    MinTrackHeight = mMinTrackheight
    
End Property

Public Property Let MinTrackWidth(ByVal mWidth As Integer)

    If mWidth > 0 Then
        mMinTrackWidth = mWidth
    End If
    
End Property

Public Property Get MinTrackWidth() As Integer

    MinTrackWidth = mMinTrackWidth
    
End Property

Public Property Let TopMost(ByVal newVal As Boolean)

If newVal Then
    mWnd.SetWindowStyle WS_EX_TOPMOST
Else
    mWnd.UnSetWindowStyle WS_EX_TOPMOST
End If
mWnd.Refresh

End Property

Public Property Get TopMost() As Boolean

    TopMost = mWnd.IsWindowStyleSet(WS_EX_TOPMOST)
    
End Property

Public Property Let Transparent(ByVal newVal As Boolean)

If newVal Then
    mWnd.SetWindowStyle WS_EX_TRANSPARENT
Else
    mWnd.UnSetWindowStyle WS_EX_TRANSPARENT
End If
mWnd.Refresh

End Property

Public Property Get Transparent() As Boolean

    Transparent = mWnd.IsWindowStyleSet(WS_EX_TRANSPARENT)
    
End Property

Public Property Let VerticalScrollbar(ByVal bSetting As Boolean)

If bSetting Then
    mWnd.SetWindowStyle WS_VSCROLL
Else
    mWnd.UnSetWindowStyle WS_VSCROLL
End If
mWnd.Refresh

End Property

Public Sub InvalidateRect(ByVal RectIn As APIRect)

    mWnd.InvalidateRect RectIn
    
End Sub



Public Property Let ParentForm(ByVal fParent As Long)

If mAPI Is Nothing Then
    Set mAPI = New APIFunctions
End If

If mWnd Is Nothing Then
    Set mWnd = New ApiWindow
End If

If mWnd.hwnd <> fParent Then
    mWnd.hwnd = fParent
    mAPI.SubclassedWindows.Add mWnd
End If

End Property

Public Property Get VerticalScrollbar() As Boolean

    VerticalScrollbar = mWnd.IsWindowStyleSet(WS_VSCROLL)

End Property

Public Property Get HorizontalScrollbar() As Boolean

    HorizontalScrollbar = mWnd.IsWindowStyleSet(WS_VSCROLL)

End Property

Private Sub mAPI_ApiError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)

    RaiseEvent ApiError(Number, Source, Description)

End Sub

Private Sub mWnd_ActiveApplicationChanged(ByVal ActivatingThisApp As Boolean, ByVal hThread As Long, Cancel As Boolean)

RaiseEvent ActiveApplicationChanged(ActivatingThisApp, hThread, Cancel)

End Sub

Private Sub mWnd_HorizontalScroll(ByVal Message As enScrollMessages, ByVal Position As Long, Cancel As Boolean)

RaiseEvent HorizontalScroll(Message, Position, Cancel)

End Sub


Private Sub mWnd_KeyPressed(ByVal VKey As Long, ByVal Repetition As Long, ByVal ScanCode As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As Boolean, ByVal AlreadyDown As Boolean, ByVal BeingPressed As Boolean, Cancel As Boolean)

RaiseEvent KeyPressed(VKey, Repetition, ScanCode, ExtendedKey, AltDown, AlreadyDown, BeingPressed, Cancel)

End Sub


Private Sub mWnd_LostCapture(ByVal hwndNewCapture As Long, Cancel As Boolean)

RaiseEvent LostCapture(hwndNewCapture, Cancel)

End Sub


Private Sub mWnd_LowMemory(ByVal TimeSpentCompacting As Long)

RaiseEvent LowMemory(TimeSpentCompacting)

End Sub


Private Sub mWnd_MinMaxSize(MaxHeight As Long, MaxWidth As Long, MaxPositionTop As Long, MaxPositionLeft As Long, MinTrackWidth As Long, MinTrackHeight As Long, MaxTrackWidth As Long, MaxTrackHeight As Long)

MaxHeight = Me.MaxHeight
MaxWidth = Me.MaxWidth
MaxPositionTop = Me.MaxPositionTop
MaxPositionLeft = Me.MaxPositionLeft
MaxTrackWidth = Me.MaxTrackWidth
MaxTrackHeight = Me.MaxTrackHeight
MinTrackWidth = Me.MinTrackWidth
MinTrackHeight = Me.MinTrackHeight

RaiseEvent MinMaxSize(MaxHeight, MaxWidth, MaxPositionTop, MaxPositionLeft, MinTrackWidth, MinTrackHeight, MaxTrackWidth, MaxTrackHeight)

End Sub


Private Sub mWnd_MouseOverMenu(ByVal Caption As String)

RaiseEvent MouseOverMenu(Caption)

End Sub


Private Sub mWnd_Move(ByVal x As Long, ByVal y As Long, Cancel As Boolean)

RaiseEvent Move(x, y, Cancel)

End Sub


Private Sub mWnd_NonClientDblClick(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)

RaiseEvent NonClientDblClick(Location, Button, x, y)

End Sub


Private Sub mWnd_NonClientMouseDown(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)

RaiseEvent NonClientMouseDown(Location, Button, x, y)

End Sub


Private Sub mWnd_NonClientMouseMove(ByVal Location As enHitTestResult, ByVal x As Single, ByVal y As Single)

RaiseEvent NonClientMouseMove(Location, x, y)

End Sub


Private Sub mWnd_NonClientMouseUp(ByVal Location As enHitTestResult, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)

RaiseEvent NonClientMouseUp(Location, Button, x, y)

End Sub


Private Sub mWnd_VerticalScroll(ByVal Message As enScrollMessages, ByVal Position As Long, Cancel As Boolean)

RaiseEvent VerticalScroll(Message, Position, Cancel)

End Sub


Private Sub mWnd_WindowMessageFired(ByVal msg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean, ProcRet As Long)

RaiseEvent WindowMessageFired(msg, wParam, lParam, Cancel, ProcRet)

End Sub


Private Sub mWnd_WindowsINIChanged(ByVal Section As String)

RaiseEvent WindowsINIChanged(Section)

End Sub


Private Sub mWnd_WindowsSettingsChanged(ByVal NewSetting As enSystemParametersInfo)

RaiseEvent WindowsSettingsChanged(NewSetting)

End Sub


Private Sub UserControl_Initialize()

Set mAPI = New APIFunctions
Set mWnd = New ApiWindow

End Sub

Private Sub UserControl_InitProperties()

'\\ Explicitly start the global class variables
mAPI.StartLink

'\\ Initialise the maximum/minimum size properties
mMaxHeight = 480
mMaxWidth = 640
mMaxPositionTop = 0
mMaxPositionLeft = 0
mMinTrackWidth = 100
mMinTrackheight = 50
mMaxTrackWidth = 640
mMaxTrackHeight = 480

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'\\ Read the maximum/minimum size properties
mMaxHeight = PropBag.ReadProperty("MaxHeight", 480)
mMaxWidth = PropBag.ReadProperty("MaxWidth", 640)
mMaxPositionTop = PropBag.ReadProperty("MaxPositionTop", 0)
mMaxPositionLeft = PropBag.ReadProperty("MaxPositionLeft", 0)
mMinTrackWidth = PropBag.ReadProperty("MinTrackWidth", 100)
mMinTrackheight = PropBag.ReadProperty("MinTrackHeight", 50)
mMaxTrackWidth = PropBag.ReadProperty("MaxTrackWidth", 640)
mMaxTrackHeight = PropBag.ReadProperty("MaxTrackHeight", 480)

End Sub


Private Sub UserControl_Terminate()

Set mWnd = Nothing
Set mAPI = Nothing

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'\\ Write the maximum/minimum size properties
Call PropBag.WriteProperty("MaxHeight", mMaxHeight, 480)
Call PropBag.WriteProperty("MaxWidth", mMaxWidth, 640)
Call PropBag.WriteProperty("MaxPositionTop", mMaxPositionTop, 0)
Call PropBag.WriteProperty("MaxPositionLeft", mMaxPositionLeft, 0)
Call PropBag.WriteProperty("MinTrackWidth", mMinTrackWidth, 100)
Call PropBag.WriteProperty("MinTrackHeight", mMinTrackheight, 50)
Call PropBag.WriteProperty("MaxTrackWidth", mMaxTrackWidth, 640)
Call PropBag.WriteProperty("MaxTrackHeight", mMaxTrackHeight, 480)

End Sub


