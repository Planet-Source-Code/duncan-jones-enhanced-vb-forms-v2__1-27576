VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A single menu item."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private m_HMenu As Long

Private mSubmenus As colMenu
Private mMenuItemInfo As APIMenuItemInfo

Private Declare Function GetMenuItemCountApi Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenuApi Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptX As Long, ByVal ptY As Long) As Long

Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Enum enMenuIndexMethod
    MF_BYCOMMAND = &H0&
    MF_BYPOSITION = &H400&
End Enum

Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long


Public Enum enMenuStates
    MF_CHECKED = &H8&
    MF_ENABLED = &H0&
    MF_HILITE = &H80&
    MF_DISABLED = &H2&
    MF_GRAYED = &H1&
End Enum

'\\ used to determine if this menu is valid...
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long

Public Property Get Caption(ByVal nPos As Long) As String
Attribute Caption.VB_Description = "Returns the string in the given menu position"

Dim sRet As String
Dim lRet As Long

sRet = String$(1024, 0)

lRet = GetMenuString(m_HMenu, nPos, sRet, 1024, MF_BYPOSITION)
If lRet = 0 Or Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiMenu:Caption", APIDispenser.LastSystemError)
End If
If lRet > 0 Then
    sRet = Left$(sRet, lRet)
End If
Caption = sRet

End Property

Public Property Get CaptionFromPoint(ByVal Parent As ApiWindow, ByVal Point As APIPoint) As String

Dim sRet As String
Dim lRet As Long

lRet = Me.IndexFromPoint(Parent, Point)
If lRet > -1 Then
    sRet = Me.Caption(lRet)
End If
CaptionFromPoint = sRet

End Property

Public Property Get Checked(ByVal nPos As Long) As Boolean

    Checked = (MenuState(nPos, MF_BYPOSITION) Or MF_CHECKED)

End Property


Public Property Get hMenu() As Long

    hMenu = m_HMenu
    
End Property

Public Property Let hMenu(ByVal newHandle As Long)

    If newHandle <> m_HMenu Then
        m_HMenu = newHandle
    End If
    
End Property
Public Property Get Grayed(ByVal nPos As Long) As Boolean

    Grayed = (MenuState(nPos, MF_BYPOSITION) Or MF_GRAYED)
    
End Property

Public Property Get IndexFromPoint(ByVal Parent As ApiWindow, ByVal Point As APIPoint) As Long

Dim lRet As Long
Dim x As Long, y As Long

With Point
    x = .x
    y = .y
End With

lRet = MenuItemFromPoint(Parent.hwnd, m_HMenu, x, y)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiMenu:IndexFromPoint", APIDispenser.LastSystemError)
    IndexFromPoint = -1
Else
    IndexFromPoint = lRet
End If

End Property

Public Property Get Popup(ByVal nPos As Integer) As Boolean

    Popup = (ItemId(nPos) = -1)

End Property

Public Property Get Separator(ByVal nPos As Long) As Boolean

    Separator = (ItemId(nPos) = 0)

End Property

Public Property Get ItemCount() As Long

Dim lRet As Long

lRet = GetMenuItemCount(m_HMenu)
If (lRet = -1) Or (Err.LastDllError > 0) Then
    Call ReportError(Err.LastDllError, "ApiMenu:ItemCount", APIDispenser.LastSystemError)
End If

ItemCount = lRet

End Property

Public Property Get ItemId(ByVal nPos As Long) As Long

    Dim lRet As Long
    
    lRet = GetMenuItemID(m_HMenu, nPos)
    If Err.LastDllError > 0 Then
        Call ReportError(Err.LastDllError, "ApiMenu:ItemId", APIDispenser.LastSystemError)
    End If
    
    ItemId = lRet
    
End Property


Public Property Get Key() As String

Key = "HMENU:" & m_HMenu

End Property


Public Property Set MenuItemInfo(thisItem As APIMenuItemInfo)

If thisItem Is Nothing Then
    Set mMenuItemInfo = Nothing
Else
    Set mMenuItemInfo = thisItem
End If

End Property

Public Property Get MenuItemInfo() As APIMenuItemInfo

If mMenuItemInfo Is Nothing Then
    Set mMenuItemInfo = New APIMenuItemInfo
End If
Set MenuItemInfo = mMenuItemInfo

End Property

Public Property Get MenuState(ByVal nIndex As Long, ByVal IndexMethod As enMenuIndexMethod) As Long

Dim lRet As Long

lRet = GetMenuState(m_HMenu, nIndex, IndexMethod)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiMenu:MenuState", APIDispenser.LastSystemError)
End If
MenuState = lRet

End Property


Public Property Get Submenus() As colMenu

Dim lSubmenus As Long
Dim lThismenu As Long
Dim hThismenu As Long

Dim mnuItem As ApiMenu
Dim mnuItemInfo As APIMenuItemInfo

Set mSubmenus = New colMenu
Set mnuItemInfo = New APIMenuItemInfo

lSubmenus = GetMenuItemCountApi(m_HMenu)
If lSubmenus > 0 Then
    For lThismenu = 0 To (lSubmenus - 1)
        hThismenu = GetSubMenuApi(m_HMenu, lThismenu)
        If hThismenu > 0 And Err.LastDllError = 0 Then
            Set mnuItem = mSubmenus.Add(hThismenu)
            mnuItem.hMenu = hThismenu
            Set mnuItem.MenuItemInfo = mnuItemInfo.GetMenuItemInfo(m_HMenu, lThismenu, True)
        End If
    Next lThismenu
End If

Set Submenus = mSubmenus

End Property


Public Property Get Valid() As Boolean
Attribute Valid.VB_Description = "Use to determine whether a menu is valid before calling any other of its methods."
Dim lRet As Long

lRet = IsMenu(m_HMenu)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiMenu:Valid", APIDispenser.LastSystemError)
Else
    Valid = (lRet <> 0)
End If

End Property

Private Sub Class_Terminate()

Set mSubmenus = Nothing
Set mMenuItemInfo = Nothing

End Sub


