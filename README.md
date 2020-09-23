<div align="center">

## Enhanced VB Forms \(v2\)


</div>

### Description

2 controls that vastly extend the cababilities of Visual basic.

*** VBEventWindow - provides a simple subclassing control.

Events:

- ActiveApplicationChanged, fired when your app gains or loses user focus

- LostCapture, fired when your app gains or loses the capture

- KeyPressed, fired when any of the keys are pressed

- LowMemory, fired when the system is running low on memory

- Move, fired when the form is moved

- VerticalScroll, HorizontalScroll, fired when the form scrollbars are set

- WindowsSettingsChanged, WindowsINIChanged , fired when the windows environment settings are changed

- NonClientMouseMove,NonClientMouseDown,NonClientMouseUp,NonClientDblClick, fired when a mouse event occurs in the non-client part of your form

- MinMaxSize, fired when the OS wants to know what size to make your form either in response to a minimise/maximise command or when the user is dragging the resize box.

- MouseOverMenu, fired when the mouse is over a top level menu

- WindowMessageFired fired for all the other windows messages

Methods:

- InvalidateRect, Sets part of the form invalid to indicate that it needs to be repainted

Properties:

- ClassName, returns the windows class name fo the form

- DeviceContext, returns the device contect of the form (for graphical operations)

- HorizontalScrollbar, VerticalScrollbar, sets or unsets scrollbars on the form

- TopMost sets the form to float over the top of other forms

- Transparent, makes the formclient area invisible

Use:

In the form load...

Private Sub Form_Load()

Me.VBEventWindow.ParentForm = Me.hWnd

End Sub

*** VBSysTrayCtl - Provides a simple control to allow your application to use the SysTray

Events:

- MouseMove, Fired when the mouse moves over the tray icon

- MouseDown, Fired when a mouse down event occurs over the tray icon

- MouseUp, Fired when a mouse up event occurs over the tray icon

- MouseDblClick, Fired when the user double clicks the Tray Icon

Methods:

- ShowIcon, displays the icon in the system tray area

- Hideicon, removes the icon from the system tray area

- Refresh, updates the icon displayed in the system tray area

Properties:

- Tooltip, the tip that is displayed if the user hovers the mouse over your systray icon

Use:

In the form load....

Private Sub Form_Load()

Me.VBSysTrayCtl1.Tooltip = "Merrion Computing"

Me.VBSysTrayCtl1.ShowIcon

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Me.VBSysTrayCtl1.Hideicon

End Sub
 
### More Info
 
The functionality of this control is quite complex.

Fuller help is provided at http://www.MerrionComputing.com/Help/index.htm


<span>             |<span>
---                |---
**Submitted On**   |2001-09-27 10:00:18
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Advanced
**User Rating**    |5.0 (55 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Enhanced V270409272001\.zip](https://github.com/Planet-Source-Code/duncan-jones-enhanced-vb-forms-v2__1-27576/archive/master.zip)

### API Declarations

This source code has examples of a huge number of API calls. It is worth downloading for that alone.





