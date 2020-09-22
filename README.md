<div align="center">

## Systray Toolbar Bug


</div>

### Description

When a form contains a toolbar (from MS Common Controls 6.0), there is a bug in vb6 that causes the systray icon to not respond to mouse events. The workaround for this is to set the toolbar.enabled = False before creating the systray icon
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ben White](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ben-white.md)
**Level**          |Beginner
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ben-white-systray-toolbar-bug__1-33673/archive/master.zip)

### API Declarations

```
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
 Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
```


### Source Code

```
'*** This is a very simplistic Example
'*** Create a form with a toolbar named Toolbar1 and place the following code in it
'-------------
' FORM CODE
'-------------
Private Sub Form_Resize()
 If Me.WindowState = vbMinimized Then
  Toolbar1.Enabled = False
  TrayAdd Me
 End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Result As Long
 Select Case TrayEvent(Me, X)
  Case 1:
   TrayRestore Me
   Toolbar1.Enabled = True
  Case 2:
 End Select
End Sub
'**** this is generic code not written, only modified by myself
'*** Create a module and place the following code in it
'-------------
' MODULE CODE
'-------------
 'user defined type required by Shell_NotifyIcon API call
 Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 64
 End Type
 'constants required by Shell_NotifyIcon API call:
 Private Const NIM_ADD = &H0
 Private Const NIM_MODIFY = &H1
 Private Const NIM_DELETE = &H2
 Private Const NIF_MESSAGE = &H1
 Private Const NIF_ICON = &H2
 Private Const NIF_TIP = &H4
 Private Const WM_MOUSEMOVE = &H200
 Private Const WM_LBUTTONDOWN = &H201  'Button down
 Private Const WM_LBUTTONUP = &H202  'Button up
 Private Const WM_LBUTTONDBLCLK = &H203 'Double-click
 Private Const WM_RBUTTONDOWN = &H204  'Button down
 Private Const WM_RBUTTONUP = &H205  'Button up
 Private Const WM_RBUTTONDBLCLK = &H206 'Double-click
 Private nid As NOTIFYICONDATA
 Public Sub TrayAdd(frmWindow As Form)
  With nid
   .cbSize = Len(nid)
   .hwnd = frmWindow.hwnd
   .uId = vbNull
   .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   .uCallBackMessage = WM_MOUSEMOVE
   .hIcon = frmWindow.Icon
   .szTip = App.Title & vbNullChar
  End With
  Shell_NotifyIcon NIM_ADD, nid
  frmWindow.Hide
 End Sub
 Public Sub TrayDel()
  Shell_NotifyIcon NIM_DELETE, nid
 End Sub
 Public Sub TrayEdit(frmWindow As Form)
  With nid
   .cbSize = Len(nid)
   .hwnd = frmWindow.hwnd
   .uId = vbNull
   .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   .uCallBackMessage = WM_MOUSEMOVE
   .hIcon = frmWindow.Icon
   .szTip = App.Title & vbNullChar
  End With
  Shell_NotifyIcon NIM_MODIFY, nid
 End Sub
 Public Sub TrayRestore(frmWindow As Form)
  Dim Result As Long
  frmWindow.WindowState = vbNormal
  Result = SetForegroundWindow(frmWindow.hwnd)
  frmWindow.Show
  Shell_NotifyIcon NIM_DELETE, nid
 End Sub
 Public Function TrayEvent(frmWindow As Form, X As Single)
  'Call this sub from the mousemove
  'event on a form
  TrayEvent = 0
  Dim msg As Long
  If frmWindow.ScaleMode <> vbPixels Then
   msg = X / Screen.TwipsPerPixelX
  Else
   msg = X
  End If
  Select Case msg
   Case WM_LBUTTONDOWN
   Case WM_LBUTTONUP
    TrayEvent = 1
   Case WM_LBUTTONDBLCLK
   Case WM_RBUTTONDOWN
    TrayEvent = 2
   Case WM_RBUTTONUP
   Case WM_RBUTTONDBLCLK
  End Select
 End Function
```

