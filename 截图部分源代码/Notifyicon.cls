VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotifyIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private ND As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIF_ICON = &H2, NIF_MESSAGE = &H1, NIF_TIP = &H4, NIM_ADD = &H0, NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200, WM_LBUTTONUP = &H202, WM_RBUTTONUP = &H205
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hicon As Long
        SzTip As String * 64
End Type

Public Event MouseEvent(style As Integer)
Public Event LMouseUp()
Public Event RMouseUp()
Public Event MouseMove(x As Single, y As Single)

Public Sub AddNotifyIcon(Form As Form, Optional hicon As Long, Optional SzTip As String)
Set mForm = Form
With ND
 .hwnd = Form.hwnd
 .uID = Form.Icon
 If SzTip = "" Then SzTip = Form.Caption
 .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
 .uCallbackMessage = WM_MOUSEMOVE
 .hicon = Form.Icon.Handle
 .SzTip = SzTip & vbNullChar
 .cbSize = Len(ND)
End With
Shell_NotifyIcon NIM_ADD, ND
End Sub
Public Sub DeleteNotifyIcon()
Call Class_Terminate
End Sub
Private Sub Class_Terminate()
Shell_NotifyIcon NIM_DELETE, ND
End Sub
Private Sub mForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim IMSG As Integer
IMSG = x / Screen.TwipsPerPixelX
Select Case IMSG
 Case Is = WM_RBUTTONUP
    RaiseEvent MouseEvent(2)
    RaiseEvent RMouseUp
 Case Is = WM_LBUTTONUP
   RaiseEvent MouseEvent(1)
   RaiseEvent LMouseUp
 Case Is = WM_MOUSEMOVE
   RaiseEvent MouseMove(x, y)
End Select
End Sub
