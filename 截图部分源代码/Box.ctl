VERSION 5.00
Begin VB.UserControl Box 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3285
      Top             =   540
   End
End
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private mMouseIn As Boolean

Public Event Mouse(MouseEvent As BoxEvent, X As Variant, Y As Variant)
Public Enum BoxEvent
Init = 0
BoxEnter = 1
BoxMove = 2
BoxDown = 3
BoxUp = 4
BoxExit = 5
Terminate = 6
GotFocus = 7
LostFocus = 8
MouseIn = 9
End Enum
Public HasFocus As Boolean, IsMouseIn As Boolean, IsMouseDown As Boolean

Private Sub Timer1_Timer()
RaiseEvent Mouse(GotFocus, 0, 0)
End Sub

Private Sub UserControl_GotFocus()
RaiseEvent Mouse(GotFocus, 0, 0)
HasFocus = True
End Sub

Private Sub UserControl_LostFocus()
HasFocus = False
RaiseEvent Mouse(LostFocus, 0, 0)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
IsMouseDown = True
RaiseEvent Mouse(BoxDown, X, Y)
  SetCapture UserControl.hwnd
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Mouse(BoxMove, X, Y)
If Not mMouseIn Then
  mMouseIn = True
  IsMouseIn = True
  RaiseEvent Mouse(BoxEnter, X, Y)
  SetCapture UserControl.hwnd
ElseIf (X < 0 Or Y < 0 Or X > XPixel(UserControl.Width) Or Y > YPixel(UserControl.Height)) Then
  mMouseIn = False
  IsMouseIn = False
  RaiseEvent Mouse(BoxExit, X, Y)
  ReleaseCapture
End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
IsMouseDown = False
RaiseEvent Mouse(BoxUp, X, Y)
  SetCapture UserControl.hwnd
End Sub
Public Function Hdc()
Hdc = UserControl.Hdc
End Function
Public Sub Refresh()
UserControl.Refresh
End Sub
Public Sub Create()
RaiseEvent Mouse(Init, 0, 0)
End Sub
Public Sub Delete()
RaiseEvent Mouse(Terminate, 0, 0)
End Sub
