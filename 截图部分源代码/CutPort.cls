VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CutPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MouseIn As Boolean
Private x1 As Single, x2 As Single, y1 As Single, y2 As Single
Private offsetX As Single, offsetY As Single
Public FocusObj As FocusList

Public Property Get x() As Single
If x1 < x2 Then x = x1 Else x = x2
End Property
Public Property Get y() As Single
If y1 < y2 Then y = y1 Else y = y2
End Property
Public Property Get Width() As Single
Width = Abs(x1 - x2)
End Property
Public Property Get Height() As Single
Height = Abs(y1 - y2)
End Property


Public Sub ReceiveMouseEvent(EventType As MouseEvent, Ex As Single, Ey As Single, IsMouseDown As Boolean)
Select Case EventType
  Case Is = MouseEnter
    
  Case Is = MouseExit
     CutBoard.MousePointer = 0
  Case Is = MouseDown
    If Ex > x - 10 And Ex < x + 10 And Ey > y - 10 And Ey < y + 10 Then
      FocusObj = LeftTop
    ElseIf Ex > x + Width - 10 And Ex < x + Width + 10 And Ey > y - 10 And Ey < y + 10 Then
      FocusObj = RightTop
    ElseIf Ex > x - 10 And Ex < x + 10 And Ey > y + Height - 10 And Ey < y + Height + 10 Then
      FocusObj = LeftDown
    ElseIf Ex > x2 - 10 And Ex < x2 + 10 And Ey > y2 - 10 And Ey < y2 + 10 Then
      FocusObj = RightDown
    ElseIf Ex > x - 10 And Ex < x + 10 Then
      FocusObj = Left
    ElseIf Ex > x2 - 10 And Ex < x2 + 10 Then
      FocusObj = Right
    ElseIf Ey > y - 10 And Ey < y + 10 Then
      FocusObj = Top
    ElseIf Ey > y2 - 10 And Ey < y2 + 10 Then
      FocusObj = Botton
    Else
      FocusObj = Middle
      offsetX = Ex - x: offsetY = Ey - y
    End If
     
  Case Is = MouseMove
    If Ex > x - 10 And Ex < x + 10 And Ey > y - 10 And Ey < y + 10 Then
      CutBoard.MousePointer = 8
    ElseIf Ex > x + Width - 10 And Ex < x + Width + 10 And Ey > y - 10 And Ey < y + 10 Then
      CutBoard.MousePointer = 6
    ElseIf Ex > x - 10 And Ex < x + 10 And Ey > y + Height - 10 And Ey < y + Height + 10 Then
      CutBoard.MousePointer = 6
    ElseIf Ex > x2 - 10 And Ex < x2 + 10 And Ey > y2 - 10 And Ey < y2 + 10 Then
      CutBoard.MousePointer = 8
    ElseIf Ex > x - 10 And Ex < x + 10 Then
      CutBoard.MousePointer = 9
    ElseIf Ex > x2 - 10 And Ex < x2 + 10 Then
      CutBoard.MousePointer = 9
    ElseIf Ey > y - 10 And Ey < y + 10 Then
      CutBoard.MousePointer = 7
    ElseIf Ey > y2 - 10 And Ey < y2 + 10 Then
      CutBoard.MousePointer = 7
    Else
      If CutBoard.ControlMode = SizeMode Then CutBoard.MousePointer = 5
    End If
    
'    If Ex > X - 10 And Ex < X + 10 And Ey > Y - 10 And Ey < Y + 10 Then
'
'    ElseIf Ex > X + Width - 10 And Ex < X + Width + 10 And Ey > Y - 10 And Ey < Y + 10 Then
'
'    Else
'
'    End If
    
    Select Case FocusObj
      Case Is = LeftTop
         x1 = Ex
         y1 = Ey
      Case Is = RightTop
         y1 = Ey
         x2 = Ex
      Case Is = LeftDown
         x1 = Ex
         y2 = Ey
      Case Is = RightDown
         x2 = Ex
         y2 = Ey
      Case Is = Middle
         If CutBoard.ControlMode <> SizeMode Then Exit Sub
         x2 = Ex - offsetX + Width
         y2 = Ey - offsetY + Height
         x1 = Ex - offsetX: y1 = Ey - offsetY
      Case Is = Left
         x1 = Ex
      Case Is = Right
         x2 = Ex
      Case Is = Top
         y1 = Ey
      Case Is = Botton
         y2 = Ey
    End Select
  Case Is = MouseUp
    FocusObj = none
End Select
End Sub
Public Sub ReFresh()

GdipDrawRectangleI ViewGraphics, NewPen(&HFF00A8EC, 1), x, y, Width, Height
Dim White As Long
White = NewBrush(&HFFFFFFFF)
GdipFillEllipse ViewGraphics, White, x - 4, y - 4, 10, 10
GdipFillEllipse ViewGraphics, White, x + Width - 4, y - 4, 10, 10
GdipFillEllipse ViewGraphics, White, x + Width - 4, y + Height - 4, 10, 10
GdipFillEllipse ViewGraphics, White, x - 4, y + Height - 4, 10, 10

GdipFillEllipse ViewGraphics, White, x + Width / 2 - 5, y - 4, 10, 10
GdipFillEllipse ViewGraphics, White, x + Width / 2 - 5, y + Height - 4, 10, 10
GdipFillEllipse ViewGraphics, White, x + Width - 4, y + Height / 2 - 5, 10, 10
GdipFillEllipse ViewGraphics, White, x - 4, y + Height / 2 - 5, 10, 10
White = NewPen(&H55808080, 1)
GdipDrawEllipse ViewGraphics, White, x - 4, y - 4, 10, 10
GdipDrawEllipse ViewGraphics, White, x + Width - 4, y - 4, 10, 10
GdipDrawEllipse ViewGraphics, White, x + Width - 4, y + Height - 4, 10, 10
GdipDrawEllipse ViewGraphics, White, x - 4, y + Height - 4, 10, 10

GdipDrawEllipse ViewGraphics, White, x + Width / 2 - 5, y - 4, 10, 10
GdipDrawEllipse ViewGraphics, White, x + Width / 2 - 5, y + Height - 4, 10, 10
GdipDrawEllipse ViewGraphics, White, x + Width - 4, y + Height / 2 - 5, 10, 10
GdipDrawEllipse ViewGraphics, White, x - 4, y + Height / 2 - 5, 10, 10
End Sub

'======================================

Public Sub SetStartPoint(sx As Single, sy As Single)
x1 = sx
y1 = sy
End Sub
Public Sub SetEndPoint(sx As Single, sy As Single)
x2 = sx: y2 = sy
End Sub


