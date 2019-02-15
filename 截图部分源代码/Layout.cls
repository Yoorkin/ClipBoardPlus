VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Layout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum MouseEvent
 MouseDown
 MouseUp
 MouseMove
 MouseEnter
 MouseExit
End Enum


'Public Sub ReceiveMouseEvent(EventType As MouseEvent, Ex As Single, Ey As Single, Button As EnumButton)
'Select Case EventType
'  Case Is = MouseEnter
'
'  Case Is = MouseExit
'
'  Case Is = MouseMove
'
'End Select
'End Sub
Private IsMouseDown As Boolean
Public Function SendMouseDown(x As Single, y As Single) As Boolean
IsMouseDown = True
Dim i
For i = UBound(Controls) To 1 Step -1
If Controls(i) Is Nothing Then Exit For
With Controls(i)
  If x > .x And y > .y And x < .x + .Width And y < .y + .Height Then
     .ReceiveMouseEvent MouseEvent.MouseDown, x, y, IsMouseDown
     FocusIndex = i
     
     If i = 1 And CutBoard.Cut.FocusObj = Middle Then
       SendMouseDown = False
     Else
       SendMouseDown = True
     End If
     Exit For
  End If
 End With
Next
CutBoard.Draw
End Function
Public Sub SendMouseMove(x As Single, y As Single)
Static IsMouseIn As Boolean
Dim i
For i = UBound(Controls) To 1 Step -1
If Controls(i) Is Nothing Then Exit For
 With Controls(i)
  If (x > .x And y > .y And x < .x + .Width And y < .y + .Height) Or FocusIndex = i Then
     .ReceiveMouseEvent MouseEvent.MouseMove, x, y, IsMouseDown
     If Not .MouseIn Then .ReceiveMouseEvent MouseEvent.MouseEnter, x, y, IsMouseDown: .MouseIn = True
     Exit For
  Else
     If .MouseIn Then .ReceiveMouseEvent MouseEvent.MouseExit, x, y, IsMouseDown: .MouseIn = False
     
  End If
 End With
Next
CutBoard.Draw
End Sub
Public Sub SendMouseUp(x As Single, y As Single)
IsMouseDown = False
Dim i
For i = UBound(Controls) To 1 Step -1
If Controls(i) Is Nothing Then Exit For
 With Controls(i)
  If (x > .x And y > .y And x < .x + .Width And y < .y + .Height) Or FocusIndex = i Then
     .ReceiveMouseEvent MouseEvent.MouseUp, x, y, IsMouseDown
     Exit For
  End If
 End With
Next
FocusIndex = 0
CutBoard.Draw
End Sub
Public Sub ReFreshTo(Graphics As Long, Cut As CutPort)
Dim i
For i = 2 To UBound(Controls)
With Controls(i)
  If .Enable Then
  Select Case .PolygonType
   Case Is = PolyType.ShapeEllipse
    If .Flood = False Then
      GdipDrawEllipse Graphics, NewPen(.Color, 2), .x - Cut.x, .y - Cut.y, .Width, .Height
    Else
      GdipFillEllipse Graphics, NewBrush(.Color), .x - Cut.x, .y - Cut.y, .Width, .Height
    End If
   Case Is = PolyType.ShapeRectangle
    If .Flood = False Then
      GdipDrawRectangle Graphics, NewPen(.Color, 2), .x - Cut.x, .y - Cut.y, .Width, .Height
    Else
      GdipFillRectangle Graphics, NewBrush(.Color), .x - Cut.x, .y - Cut.y, .Width, .Height
    End If
  End Select
  End If
End With
Next
End Sub
Public Sub ReFresh()
Dim i
For i = 1 To UBound(Controls)
  Controls(i).ReFresh
Next
End Sub

Public Function CreateCutPort(StartX As Single, StartY As Single)
ReDim Preserve Controls(UBound(Controls) + 1)
Set Controls(UBound(Controls)) = New CutPort
 Controls(UBound(Controls)).SetStartPoint StartX, StartY
Set CreateCutPort = Controls(UBound(Controls))
End Function
Public Function CreatePolygon(StartX As Single, StartY As Single, IsFlood As Boolean, Shape As PolyType)
ReDim Preserve Controls(UBound(Controls) + 1)
Set Controls(UBound(Controls)) = New Polygon
With Controls(UBound(Controls))
 .SetStartPoint StartX, StartY
 .Flood = IsFlood
 .PolygonType = Shape
 .Color = UsingColor
End With
Set CreatePolygon = Controls(UBound(Controls))
End Function


Private Sub Class_Initialize()
ReDim Preserve Controls(0)
End Sub
