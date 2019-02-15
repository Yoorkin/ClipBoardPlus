VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public x As Long, y As Long, Width As Long, Height As Long
Private x1 As Single, x2 As Single, y1 As Single, y2 As Single
Public Sub SetStartPoint(sx As Single, sy As Single)
x1 = sx
y1 = sy
End Sub
Public Sub SetEndPoint(sx As Single, sy As Single)
If x1 < sx Then
  x = x1
  Width = sx - x1
Else
  x = sx
  Width = x1 - sx
End If

If y1 < sy Then
  y = y1
  Height = sy - y1
Else
  y = sy
  Height = y1 - sy
End If
End Sub

Public Sub Move(mx As Single, my As Single)
x = mx
y = my
End Sub
