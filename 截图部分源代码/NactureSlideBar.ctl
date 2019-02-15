VERSION 5.00
Begin VB.UserControl NactureSlideBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HitBehavior     =   0  '无
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "NactureSlideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private MForecolor As OLE_COLOR, _
            MBackcolor As OLE_COLOR, _
            MValueMax As Long, _
            MValueMin As Long, _
            MValue As Long, _
            MHasPoint As Boolean
Private M_Down As Boolean, _
            MValueChange As Long
             
'****事件***********************************************************************************
'
'
'
'*******************************************************************************************
Public Event Scroll(Value As Long)
Public Event Change(Value As Long)
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
M_Down = True
MValue = x - 2
ReFresh True, MValue
SetCapture UserControl.hwnd
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim V
If M_Down = True Then
    If x - 2 < 1 Then
       MValue = 0
    ElseIf x - 2 > UserControl.Width \ 15 - 24 Then
       MValue = UserControl.Width \ 15 - 24
    Else
       MValue = x - 2
    End If
    ReFresh True, MValue
    RaiseEvent Scroll(MValue)
End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
M_Down = False
ReFresh False, MValue
ReleaseCapture
RaiseEvent Change(MValue)
End Sub

'****属性***********************************************************************************
'
'
'
'*******************************************************************************************

'背景颜色
Public Property Get Backcolor() As OLE_COLOR
Attribute Backcolor.VB_ProcData.VB_Invoke_Property = ";外观"
Backcolor = MBackcolor
End Property
Public Property Let Backcolor(newColor As OLE_COLOR)
MBackcolor = newColor
PropertyChanged "Backcolor": ReFresh
End Property
'前景颜色
Public Property Get Forecolor() As OLE_COLOR
Forecolor = MForecolor
End Property
Public Property Let Forecolor(newColor As OLE_COLOR)
MForecolor = newColor
PropertyChanged "Forecolor": ReFresh
End Property
''滑动最小值
'Public Property Get ValueMin() As Long
'ValueMin = MValueMin
'End Property
'Public Property Let ValueMin(NewValue As Long)
'MValueMin = NewValue
'MValueChange = (UserControl.Width \ 15 - 24) / (MValueMax - MValueMin)
'PropertyChanged "ValueMin"
'ReFresh False, MValue
'End Property
'滑动最大值
Public Property Get ValueMax() As Long
ValueMax = MValueMax
End Property
Public Property Let ValueMax(NewValue As Long)
MValueMax = NewValue
'MValueChange = (UserControl.Width \ 15 - 24) / (MValueMax - MValueMin)
PropertyChanged "ValueMax"
UserControl.Width = (ValueMax + 24) * 15
If MValue > MValueMax Then MValue = MValueMax
ReFresh False, MValue
End Property
'默认值
Public Property Get Value() As Long
Value = MValue
End Property
Public Property Let Value(NewValue As Long)
MValue = NewValue
PropertyChanged "Value"
If MValue > MValueMax Then MValue = MValueMax
ReFresh False, MValue
End Property
''是否绘制分界点
'Public Property Get HasPoint() As Boolean
'HasPoint = MHasPoint
'End Property
'Public Property Let HasPoint(NewPoint As Boolean)
'MHasPoint = NewPoint
'PropertyChanged "HasPoint": ReFresh False, MValue
'End Property
'****属性读取*******************************************************************************
'
'
'
'****************************************************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
    MBackcolor = .ReadProperty("Backcolor", 0)
    MForecolor = .ReadProperty("Forecolor", 0)
    Let MValueMax = .ReadProperty("ValueMax", 100)
'    Let MValueMin = .ReadProperty("ValueMin", 0)
    Let MValue = .ReadProperty("Value", 0)
'    Let MHasPoint = .ReadProperty("HasPoint", True)
 End With
'MValueChange = (UserControl.Width \ 15 - 24) / (MValueMax - MValueMin)
ReFresh False, MValue
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
    .WriteProperty "Backcolor", MBackcolor, 0
    .WriteProperty "Forecolor", MForecolor, 0
    .WriteProperty "Value", MValue, 0
    .WriteProperty "ValueMax", MValueMax, 0
'    .WriteProperty "ValueMin", MValueMin, 0
'    .WriteProperty "HasPoint", False
 End With
End Sub
'****绘图*******************************************************************************
'
'
'
'****************************************************************************************
Public Sub ReFresh(Optional Click As Boolean = False, Optional Value = 0)
Dim Token As Long, _
        Inputbuf As GdiplusStartupInput, _
        graphics As Long, _
        BackBrush As Long, _
        ForeBrush As Long, _
        SidePen2 As Long, _
        AForeBrush As Long
Dim i
        
        Inputbuf.GdiplusVersion = 1
        GdiplusStartup Token, Inputbuf
        GdipCreateFromHDC UserControl.Hdc, graphics
        GdipCreatePen1 OLEColorChange(MForecolor), 2, UnitPixel, SidePen2
        GdipCreateSolidFill OLEColorChange(MBackcolor), BackBrush
        GdipCreateSolidFill OLEColorChange(Forecolor), ForeBrush
        GdipFillRectangle graphics, BackBrush, 0, 0, UserControl.Width \ 15, UserControl.Height \ 15
               
    
        GdipFillRectangle graphics, ForeBrush, 12, 12, Value - 2, 4 '左横条
        GdipDrawLine graphics, SidePen2, Value + 14, 14, UserControl.Width \ 15 - 12, 14
        
'        If HasPoint = True And MValueChange <> 0 Then
'           Do Until i > (UserControl.Width \ 15 - 24) \ MValueChange
'           'DoEvents
'              GdipFillRectangle Graphics, ForeBrush, i * MValueChange + 10, 12, 4, 4
'              i = i + 1
'           Loop
'        End If
        
        GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
        GdipFillEllipse graphics, BackBrush, Value - 3, 7, 12, 12 '遮盖
        
        If Click = True Then
           GdipFillEllipse graphics, ForeBrush, Value, 8, 10, 10
        Else
           GdipDrawEllipse graphics, SidePen2, Value, 8, 10, 10
        End If
         
         UserControl.ReFresh
         
         GdipDeletePen SidePen2
         GdipDeleteBrush BackBrush
         GdipDeleteBrush ForeBrush
         GdipDeleteGraphics graphics
         GdiplusShutdown Token
         
End Sub
'****杂项*******************************************************************************
'
'
'
'****************************************************************************************

Private Sub UserControl_Initialize()
MBackcolor = RGB(255, 255, 255)
MForecolor = RGB(27, 187, 205)
MValueMax = 100
ReFresh
End Sub


Private Sub UserControl_Resize()
UserControl.Height = 28 * 15
UserControl.Width = (ValueMax + 24) * 15
If MValue > MValueMax Then MValue = MValueMax
ReFresh False, MValue
End Sub
Public Function OLEColorChange(Color As OLE_COLOR) As Long
Dim C, i, Ccount
C = Hex(Color)
If Len(C) < 6 Then
    Ccount = 6 - Len(C)
    For i = 1 To Ccount
     C = "0" & C
    Next
End If
OLEColorChange = "&hff" & Mid(C, 5, 2) & Mid(C, 3, 2) & Mid(C, 1, 2)
End Function

