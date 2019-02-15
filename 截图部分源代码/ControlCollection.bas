Attribute VB_Name = "ControlCollection"
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public ShadowGraphics As Long, ViewGraphics As Long
Public LineBrush As Long, ColorBoxGraphics(8) As Long
Public Linecolor(5, 1) As Long, UsingColor As Long

Public ToolBoxGraphics(8) As Long, ToolIcon(8) As Long, WToolIcon(8) As Long, CloseImage As Long
Public Controls() As Object, FocusIndex As Integer
Public Enum PolyType
 ShapeRectangle = 1
 ShapeEllipse = 0
End Enum
Public Enum FocusList
 none
 LeftTop
 RightTop
 LeftDown
 RightDown
 Top
 Botton
 Left
 Right
 Middle
End Enum
 
Public Function XPixel(Twip As Variant)  '缇转化为像素
XPixel = Twip \ Screen.TwipsPerPixelX
End Function
Public Function YPixel(Twip As Variant)
YPixel = Twip \ Screen.TwipsPerPixelY
End Function
Public Function XTwip(Pixel As Variant)  '像素转化为缇
XTwip = Pixel * Screen.TwipsPerPixelX
End Function
Public Function YTwip(Pixel As Variant)
YTwip = Pixel * Screen.TwipsPerPixelY
End Function

