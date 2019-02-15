VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CutBoard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1adsf"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   854
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ToolBar 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   6435
      ScaleHeight     =   420
      ScaleWidth      =   3660
      TabIndex        =   2
      Top             =   1260
      Visible         =   0   'False
      Width           =   3660
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   7
         Left            =   2430
         TabIndex        =   18
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   1
         Left            =   405
         TabIndex        =   4
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   2
         Left            =   810
         TabIndex        =   6
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   3
         Left            =   1215
         TabIndex        =   7
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   4
         Left            =   1620
         TabIndex        =   8
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   5
         Left            =   2025
         TabIndex        =   5
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   6
         Left            =   2835
         TabIndex        =   9
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin 闪截.Box ToolBox 
         Height          =   420
         Index           =   8
         Left            =   3240
         TabIndex        =   10
         Top             =   0
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin VB.PictureBox ColorSheet 
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   -45
         ScaleHeight     =   645
         ScaleWidth      =   3885
         TabIndex        =   11
         Top             =   360
         Width           =   3885
         Begin 闪截.Box ColorBox 
            Height          =   375
            Index           =   0
            Left            =   2880
            TabIndex        =   12
            Top             =   90
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin 闪截.Box ColorBox 
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   90
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin 闪截.Box ColorBox 
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   14
            Top             =   90
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin 闪截.Box ColorBox 
            Height          =   375
            Index           =   3
            Left            =   2160
            TabIndex        =   15
            Top             =   90
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin 闪截.Box ColorBox 
            Height          =   375
            Index           =   4
            Left            =   2520
            TabIndex        =   16
            Top             =   90
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin VB.Label Label1 
            BackColor       =   &H00EEEEEE&
            Caption         =   "选择颜色"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   360
            TabIndex        =   17
            Top             =   135
            Width           =   780
         End
      End
   End
   Begin VB.PictureBox PicShadow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   2025
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.PictureBox PicGraphics 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   4290
      Left            =   1530
      ScaleHeight     =   286
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3345
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7965
      Top             =   3150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.png|*.bmp|*.jpg"
   End
End
Attribute VB_Name = "CutBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Cut As CutPort, Using As Object
Public Enum EnumControlMode
 CreateMode = 0
 SizeMode = 1
End Enum
Public Enum EnumObject
 OBJCutBoard
 ObjNone
 SolidRectangle
 SolidEllipse
 Rectangle
 Ellipse
End Enum

Public IsMouseDown As Boolean, ControlMode As EnumControlMode, UsingObject As EnumObject '鼠标是否按下/控制模式/要创建的对象
Public BoardLayout As New Layout




Private Sub Form_Load()
UsingColor = &HFF00A8EC
Me.Move 0, 0, Screen.Width, Screen.Height
With PicGraphics
 .Width = Screen.Width / Screen.TwipsPerPixelX
 .Height = Screen.Height / Screen.TwipsPerPixelY
End With
With PicShadow
 .Width = Screen.Width / Screen.TwipsPerPixelX
 .Height = Screen.Height / Screen.TwipsPerPixelY
End With

InitGDIPlus
Linecolor(0, 0) = &HFF00A8EC
Linecolor(0, 1) = &HFF0064EC
Linecolor(1, 0) = &HFF00B034
Linecolor(1, 1) = &HFF006A25
Linecolor(2, 0) = &HFF832197
Linecolor(2, 1) = &HFF3D107B
Linecolor(3, 0) = &HFFDF0024
Linecolor(3, 1) = &HFF960014
Linecolor(4, 0) = &HFFE8641B
Linecolor(4, 1) = &HFFA04716

Dim i As Integer
For i = 0 To ToolBox.UBound
  GdipCreateFromHDC ToolBox(i).Hdc, ToolBoxGraphics(i)
Next
For i = 0 To 8
  GdipLoadImageFromFile StrPtr(App.Path + "\Icon\" + Replace(str(i), " ", "") + ".png"), ToolIcon(i)
  GdipLoadImageFromFile StrPtr(App.Path + "\Icon\W" + Replace(str(i), " ", "") + ".png"), WToolIcon(i)
  GdipSetInterpolationMode ToolBoxGraphics(i), InterpolationModeBilinear
  ToolBox_Mouse i, BoxExit, 0, 0
  Debug.Print App.Path + "\Icon\W" + str(i) + ".png"
Next
For i = 0 To 4
  GdipCreateFromHDC ColorBox(i).Hdc, ColorBoxGraphics(i)
  GdipSetSmoothingMode ColorBoxGraphics(i), SmoothingModeAntiAlias
Next
GdipLoadImageFromFile StrPtr(App.Path + "\Icon\Close.png"), CloseImage
Call Start
End Sub

Public Sub Start()
Dim ScreenDc As Long
ScreenDc = GetDC(0)
BitBlt PicGraphics.Hdc, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, ScreenDc, 0, 0, vbSrcCopy
BitBlt PicShadow.Hdc, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, ScreenDc, 0, 0, vbSrcCopy
GdipCreateFromHDC Me.Hdc, ViewGraphics
GdipCreateFromHDC PicShadow.Hdc, ShadowGraphics
GdipFillRectangle ShadowGraphics, NewBrush(&HAA000000), 0, 0, Screen.Width / Screen.TwipsPerPixelY, Screen.Height / Screen.TwipsPerPixelY
GdipSetSmoothingMode ViewGraphics, SmoothingModeAntiAlias
GdipCreateLineBrush NewPointF(0, 0), NewPointF(XPixel(ToolBox(0).Width), XPixel(ToolBox(0).Height)), &HFF00A8EC, &HFF0064EC, WrapModeTileFlipX, LineBrush
Draw
Me.Show
End Sub

Public Sub Draw()
BitBlt Me.Hdc, 0, 0, Screen.Width * Screen.TwipsPerPixelX, Screen.Height * Screen.TwipsPerPixelY, PicShadow.Hdc, 0, 0, vbSrcCopy
If Not Cut Is Nothing Then
BitBlt Me.Hdc, Cut.X, Cut.Y, Cut.Width, Cut.Height, PicGraphics.Hdc, Cut.X, Cut.Y, vbSrcCopy
BoardLayout.ReFresh
End If
Me.ReFresh
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.ToolBar.Visible = True
Dim CatchObject  As Boolean
CatchObject = BoardLayout.SendMouseDown(X, Y)
If ControlMode = CreateMode And Not CatchObject Then  '创建模式
  Select Case UsingObject
     Case Is = OBJCutBoard
       Set Cut = BoardLayout.CreateCutPort(X, Y)
     Case Is = SolidRectangle
       Set Using = BoardLayout.CreatePolygon(X, Y, True, PolyType.ShapeRectangle)
     Case Is = SolidEllipse
       Set Using = BoardLayout.CreatePolygon(X, Y, True, PolyType.ShapeEllipse)
     Case Is = Rectangle
       Set Using = BoardLayout.CreatePolygon(X, Y, False, PolyType.ShapeRectangle)
     Case Is = Ellipse
       Set Using = BoardLayout.CreatePolygon(X, Y, False, PolyType.ShapeEllipse)
  End Select
ElseIf ControlMode = SizeMode Then '调整模式
  
  
  

End If
IsMouseDown = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BoardLayout.SendMouseMove X, Y
  If ControlMode = CreateMode Then
    If IsMouseDown Then
      Select Case UsingObject
       Case Is = EnumObject.OBJCutBoard
         Cut.SetEndPoint X, Y: Draw: ToolBar.Move Cut.X + Cut.Width + 20, Cut.Y + Cut.Height + 20
       Case Else
         If Not Using Is Nothing Then Using.SetEndPoint X, Y: Draw
      End Select
    End If
  ElseIf ControlMode = SizeMode Then

  End If

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
BoardLayout.SendMouseUp X, Y
IsMouseDown = False
Set Using = Nothing
If UsingObject = OBJCutBoard Then
ControlMode = SizeMode
UsingObject = ObjNone
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
GdipDeleteGraphics ViewGraphics
GdipDeleteGraphics ShadowGraphics
Dim i
For i = 0 To ToolBox.UBound
 GdipDeleteGraphics ToolBoxGraphics(i)
Next
For i = 0 To UBound(ToolIcon)
 GdipDisposeImage ToolIcon(i)
 GdipDisposeImage WToolIcon(i)
Next
     GdipDeleteBrush LineBrush
     TerminateGDIPlus
End Sub


Private Sub ColorBox_Mouse(Index As Integer, MouseEvent As BoxEvent, X As Variant, Y As Variant)
 GdipGraphicsClear ColorBoxGraphics(Index), &HFFEEEEEE
 GdipFillEllipse ColorBoxGraphics(Index), NewBrush(Linecolor(Index, 0)), 5, 2, 15, 15

If MouseEvent = BoxDown Then
  UsingColor = Linecolor(Index, 0)
  GdipDeleteBrush LineBrush
  GdipCreateLineBrush NewPointF(0, 0), NewPointF(30, 30), Linecolor(Index, 0), Linecolor(Index, 1), WrapModeTileFlipX, LineBrush
End If
If ColorBox(Index).IsMouseDown Then
 GdipFillRectangle ColorBoxGraphics(Index), NewBrush(&H22000000), 0, 0, 30, 30
End If
ColorBox(Index).ReFresh
End Sub


Private Sub ToolBox_Mouse(Index As Integer, MouseEvent As BoxEvent, X As Variant, Y As Variant)
Dim SaveImg As Long
With ToolBox(Index)
 If .IsMouseIn Then
   GdipFillRectangleI ToolBoxGraphics(Index), NewBrush(&HFFE0E0E0), 0, 0, ToolBox(Index).Width, ToolBox(Index).Height
   If .IsMouseDown Then GdipFillRectangleI ToolBoxGraphics(Index), NewBrush(&H22000000), 0, 0, ToolBox(Index).Width, ToolBox(Index).Height
 Else
   GdipFillRectangleI ToolBoxGraphics(Index), NewBrush(&HFFFFFFFF), 0, 0, ToolBox(Index).Width, ToolBox(Index).Height
 End If
 If .HasFocus Then
     GdipFillRectangleI ToolBoxGraphics(Index), LineBrush, 0, 0, XPixel(ToolBox(Index).Width), YPixel(ToolBox(Index).Height)
     GdipDrawImageRect ToolBoxGraphics(Index), WToolIcon(Index), 4, 4, XPixel(.Width) - 8, YPixel(.Height) - 8
 Else
     GdipDrawImageRect ToolBoxGraphics(Index), ToolIcon(Index), 4, 4, XPixel(.Width) - 8, YPixel(.Height) - 8
 End If
End With
If MouseEvent = BoxUp Then
Select Case Index
  Case Is = 0
    ControlMode = SizeMode
  Case Is = 1
    ToolBar.Height = 48
  Case Is = 2
   UsingObject = EnumObject.SolidRectangle
   ControlMode = CreateMode
  Case Is = 3
   UsingObject = EnumObject.SolidEllipse
   ControlMode = CreateMode
  Case Is = 4
   UsingObject = EnumObject.Rectangle
   ControlMode = CreateMode
  Case Is = 5
   UsingObject = Ellipse
   ControlMode = CreateMode
  Case Is = 6
   Unload CutBoard
   End
  Case Is = 8

   Me.WindowState = 0
   GdipDeleteGraphics ViewGraphics
   Me.Move 0, 0, XTwip(Cut.Width), YTwip(Cut.Height)
   Me.Cls
   GdipCreateFromHDC Me.Hdc, ViewGraphics
   GdipSetSmoothingMode ViewGraphics, SmoothingModeAntiAlias
   BitBlt Me.Hdc, 0, 0, Cut.Width, Cut.Height, PicGraphics.Hdc, Cut.X, Cut.Y, vbSrcCopy
   BoardLayout.ReFreshTo ViewGraphics, Cut
   Me.ReFresh
   GdipCreateBitmapFromHBITMAP Me.Image.Handle, 0, SaveImg
   SaveImageToPNG SaveImg, App.Path + "\Tem.png"
   Unload CutBoard
   End
  Case Is = 7
    With CommonDialog1
     .DialogTitle = "保存到"
     .InitDir = App.Path
     .Filter = "*.Png|a|*.bmp|a|*.jpg|a|*.GIF"
     .ShowSave
   End With
     
   Me.WindowState = 0
   GdipDeleteGraphics ViewGraphics
   Me.Move 0, 0, XTwip(Cut.Width), YTwip(Cut.Height)
   Me.Cls
   GdipCreateFromHDC Me.Hdc, ViewGraphics
   GdipSetSmoothingMode ViewGraphics, SmoothingModeAntiAlias
   BitBlt Me.Hdc, 0, 0, Cut.Width, Cut.Height, PicGraphics.Hdc, Cut.X, Cut.Y, vbSrcCopy
   BoardLayout.ReFreshTo ViewGraphics, Cut
   Me.ReFresh
   GdipCreateBitmapFromHBITMAP Me.Image.Handle, 0, SaveImg
   Select Case CommonDialog1.FilterIndex
    Case Is = 1
     SaveImageToPNG SaveImg, CommonDialog1.filename + ".png"
    Case Is = 2
     SaveImageToBMP SaveImg, CommonDialog1.filename + ".bmp"
    Case Is = 3
     Me.Hide
     jpgSave.Show
     jpgSave.Img = SaveImg
     Exit Sub
    Case Is = 4
     SaveImageToGIF SaveImg, CommonDialog1.filename + ".GIF"
   End Select
   Debug.Print CommonDialog1.FilterIndex
   Unload CutBoard
   End
End Select
End If
ToolBox(Index).ReFresh
End Sub
