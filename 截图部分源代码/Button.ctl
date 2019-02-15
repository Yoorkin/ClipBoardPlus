VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mFont As Font, mCaption As String, mFontColor As OLE_COLOR, mForecolor As OLE_COLOR, mBackColor As OLE_COLOR
Private mFocus As Boolean, mPress As Boolean, mEnable As Boolean, mMouseIn As Boolean
Private SFontColor As OLE_COLOR, SBackColor As OLE_COLOR
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long



Private Sub UserControl_InitProperties()
Set mFont = Ambient.Font
mFont.Size = 16
mBackColor = RGB(225, 225, 225)
mForecolor = RGB(0, 121, 215)
mFontColor = &H606060
mCaption = Extender.name
Enable = False
Enable = True
End Sub







Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
With PropBag
 Set mFont = .ReadProperty("Font")
 mBackColor = .ReadProperty("Backcolor")
 mFontColor = .ReadProperty("Fontcolor")
 mForecolor = .ReadProperty("Forecolor")
 SFontColor = .ReadProperty("sFontcolor")
 SBackColor = .ReadProperty("sBackcolor")
 mCaption = .ReadProperty("Caption")
 mEnable = .ReadProperty("Enable")
End With
Call UserControl_Paint
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "Font", mFont, Ambient.Font
 .WriteProperty "Backcolor", mBackColor, 0
 .WriteProperty "Forecolor", mForecolor, 0
 .WriteProperty "Fontcolor", mFontColor, 0
 .WriteProperty "sBackcolor", SBackColor, 0
 .WriteProperty "sFontcolor", SFontColor, 0
 .WriteProperty "Caption", mCaption, 0
 .WriteProperty "Enable", mEnable
End With
End Sub

Public Property Let Enable(E As Boolean)
If Not E Then
   SFontColor = mFontColor: mFontColor = RGB(131, 131, 131)
   SBackColor = mBackColor: mBackColor = RGB(210, 210, 210)
Else
   mFontColor = SFontColor
   mBackColor = SBackColor
End If
mEnable = E
PropertyChanged "Enable"
PropertyChanged "sBackcolor"
PropertyChanged "sFontcolor"
PropertyChanged "Enable"
Call UserControl_Paint
End Property
Public Property Get Enable() As Boolean
Enable = mEnable
End Property
Public Property Set Font(F As Font)
Set mFont = F
End Property
Public Property Get Font() As Font
Set Font = mFont
PropertyChanged "Font"
Call UserControl_Paint
End Property
Public Property Get Caption() As String
Caption = mCaption
End Property
Public Property Let Caption(NewCaption As String)
mCaption = NewCaption
PropertyChanged "Caption"
Call UserControl_Paint
End Property
Public Property Get Fontcolor() As OLE_COLOR
Fontcolor = mFontColor
End Property
Public Property Let Fontcolor(newColor As OLE_COLOR)
mFontColor = newColor
PropertyChanged "Fontcolor"
Call UserControl_Paint
End Property
Public Property Get Forecolor() As OLE_COLOR
Forecolor = mForecolor
End Property
Public Property Let Forecolor(newColor As OLE_COLOR)
mForecolor = newColor
PropertyChanged "Forecolor"
Call UserControl_Paint
End Property
Public Property Get Backcolor() As OLE_COLOR
Backcolor = mBackColor
End Property
Public Property Let Backcolor(newColor As OLE_COLOR)
mBackColor = newColor
PropertyChanged "Backcolor"
Call UserControl_Paint
End Property


Private Sub UserControl_Paint()
Dim mToken  As Long, Inputbuf As GdiplusStartupInput, _
        Graphics As Long, ButtonDeeph As Byte
Dim Fontfam As Long, _
       Strformat As Long, _
       MyFont As Long, _
       Rclayout As RECTF, _
       RECT As RECTF
    
    Inputbuf.GdiplusVersion = 1
    GdiplusStartup mToken, Inputbuf
    GdipCreateFromHDC UserControl.hDC, Graphics '»­²¼


        With Rclayout
            .Left = 0
            .Right = UserControl.Width \ Screen.TwipsPerPixelX
            .Top = 0
            .Bottom = UserControl.Height \ Screen.TwipsPerPixelY
        End With
    GdipGraphicsClear Graphics, OLEColorChange(mBackColor)
    
    Dim DashPen As Long
    GdipCreatePen1 &H40000000, 1, UnitPixel, DashPen
    
      
      If mMouseIn Then
       ButtonDeeph = 60
       Debug.Print Time & "MPress====" & ButtonDeeph
       GdipSetPenColor DashPen, OLEColorChange(mForecolor)
       GdipDrawRectangleI Graphics, DashPen, 0, 0, Rclayout.Right - 1, Rclayout.Bottom - 1
       GdipFillRectangleI Graphics, NewBrush(OLEColorChange(mForecolor, ButtonDeeph)), 0, 0, Rclayout.Right - 1, Rclayout.Bottom - 1
      End If
      If mPress Then GdipFillRectangleI Graphics, NewBrush(OLEColorChange(&H808080, ButtonDeeph)), 0, 0, Rclayout.Right - 1, Rclayout.Bottom - 1
  
      GdipDrawRectangleI Graphics, DashPen, Rclayout.Left, Rclayout.Top, Rclayout.Right - 1, Rclayout.Bottom - 1
      
      If mFocus Then
        GdipSetPenWidth DashPen, 1
        GdipSetPenColor DashPen, &HF0000000
        GdipSetPenDashStyle DashPen, DashStyleDot
        GdipDrawRectangleI Graphics, DashPen, 2, 2, Rclayout.Right - 5, Rclayout.Bottom - 5
      End If
   
   If Not mFont Is Nothing Then
            GdipCreateFontFamilyFromName StrPtr(mFont.name), 0, Fontfam
            GdipCreateFont Fontfam, mFont.Size, FontStyleRegular, UnitPixel, MyFont
            GdipCreateStringFormat 0, 0, Strformat
            GdipSetStringFormatAlign Strformat, StringAlignmentCenter
            GdipMeasureString Graphics, StrPtr(mCaption), Len(mCaption), MyFont, Rclayout, Strformat, RECT, 0, 0
            RECT.Top = (Rclayout.Bottom - RECT.Bottom) / 2
            GdipDrawString Graphics, StrPtr(mCaption), Len(mCaption), MyFont, RECT, Strformat, NewBrush(OLEColorChange(mFontColor))
    End If
    UserControl.Refresh

    GdipDeleteGraphics Graphics
    GdipDeleteFontFamily Fontfam
    GdipDeleteStringFormat Strformat
    GdipDeleteFont MyFont
    GdiplusShutdown mToken
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print Time & "KeyDown"
If KeyCode = 32 Then
mPress = True
Call UserControl_MouseMove(0, 0, 0, 0)
Debug.Print Time & "HighLight"
Call UserControl_MouseDown(0, 0, 0, 0)
End If
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call UserControl_MouseDown(0, 0, 0, 0)
Call UserControl_MouseUp(0, 0, 0, 0)
End If
Debug.Print Time & "KeyPrress"
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then Call UserControl_MouseUp(0, 0, 0, 0)
Debug.Print Time & "KeyUp"
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not mEnable Then Exit Sub
SetCapture UserControl.hwnd
If mMouseIn = False Then
    mMouseIn = True
    Call UserControl_Paint
ElseIf X > UserControl.Width \ Screen.TwipsPerPixelX Or Y > UserControl.Height \ Screen.TwipsPerPixelY Or X < 0 Or Y < 0 Then
    ReleaseCapture
    mMouseIn = False
    Call UserControl_Paint
End If
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not mEnable Then Exit Sub
mPress = True
Call UserControl_Paint
RaiseEvent MouseDown(Button, Shift, X, Y)
Debug.Print Time & "MouseDown"
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not mEnable Then Exit Sub
mPress = False
Call UserControl_Paint
RaiseEvent MouseUp(Button, Shift, X, Y)
RaiseEvent Click
Debug.Print Time & "MouseUp"
End Sub
Private Sub UserControl_GotFocus()
If Not mEnable Then Exit Sub
mFocus = Not mFocus
Call UserControl_Paint
End Sub
Private Sub UserControl_LostFocus()
If Not mEnable Then Exit Sub
mFocus = Not mFocus
mMouseIn = False
Call UserControl_Paint
End Sub


Private Sub UserControl_Resize()
Call UserControl_Paint
End Sub
Public Function OLEColorChange(Color As OLE_COLOR, Optional ColorAlpha As Byte = 255) As Long
Dim C, i, Ccount
C = Hex(Color)
If Len(C) < 6 Then
    Ccount = 6 - Len(C)
    For i = 1 To Ccount
     C = "0" & C
    Next
End If
OLEColorChange = "&h" & Hex(ColorAlpha) & Mid(C, 5, 2) & Mid(C, 3, 2) & Mid(C, 1, 2)
End Function

