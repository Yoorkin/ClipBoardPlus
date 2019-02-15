VERSION 5.00
Begin VB.Form jpgSave 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   0  'None
   Caption         =   "保存为JPG"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin 闪截.Button Button1 
      Height          =   465
      Left            =   2655
      TabIndex        =   3
      Top             =   1485
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   14803425
      Forecolor       =   14121216
      Fontcolor       =   6316128
      sBackcolor      =   14803425
      sFontcolor      =   6316128
      Caption         =   "保存"
      Enable          =   -1  'True
   End
   Begin 闪截.NactureSlideBar NactureSlideBar 
      Height          =   420
      Left            =   1665
      TabIndex        =   0
      Top             =   630
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
      Backcolor       =   15724527
      Forecolor       =   15571756
      Value           =   50
      ValueMax        =   100
   End
   Begin VB.Label Label 
      BackColor       =   &H00EFEFEF&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ED9B2C&
      Height          =   330
      Left            =   3645
      TabIndex        =   2
      Top             =   720
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00EFEFEF&
      BorderColor     =   &H00ED9B2C&
      Height          =   2085
      Left            =   0
      Top             =   0
      Width           =   4200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EFEFEF&
      Caption         =   "图片质量"
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   630
      TabIndex        =   1
      Top             =   720
      Width           =   825
   End
End
Attribute VB_Name = "jpgSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Img As Long
Private Sub Button1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     SaveImageToJPG Img, CutBoard.CommonDialog1.filename + ".jpg", NactureSlideBar.Value
     Unload CutBoard
     End
End Sub

Private Sub NactureSlideBar_Scroll(Value As Long)
Label.Caption = NactureSlideBar.Value
End Sub
