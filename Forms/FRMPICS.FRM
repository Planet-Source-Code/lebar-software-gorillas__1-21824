VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmPics 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   10515
   ClientLeft      =   270
   ClientTop       =   420
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10515
   ScaleWidth      =   15000
   Begin PicClip.PictureClip SunClip 
      Left            =   13320
      Top             =   3000
      _ExtentX        =   2170
      _ExtentY        =   820
      _Version        =   393216
      Cols            =   2
      Picture         =   "FRMPICS.frx":0000
   End
   Begin PicClip.PictureClip BossExplodeClip 
      Left            =   120
      Top             =   5520
      _ExtentX        =   11033
      _ExtentY        =   10636
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      Picture         =   "FRMPICS.frx":1E5A
   End
   Begin PicClip.PictureClip BossClip1 
      Left            =   11880
      Top             =   240
      _ExtentX        =   5186
      _ExtentY        =   1879
      _Version        =   393216
      Cols            =   3
      Picture         =   "FRMPICS.frx":7CCB4
   End
   Begin PicClip.PictureClip BossClip2 
      Left            =   11880
      Top             =   1560
      _ExtentX        =   5186
      _ExtentY        =   1879
      _Version        =   393216
      Cols            =   3
      Picture         =   "FRMPICS.frx":8701A
   End
   Begin PicClip.PictureClip CloudClip 
      Left            =   120
      Top             =   120
      _ExtentX        =   20003
      _ExtentY        =   9525
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      Picture         =   "FRMPICS.frx":91380
   End
   Begin VB.Image GExplosion 
      Appearance      =   0  'Flat
      Height          =   900
      Index           =   0
      Left            =   8160
      Picture         =   "FRMPICS.frx":158932
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image GExplosion 
      Appearance      =   0  'Flat
      Height          =   1080
      Index           =   1
      Left            =   9255
      Picture         =   "FRMPICS.frx":15BD04
      Top             =   6240
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image GExplosion 
      Appearance      =   0  'Flat
      Height          =   1290
      Index           =   2
      Left            =   10560
      Picture         =   "FRMPICS.frx":160786
      Top             =   6120
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image BExplosion 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   0
      Left            =   9000
      Picture         =   "FRMPICS.frx":167098
      Top             =   9960
      Width           =   495
   End
   Begin VB.Image BExplosion 
      Appearance      =   0  'Flat
      Height          =   540
      Index           =   1
      Left            =   9720
      Picture         =   "FRMPICS.frx":167C92
      Top             =   9840
      Width           =   585
   End
   Begin VB.Image BExplosion 
      Appearance      =   0  'Flat
      Height          =   705
      Index           =   2
      Left            =   10440
      Picture         =   "FRMPICS.frx":168DB4
      Top             =   9720
      Width           =   780
   End
   Begin VB.Image BExplosion 
      Appearance      =   0  'Flat
      Height          =   780
      Index           =   3
      Left            =   11400
      Picture         =   "FRMPICS.frx":16AA9A
      Top             =   9600
      Width           =   870
   End
   Begin VB.Image BExplosion 
      Appearance      =   0  'Flat
      Height          =   900
      Index           =   4
      Left            =   12840
      Picture         =   "FRMPICS.frx":16CE9C
      Top             =   6360
      Width           =   1035
   End
   Begin VB.Image BExplosion 
      Appearance      =   0  'Flat
      Height          =   960
      Index           =   5
      Left            =   10680
      Picture         =   "FRMPICS.frx":16FF9E
      Top             =   8160
      Width           =   1065
   End
End
Attribute VB_Name = "frmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
