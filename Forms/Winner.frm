VERSION 5.00
Begin VB.Form frmWinner 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "And the winner is..."
   ClientHeight    =   2925
   ClientLeft      =   1545
   ClientTop       =   3525
   ClientWidth     =   7770
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
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Winner.frx":0000
   ScaleHeight     =   2925
   ScaleWidth      =   7770
   Begin VB.Timer timerWinner 
      Interval        =   200
      Left            =   825
      Top             =   3495
   End
   Begin VB.CommandButton btnWinner 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   615
      TabIndex        =   0
      Top             =   495
      Width           =   6720
   End
   Begin VB.Image picWinner 
      Appearance      =   0  'Flat
      Height          =   2925
      Left            =   0
      Picture         =   "Winner.frx":4A48A
      Top             =   0
      Width           =   7800
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PicNum As Integer

Private Sub btnWinner_Click()
      Unload Me
End Sub

Private Sub Form_Load()
      CenterForm Me
End Sub

Private Sub timerWinner_Timer()
      
      If PicNum = 0 Then picWinner.Visible = True
      If PicNum = 1 Then picWinner.Visible = False
      
      PicNum = PicNum + 1
      
      If PicNum > 1 Then
            PicNum = 0
      End If
      
End Sub

