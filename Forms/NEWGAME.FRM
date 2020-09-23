VERSION 5.00
Begin VB.Form frmNewGame 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "New Game"
   ClientHeight    =   5175
   ClientLeft      =   4215
   ClientTop       =   3585
   ClientWidth     =   5115
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
   ScaleHeight     =   5175
   ScaleWidth      =   5115
   Begin VB.TextBox txtTotalPoints 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2805
      TabIndex        =   1
      Text            =   "3"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtBDiff 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2805
      TabIndex        =   4
      Text            =   "120"
      Top             =   2175
      Width           =   855
   End
   Begin VB.PictureBox Frame3D2 
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   360
      ScaleHeight     =   1275
      ScaleWidth      =   4515
      TabIndex        =   5
      Top             =   240
      Width           =   4575
      Begin VB.ComboBox txtPlayer2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         TabIndex        =   11
         Text            =   "Player 2"
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox txtPlayer1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         TabIndex        =   10
         Text            =   "Player 1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of player 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   735
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of player 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   375
         Width           =   1245
      End
   End
   Begin VB.PictureBox Frame3D1 
      Height          =   1575
      Left            =   360
      ScaleHeight     =   1515
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   2760
      Width           =   4215
      Begin VB.TextBox txtGravity 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   0
         Text            =   "9.8"
         Top             =   472
         Width           =   855
      End
      Begin VB.CommandButton Command3D2 
         Caption         =   "&Earth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   975
         Width           =   1455
      End
      Begin VB.CommandButton Command3D3 
         Caption         =   "&Moon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity in meters/sec (Earth = 9.8):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   495
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command3D1 
      Caption         =   "&Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   4560
      Width           =   4080
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Play to how many total points:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   1815
      Width           =   2100
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Building height difference:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   870
      TabIndex        =   9
      Top             =   2190
      Width           =   1830
   End
End
Attribute VB_Name = "frmNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3D1_Click()
    
      Player1 = txtPlayer1
      Player2 = txtPlayer2
      
      WriteRegistry "Gorillas", "Player1", Player1
      WriteRegistry "Gorillas", "Player2", Player2
      
      NumGames = Val(Left$(txtTotalPoints, 2))
            
      If NumGames = 0 Then
            NumGames = 3
      End If
      
      WriteRegistry "Gorillas", "Total Games", CStr(NumGames)
      
      gravity = Val(txtGravity)
      
      If gravity = 0 Then
            gravity = 9.8
      End If
      
      WriteRegistry "Gorillas", "Gravity", CStr(gravity)
      
      If RandomHeight <= 0 Then
            txtBDiff = "120"
      End If
      
      RandomHeight = Val(txtBDiff)
      
      WriteRegistry "Gorillas", "RandomHeight", CStr(RandomHeight)
      
      Unload Me

End Sub

Private Sub Command3D2_Click()
      txtGravity = "9.8"
End Sub

Private Sub Command3D3_Click()
      txtGravity = "2.6"
End Sub

Private Sub Form_Load()

      CenterForm Me

      txtPlayer1.AddItem "Computer"
      txtPlayer2.AddItem "Computer"
    
      txtPlayer1 = "Player 1"
      If Player1 <> "" Then txtPlayer1 = Player1
      
      txtPlayer2 = "Player 2"
      If Player2 <> "" Then txtPlayer2 = Player2

End Sub
