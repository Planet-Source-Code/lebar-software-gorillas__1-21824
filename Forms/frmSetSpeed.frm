VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetSpeed 
   Caption         =   "Set Game Speed"
   ClientHeight    =   1350
   ClientLeft      =   3840
   ClientTop       =   3945
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider Slider1 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   503
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   16
      SelStart        =   5
      Value           =   5
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   2175
      TabIndex        =   1
      Top             =   675
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   420
      Left            =   615
      TabIndex        =   0
      Top             =   675
      Width           =   1365
   End
   Begin VB.Label lbl_MaxSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmSetSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
      MaxSpeed = Slider1.Value * 5
      Unload Me
End Sub

Private Sub Command2_Click()
      Unload Me
End Sub

Private Sub Form_Load()
      Slider1.Value = MaxSpeed \ 5
      lbl_MaxSpeed = MaxSpeed
End Sub

Private Sub Slider1_Scroll()
      MaxSpeed = Slider1.Value * 5
      lbl_MaxSpeed = MaxSpeed
End Sub
