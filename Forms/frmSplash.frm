VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5160
   ClientLeft      =   3240
   ClientTop       =   2970
   ClientWidth     =   8865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   8670
      Begin VB.Label lblIntro 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   8415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
      Unload Me
End Sub

Private Sub Form_Load()

      Dim txt As String

      txt = "Q B a s i c    G O R I L L A S" & Chr$(13)
      txt = txt + "Is Now" & Chr$(13)
      txt = txt + "V i s u a l  B a s i c    G O R I L L A S" & Chr$(13)
      txt = txt + "" & Chr$(13)
      txt = txt + "QBasic Gorillas is Copyright (C) Microsoft Corporation 1990" & Chr$(13)
      txt = txt + "QBasic to Visual Basic Conversion done by Troy Rabel" & Chr$(13)
      txt = txt + "" & Chr$(13)
      txt = txt + "Your mission is to hit your opponent with the exploding" & Chr$(13)
      txt = txt + "banana by varying the angle and power of your throw, taking" & Chr$(13)
      txt = txt + "into account wind speed, gravity, and the city skyline." & Chr$(13)
      txt = txt + "The wind speed is shown by a directional arrow at the" & Chr$(13)
      txt = txt + "bottom of the playing field, its length relative to its strength." & Chr$(13)
      txt = txt + "" & Chr$(13)
      txt = txt + "Press any key to continue" & Chr$(13)
      frmSplash.lblIntro.Caption = txt
      frmSplash.lblIntro.Visible = True

End Sub

Private Sub Frame1_Click()
      Unload Me
End Sub

Private Sub lblIntro_Click()
      Unload Me
End Sub
