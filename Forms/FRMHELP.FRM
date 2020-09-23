VERSION 5.00
Begin VB.Form frmHelp 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1005
   ClientLeft      =   6810
   ClientTop       =   2715
   ClientWidth     =   5430
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1005
   ScaleWidth      =   5430
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   
      ' Get rid of the forms Caption so title bar does not display:
      Me.Caption = ""

      ' Give the form and label a light yellow background:
      Me.BackColor = &H80FFFF
      Label1.BackColor = &H80FFFF

      ' Inset label by 1 pixel:
      Label1.Left = 1 * Screen.TwipsPerPixelX
      Label1.Top = 0

End Sub
