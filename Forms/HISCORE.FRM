VERSION 5.00
Begin VB.Form frmHiScores 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6225
   ClientLeft      =   2475
   ClientTop       =   2925
   ClientWidth     =   10830
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "HISCORE.frx":0000
   ScaleHeight     =   6225
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   9750
      Top             =   4905
   End
   Begin VB.PictureBox Panel3D1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   600
      ScaleHeight     =   4785
      ScaleWidth      =   9585
      TabIndex        =   0
      Top             =   480
      Width           =   9615
      Begin VB.CommandButton cmd_Clear 
         Caption         =   "&Clear All"
         Height          =   510
         Left            =   7155
         TabIndex        =   42
         Top             =   375
         Width           =   2055
      End
      Begin VB.CommandButton cmd_Done 
         Caption         =   "&Done"
         Height          =   510
         Left            =   7155
         TabIndex        =   43
         Top             =   1230
         Width           =   2055
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   0
         Left            =   4230
         TabIndex        =   41
         Top             =   75
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   4230
         TabIndex        =   40
         Top             =   555
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   2
         Left            =   4230
         TabIndex        =   39
         Top             =   1035
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   3
         Left            =   4230
         TabIndex        =   38
         Top             =   1515
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   4
         Left            =   4230
         TabIndex        =   37
         Top             =   1995
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   5
         Left            =   4230
         TabIndex        =   36
         Top             =   2475
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   6
         Left            =   4230
         TabIndex        =   35
         Top             =   2955
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   7
         Left            =   4230
         TabIndex        =   34
         Top             =   3435
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   8
         Left            =   4230
         TabIndex        =   33
         Top             =   3900
         Width           =   2040
      End
      Begin VB.Label pnlDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   9
         Left            =   4230
         TabIndex        =   32
         Top             =   4380
         Width           =   2040
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   9
         Left            =   2895
         TabIndex        =   31
         Top             =   4380
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   8
         Left            =   2895
         TabIndex        =   30
         Top             =   3900
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   7
         Left            =   2895
         TabIndex        =   29
         Top             =   3435
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   6
         Left            =   2895
         TabIndex        =   28
         Top             =   2955
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   5
         Left            =   2895
         TabIndex        =   27
         Top             =   2475
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   4
         Left            =   2895
         TabIndex        =   26
         Top             =   1995
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   3
         Left            =   2895
         TabIndex        =   25
         Top             =   1515
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   2
         Left            =   2895
         TabIndex        =   24
         Top             =   1035
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   2895
         TabIndex        =   23
         Top             =   555
         Width           =   1110
      End
      Begin VB.Label pnlScore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   0
         Left            =   2895
         TabIndex        =   22
         Top             =   75
         Width           =   1110
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   9
         Left            =   540
         TabIndex        =   21
         Top             =   4380
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   8
         Left            =   525
         TabIndex        =   20
         Top             =   3900
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   7
         Left            =   540
         TabIndex        =   19
         Top             =   3435
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   6
         Left            =   525
         TabIndex        =   18
         Top             =   2955
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   5
         Left            =   540
         TabIndex        =   17
         Top             =   2475
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   4
         Left            =   525
         TabIndex        =   16
         Top             =   1995
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   3
         Left            =   540
         TabIndex        =   15
         Top             =   1515
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   2
         Left            =   525
         TabIndex        =   14
         Top             =   1035
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   13
         Top             =   555
         Width           =   2040
      End
      Begin VB.Label pnlPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   0
         Left            =   525
         TabIndex        =   11
         Top             =   75
         Width           =   2040
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " 1."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   4
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   5
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   6
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "8."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   8
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "9."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   9
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "10."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   10
         Top             =   4440
         Width           =   375
      End
   End
   Begin VB.PictureBox picLights2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5850
      Left            =   0
      Picture         =   "HISCORE.frx":CC89A
      ScaleHeight     =   5850
      ScaleWidth      =   10725
      TabIndex        =   12
      Top             =   0
      Width           =   10725
   End
End
Attribute VB_Name = "frmHiScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim x As Integer

Private Sub cmd_Clear_Click()

      On Error Resume Next
      
      frmGorilla.datHiScores.Recordset.MoveLast
      frmGorilla.datHiScores.Refresh
      
      Do While Not frmGorilla.datHiScores.Recordset.EOF
            frmGorilla.datHiScores.Recordset.Delete
            frmGorilla.datHiScores.Recordset.MoveLast
      Loop
      
      frmGorilla.datHiScores.Recordset.MoveFirst
      
      For x = 0 To 9
            frmHiScores.pnlPlayerName(x) = ""
            frmHiScores.pnlScore(x) = ""
            frmHiScores.pnlDate(x) = ""
      Next x

End Sub

Private Sub cmd_Done_Click()
      Unload Me
End Sub

Private Sub Form_Load()
      
      CenterForm Me
      
      On Error Resume Next
      
      frmGorilla.datHiScores.Recordset.MoveFirst
      
      Do While Not frmGorilla.datHiScores.Recordset.EOF
            pnlPlayerName(o) = frmGorilla.datHiScores.Recordset!PlayerName
            pnlScore(o) = frmGorilla.datHiScores.Recordset!Games
            pnlDate(o) = frmGorilla.datHiScores.Recordset!Date
            frmGorilla.datHiScores.Recordset.MoveNext
            o = o + 1
      Loop
      
      'The hi scores are sorted by Games (score)
      'If there are duplicate scores then sort the duplicates by date putting the newest date above the others

End Sub

Private Sub Timer1_Timer()

      If x = 1 Then picLights2.Visible = True
      If x = 2 Then picLights2.Visible = False
      x = x + 1
      If x > 2 Then x = 1

End Sub
