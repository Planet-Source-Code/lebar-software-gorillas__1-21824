VERSION 5.00
Begin VB.Form frmGorilla 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   Caption         =   "Gorillas - 1024 x 780"
   ClientHeight    =   8115
   ClientLeft      =   240
   ClientTop       =   1770
   ClientWidth     =   14925
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
   Icon            =   "Gorilla.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   995
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   7500
      Top             =   9945
   End
   Begin VB.PictureBox pnl_Control 
      Align           =   2  'Align Bottom
      Height          =   1170
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   14865
      TabIndex        =   0
      Top             =   6945
      Width           =   14925
      Begin VB.PictureBox picWindDirection 
         BackColor       =   &H00000000&
         DrawWidth       =   2
         Height          =   225
         Left            =   6120
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   203
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   615
         Width           =   3105
      End
      Begin VB.TextBox txtVelocity2 
         Height          =   285
         Left            =   14010
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   570
         Width           =   1095
      End
      Begin VB.TextBox txtAngle2 
         Height          =   285
         Left            =   12030
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox txtVelocity1 
         Height          =   285
         Left            =   2985
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   570
         Width           =   1095
      End
      Begin VB.TextBox txtAngle1 
         Height          =   285
         Left            =   1005
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   570
         Width           =   975
      End
      Begin VB.Label pnlScore 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   360
         Index           =   1
         Left            =   7920
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.Label pnlScore 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   360
         Index           =   0
         Left            =   6720
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.Label pnl_Player2Name 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8835
         TabIndex        =   7
         Top             =   90
         Width           =   6255
      End
      Begin VB.Label pnl_Player1Name 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   90
         Width           =   6255
      End
      Begin VB.Label lblAngle2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   11400
         TabIndex        =   4
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblVelocity2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Velocity:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13200
         TabIndex        =   3
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblAngle1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblVelocity1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Velocity:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   750
      End
      Begin VB.Line Line1 
         X1              =   7680
         X2              =   7680
         Y1              =   120
         Y2              =   480
      End
   End
   Begin VB.Data datHiScores 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Scores"
      Top             =   10560
      Width           =   2535
   End
   Begin VB.Image picFireBall 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   9480
      Picture         =   "Gorilla.frx":030A
      Top             =   10320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picBuildingExplosion 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   9600
      Picture         =   "Gorilla.frx":068C
      Top             =   10200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image picGorillaExplosion 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   5040
      Picture         =   "Gorilla.frx":1286
      Top             =   9840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbl_GameOver 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   6360
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Image picBananna 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   9840
      Picture         =   "Gorilla.frx":3578
      Top             =   10320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picSun 
      Appearance      =   0  'Flat
      Height          =   465
      Left            =   7470
      Picture         =   "Gorilla.frx":392E
      Top             =   180
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image picBullet 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5040
      Picture         =   "Gorilla.frx":4874
      Top             =   10440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picGorilla 
      Appearance      =   0  'Flat
      Height          =   1125
      Index           =   0
      Left            =   480
      Picture         =   "Gorilla.frx":4BF6
      Stretch         =   -1  'True
      Top             =   8220
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image picGorilla 
      Appearance      =   0  'Flat
      Height          =   1125
      Index           =   1
      Left            =   13920
      Picture         =   "Gorilla.frx":7FF8
      Stretch         =   -1  'True
      Top             =   8235
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_NewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnu_Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnuSpeed 
         Caption         =   "&Game Speed"
      End
      Begin VB.Menu mnu_Sound 
         Caption         =   "&Sound [Off]"
      End
      Begin VB.Menu mnu_Sep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ShowInputs 
         Caption         =   "Show &Inputs"
      End
      Begin VB.Menu mnu_LEDColor 
         Caption         =   "LED &Color"
         Begin VB.Menu mnu_Colors 
            Caption         =   "&Red"
            Index           =   0
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "&Green"
            Index           =   1
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "&Blue"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "&Yellow"
            Index           =   3
         End
      End
      Begin VB.Menu mnu_Sep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_HiScores 
         Caption         =   "&Hi Scores"
      End
      Begin VB.Menu mnuScoreSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowSplash 
         Caption         =   "Show &Splash Screen"
      End
   End
End
Attribute VB_Name = "frmGorilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim y As Integer
Dim rot As Integer

Dim DBName As String
Dim gLIBDB As Database
Dim iCurrentRecord As Integer

Private Sub DoExplosion(x, y)

      Radius = ScrHeight / 30
      Inc# = 0.41
    
      If mnu_Sound.Checked Then
            If Right(App.Path, 1) <> "\" Then
                  PlaySound App.Path & "\sounds\Bhit.wav"
            Else
                  PlaySound App.Path & "sounds\Bhit.wav"
            End If
      End If
    
      DoEvents

      For c# = (Radius + 25) To 0 Step (-1 * Inc#)
            frmGorilla.Circle (x, y), c#, QBColor(9)
      Next c#
    
      For bx = 0 To 5
            picBuildingExplosion.Picture = frmPics.BExplosion(bx).Picture
            picBuildingExplosion.Move x - (picBuildingExplosion.Width \ 2), y - (picBuildingExplosion.Width \ 2)
            
            If Not picBuildingExplosion.Visible Then
                  picBuildingExplosion.Visible = True
            End If
            
            DoEvents
            rest MaxSpeed \ 2
      Next bx
    
      picBuildingExplosion.Visible = False
    
End Sub

Private Function DoShot(PlayerNum, x, y, Angle, Velocity)

      If GameOver Then
            picSun.Visible = False
            pnlWindDirection.Visible = False
            lblAngle1.Visible = False
            txtAngle1.Visible = False
            lblVelocity1.Visible = False
            txtVelocity1.Visible = False
            Exit Function
      End If

      On Error Resume Next

      If PlayerNum = 1 Then
            Angle = 180 - Angle
      End If
      
      SunHit = False
      PlayerHit = PlotShot(x, y - 40, Angle, Velocity, PlayerNum)

End Function

Private Sub EnableCursor(iSetting As Integer)

      Select Case iSetting
            Case True
                  Do While ShowCursor(True) <= 0
                  Loop
            Case False
                  Do While ShowCursor(False) >= 0
                  Loop
      End Select

End Sub

Private Function ExplodeGorilla(x As Integer, y As Integer)
  
   'On Error GoTo DoError:
    
      YAdj = Scl(12)
      XAdj = Scl(5)
      Sclx = ScrWidth / 320
      Scly = ScrHeight / 200
      
      If x < ScrWidth / 2 Then
            PlayerHit = 1
      Else
            PlayerHit = 2
      End If
      
      Radius = ScrHeight / 20
      Inc# = 0.41
    
      If mnu_Sound.Checked Then
            If Right(App.Path, 1) <> "\" Then
                  PlaySound App.Path & "\sounds\ghit.wav"
            Else
                  PlaySound App.Path & "sounds\ghit.wav"
            End If
      End If
      
      DoEvents

      If PlayerHit = 1 Then diff = 41
      If PlayerHit = 2 Then diff = 0
    
      picGorilla(PlayerHit - 1).Visible = True

      For c# = 0 To 4
            picGorilla(PlayerHit - 1).Picture = frmPics.BossExplodeClip.GraphicCell(c#)
            DoEvents
            rest MaxSpeed
      Next c#

      For c# = 0 To 16
            picGorilla(PlayerHit - 1).Picture = frmPics.CloudClip.GraphicCell(c#)
            DoEvents
            rest MaxSpeed
      Next c#

      picGorilla(PlayerHit - 1).Picture = frmPics.BossExplodeClip.GraphicCell(5)
      DoEvents

      picGorillaExplosion.Visible = False
      picGorillaExplosion.Picture = frmPics.GExplosion(0).Picture
      ExplodeGorilla = PlayerHit
      
      Select Case PlayerHit
            Case 1
                  picGorilla(0).Picture = frmPics.BossClip1.GraphicCell(0)
            Case 2
                  picGorilla(1).Picture = frmPics.BossClip2.GraphicCell(0)
      End Select
      
      Exit Function

DoError:
      MsgBox CStr(Err)
      MsgBox Error$
      Resume Next

End Function

Private Sub Command1_Click()
      PlaySound App.Path & "\sounds\Bhit.wav"
End Sub

Private Sub Form_Load()
    
      EnableCursor True
      
      Dim r
      
      '---------------------------------------------------Get Settings -------------------------------------------------------------------
      r = waveOutGetNumDevs()
    
      If r = 0 Then
            mnu_Sound.Enabled = False
            mnu_Sound.Checked = False
      Else
            'Get Sound
            'When you issue a mnu_Sound_Click call then if mnu_Sound.checked = True it will become false
            'and visa versa. So therefore, if we want to set mnu_Sound.checked = True and then use mnu_Sound_Click, we
            'must set mnu_Sound.checked to False first, and then use mnu_Sound_Click and it will become True.
            '
            'In short: Set mnu_Sounds.checked to the opposite of what you really want prior to calling mnu_Sounds_Click
            '
            If ReadRegistry("Gorillas", "Use Sound") <> "" Then
                  mnu_Sound.Checked = Not CBool(ReadRegistry("Gorillas", "Use Sound"))
            Else
                  mnu_Sound.Checked = False
                  CreateRegEntry "Gorillas", "Use Sound", "True"
            End If
            mnu_Sound_Click
      End If
      
      'Get Show Inputs
      If ReadRegistry("Gorillas", "Show Inputs") <> "" Then
            mnu_ShowInputs.Checked = Not CBool(ReadRegistry("Gorillas", "Show Inputs"))
      Else
            mnu_ShowInputs.Checked = False
            CreateRegEntry "Gorillas", "Show Inputs", "True"
      End If
      mnu_ShowInputs_Click
      
      'Get Show Inputs
      If ReadRegistry("Gorillas", "Splash Screen") <> "" Then
            mnuShowSplash.Checked = Not CBool(ReadRegistry("Gorillas", "Splash Screen"))
      Else
            mnuShowSplash.Checked = False
            CreateRegEntry "Gorillas", "Splash Screen", "True"
      End If
      mnuShowSplash_Click
      
      GetSettings
      
      '----------------------------------------------------------------------------------------------------------------------------------------
    
      DBName = App.Path + "\Gorilla.mdb"
      frmGorilla.datHiScores.DatabaseName = DBName
      x% = OpenDB(DBName)
      frmGorilla.datHiScores.RecordSource = "select * from Scores order by Games DESC;"
      frmGorilla.datHiScores.Refresh
    
      Set gDS = frmGorilla.datHiScores.Recordset.Clone()
    
      pi# = 4 * Atn(1#)
      
      Mode = 9

      ScrWidth = Me.ScaleWidth
      ScrHeight = Me.ScaleHeight
      GHeight = 25
    
      LEDColor = 11
    
      ReDim BCoor(0 To 30) As XYPoint

      SetSize

      txtAngle1 = ""
      txtAngle2 = ""
      txtVelocity1 = ""
      txtVelocity2 = ""

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      'If GameInPlay Then 'EnableCursor False
End Sub

Private Sub Form_Resize()
      
      frmGorilla.Refresh

      If GameInPlay Then
            MakeCityScape BCoor(), CityNum, 0, 0, 0
            PlaceGorillas BCoor()
      End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
      'EnableCursor True
      End
End Sub

Private Sub MakeCityScape(BCoor() As XYPoint, Slope As Integer, Color1 As Integer, Color2 As Integer, Color3 As Integer)

      x = 2
    
      Select Case Slope
            Case 1
                  NewHt = 15                 'Upward slope
            Case 2
                  NewHt = 130                'Downward slope
            Case 3 To 5
                  NewHt = 15            '"V" slope - most common
            Case 6
                  NewHt = 130                'Inverted "V" slope
      End Select
  
      BottomLine = Me.ScaleHeight - pnl_Control.Height
      HtInc = 10                         'Increase value for new height
      DefBWidth = 70                     'Default building height
      'RandomHeight = 220                 'Random height difference
      WWidth = 3                         'Window width
      WHeight = 6                        'Window height
      WDifV = 15                         'Counter for window spacing - vertical
      WDifh = 10                         'Counter for window spacing - horizontal

      CurBuilding = 1
      Do
    
            Select Case Slope
                  Case 1
                        NewHt = NewHt + HtInc
                  Case 2
                        NewHt = NewHt - HtInc
                  Case 3 To 5
                        If x > ScrWidth \ 2 Then
                              NewHt = NewHt - 2 * HtInc
                        Else
                              NewHt = NewHt + 2 * HtInc
                        End If
                  Case 4
                        If x > ScrWidth \ 2 Then
                              NewHt = NewHt + 2 * HtInc
                        Else
                              NewHt = NewHt - 2 * HtInc
                        End If
            End Select
        
            'Set width of building and check to see if it would go off the screen
            BWidth = Ran(DefBWidth) + DefBWidth
            If x + BWidth > ScrWidth Then BWidth = ScrWidth - x - 2
        
            'Set height of building and check to see if it goes below screen
            BHeight = Ran(RandomHeight) + NewHt
            If BHeight < HtInc Then BHeight = HtInc
        
            'Check to see if Building is too high
            If BottomLine - BHeight <= MaxHeight + GHeight Then BHeight = MaxHeight + GHeight - 5
        
            'Set the coordinates of the building into the array
            BCoor(CurBuilding).BLeft = x
            BCoor(CurBuilding).XCoor = x
            BCoor(CurBuilding).YCoor = BottomLine - BHeight
            BCoor(CurBuilding).BTop = BottomLine - BHeight
            BCoor(CurBuilding).WCoor = BWidth
            BCoor(CurBuilding).BWidth = BWidth
            BCoor(CurBuilding).BHeight = BHeight
    
            BuildingColor = Ran(3) + 4
            BCoor(CurBuilding).BColor = BuildingColor
        
            'Draw the building, outline first, then filled
            Line (x - 1, BottomLine + 1)-(x + BWidth + 1, BottomLine - BHeight - 1), QBColor(9), B
            Line (x, BottomLine)-(x + BWidth, BottomLine - BHeight), QBColor(BuildingColor), BF
        
            'Draw the windows
            c = x + 3
            Do
                  For i = BHeight - 3 To 7 Step -WDifV
                        If Mode <> 9 Then
                              WinColr = (Ran(2) - 2) * -3
                              WinColr = WinColr
                        ElseIf Ran(4) = 1 Then
                              WinColr = 8
                        Else
                              WinColr = WINDOWCOLOR
                        End If
                        frmGorilla.Line (c, BottomLine - i)-(c + WWidth, BottomLine - i + WHeight), QBColor(WinColr), BF
                  Next i
                  c = c + WDifh
            Loop Until c >= x + BWidth - 3
        
            x = x + BWidth + 2
        
            CurBuilding = CurBuilding + 1
    
      Loop Until x > ScrWidth - HtInc

      LastBuilding = CurBuilding - 1
    
      SetWindSpeed True
    
End Sub

Private Sub mnu_Colors_Click(Index As Integer)

      'Set this to the color of your liking......
     
      mnu_Colors(0).Checked = False
      mnu_Colors(1).Checked = False
      mnu_Colors(2).Checked = False
      mnu_Colors(3).Checked = False
      mnu_Colors(Index).Checked = True
    
      Select Case Index
            Case 0
                  LEDColor = 12
            Case 1
                  LEDColor = 10
            Case 2
                  LEDColor = 11
            Case 3
                  LEDColor = 14
      End Select

      pnlScore(0).ForeColor = QBColor(LEDColor)
      pnlScore(1).ForeColor = QBColor(LEDColor)

      txtVelocity1.ForeColor = QBColor(LEDColor)
      txtVelocity2.ForeColor = QBColor(LEDColor)

      SetWindSpeed False
      
      WriteRegistry "Gorillas", "LED Color", CStr(LEDColor)

End Sub

Private Sub mnu_Exit_Click()
      'EnableCursor True
      End
End Sub

Private Sub mnu_HiScores_Click()
      frmHiScores.Show 1
      'MakeCityScape BCoor(), CityNum, 0, 0, 0
End Sub

Private Sub mnu_NewGame_Click()

      frmNewGame.Show 1
      GameOver = False
      GameInPlay = True
      lbl_GameOver.Visible = False
    
      lblAngle1.Enabled = True
      txtAngle1.Visible = True
      lblVelocity1.Enabled = True
      txtVelocity1.Visible = True
    
      Score0 = 0
      Score1 = 0
      pnlScore(0) = "0"
      pnlScore(1) = "0"
    
      picGorilla(0).Visible = False
      picGorilla(1).Visible = False
      Cls
      Me.Refresh
      ScrWidth = Me.ScaleWidth
      ScrHeight = Me.ScaleHeight
      ReDim BCoor(0 To 30) As XYPoint
      picGorilla(0).Visible = False
      picGorilla(1).Visible = False
      picGorilla(0).Refresh
      picGorilla(1).Refresh
      DoEvents
    
      CityNum = Ran(6)
      MakeCityScape BCoor(), CityNum, 0, 0, 0
      
      PlaceGorillas BCoor()
    
      picGorilla(0).Visible = True
      picGorilla(1).Visible = True
      picGorilla(0).Refresh
      picGorilla(1).Refresh
      DoEvents
      DoSun False
    
      picSun.Left = (ScrWidth \ 2) - (picSun.Width \ 2)
      picSun.Visible = True

      PlayerNum = 0

      pnl_Player1Name.Enabled = True
      pnl_Player1Name = Player1
      pnl_Player2Name = Player2

      pnl_Player1Name.BackColor = HiLiteBack
      pnl_Player1Name.ForeColor = HiLiteFore
      pnl_Player2Name.BackColor = Grey
      pnl_Player2Name.ForeColor = Black

      'EnableCursor False
    
      Timer1.Interval = 5
      Timer1.Enabled = True
    
      txtAngle1.SetFocus

End Sub

Private Sub mnu_ShowInputs_Click()

      'Set this to True while playing the game if you do not
      'want your opponent to see your inputs......
      
      mnu_ShowInputs.Checked = Not mnu_ShowInputs.Checked
    
      If mnu_ShowInputs.Checked Then
            txtAngle1.ForeColor = QBColor(0)
            txtVelocity1.ForeColor = QBColor(0)
            txtAngle2.ForeColor = QBColor(0)
            txtVelocity2.ForeColor = QBColor(0)
      Else
            txtAngle1.ForeColor = QBColor(15)
            txtVelocity1.ForeColor = QBColor(15)
            txtAngle2.ForeColor = QBColor(15)
            txtVelocity2.ForeColor = QBColor(15)
      End If
    
      WriteRegistry "Gorillas", "Show Inputs", mnu_ShowInputs.Checked
      
      DoEvents

End Sub

Private Sub mnu_Sound_Click()
    
      'Set this to True if you want to hear explosions and all.....
      mnu_Sound.Checked = Not mnu_Sound.Checked
    
      If Not mnu_Sound.Checked Then
            mnu_Sound.Caption = "Sound [Off]"
      Else
            mnu_Sound.Caption = "Sound [On]"
      End If
      
      WriteRegistry "Gorillas", "Use Sound", mnu_Sound.Checked

End Sub

Private Function OpenDB(DBName As String) As Integer

   Dim Connect As String

      On Error GoTo OpenDBErr

      Set gLIBDB = OpenDatabase(DBName)

      'success
      OpenDB = True
      GoTo OpenDBEnd

OpenDBErr:
      OpenDB = False
      Resume OpenDBEnd

OpenDBEnd:
      
End Function

Private Sub PlaceGorillas(BCoor() As XYPoint)
    
      XAdj = 14
      YAdj = 30
      Sclx = ScrWidth / 320
      Scly = ScrHeight / 200
    
      GorillaX(1) = 0
      GorillaY(1) = 0
      GorillaX(2) = 0
      GorillaY(2) = 0
  
      'Place gorillas on second or third building from edge
      For i = 1 To 2
            If i = 1 Then
                  BNum = Ran(2) + 1
            Else:
                  BNum = LastBuilding - Ran(2)
            End If
            BWidth = BCoor(BNum + 1).XCoor - BCoor(BNum).XCoor
            GorillaX(i) = BCoor(BNum).XCoor + BWidth / 2 - XAdj
            GorillaY(i) = (BCoor(BNum).YCoor - YAdj) - 7
            If GSize = 0 Then picGorilla(i - 1).Move GorillaX(i) - 10, GorillaY(i) - 38
      Next i

      OrgGorillaX1 = GorillaX(1)
      OrgGorillaY1 = GorillaY(1)
      OrgGorillaX2 = GorillaX(2)
      OrgGorillaY2 = GorillaY(2)

End Sub

Private Function PlotShot(StartX, StartY, Angle, Velocity, PlayerNum)

      Dim ox As Integer
      Dim oy As Integer
      
      If GameOver Then
            picSun.Visible = False
            pnlWindDirection.Visible = False
            lblAngle1.Visible = False
            txtAngle1.Visible = False
            lblVelocity1.Visible = False
            txtVelocity1.Visible = False
            Exit Function
      End If
    
      x = 0
      y = 0
    
      Angle = Angle / 180 * pi#  'Convert degree angle to radians
      Radius = Mode Mod 7
    
      InitXVel# = Cos(Angle) * Velocity
      InitYVel# = Sin(Angle) * Velocity
    
      oldx = StartX
      oldy = StartY
      ox = StartX: oy = StartY

      'draw gorilla toss
      If PlayerNum = 0 Then
            picGorilla(PlayerNum).Picture = frmPics.BossClip1.GraphicCell(1)
            picGorilla(PlayerNum).Refresh
            DoEvents
        
            'play sound if enabled
            If mnu_Sound.Checked Then
                  If Right(App.Path, 1) <> "\" Then
                        PlaySound App.Path & "\sounds\toss.wav"
                  Else
                        PlaySound App.Path & "sounds\toss.wav"
                  End If
            End If
        
            DoEvents
       
            rest MaxSpeed \ 2
            
            picGorilla(PlayerNum).Picture = frmPics.BossClip1.GraphicCell(2)
            picGorilla(PlayerNum).Refresh
            DoEvents
            
            rest MaxSpeed \ 2
            
                    
            'redraw gorilla
            picGorilla(PlayerNum).Picture = frmPics.BossClip1.GraphicCell(0)
            picGorilla(PlayerNum).Refresh
      Else
            picGorilla(PlayerNum).Picture = frmPics.BossClip2.GraphicCell(1)
            picGorilla(PlayerNum).Refresh
            DoEvents
        
            'Play Sound
            If mnu_Sound.Checked Then
                  If Right(App.Path, 1) <> "\" Then
                        PlaySound App.Path & "\sounds\toss.wav"
                  Else
                        PlaySound App.Path & "sounds\toss.wav"
                  End If
            End If
        
            DoEvents
        
            rest MaxSpeed \ 2
            picGorilla(PlayerNum).Picture = frmPics.BossClip2.GraphicCell(2)
            picGorilla(PlayerNum).Refresh
            DoEvents
            rest MaxSpeed \ 2
        
            'redraw gorilla
            picGorilla(PlayerNum).Picture = frmPics.BossClip2.GraphicCell(0)
            picGorilla(PlayerNum).Refresh
      End If
  
      adjust = Scl(4)                   'For scaling CGA
      xedge = Scl(9) * (2 - PlayerNum)  'Find leading edge of banana for check

      Impact = False
      ShotInSun = False
      OnScreen = True
      PlayerHit = 0
      NeedErase = False

      StartXPos = StartX
      StartYPos = StartY - adjust - 3
    
      If PlayerNum = 1 Then
            StartXPos = StartXPos + Scl(25)
            Direction = Scl(4)
      Else
            Direction = Scl(-4)
      End If
    
      pointval = 0
      PlayerHit = False

      If Velocity < 2 Then              'Shot too slow - hit self
            x = StartX
            y = StartY
            pointval = QBColor(15)
      End If
   
      Do While (Not Impact) And OnScreen
    
            'rest MaxSpeed
        
            x = StartXPos + (InitXVel# * t#) + (0.5 * (wind / 5) * t# ^ 2)
            y = StartYPos + ((-1 * (InitYVel# * t#)) + (0.5 * gravity * t# ^ 2)) * (ScrHeight / 350)
        
            BulletLeft = x + (picBullet.Width \ 2)
            BulletTop = y + (picBullet.Height \ 2)
                
            If (x >= ScrWidth - Scl(10)) Or (x <= 3) Or (y >= ScrHeight - 3) Then
                  OnScreen = False
            Else
                  OnScreen = True
            End If
    
            PlayerHit = False

            If OnScreen And y > 0 Then
                  picBullet.Visible = False
                  pointval = Point(x + (picBullet.Width \ 2), y + (picBullet.Height \ 2))
                  DoEvents
                  picBullet.Visible = True
                  
                  Select Case pointval
                        Case QBColor(5)
                              Impact = True
                        Case QBColor(6)
                              Impact = True
                        Case QBColor(7)
                              Impact = True
                        Case QBColor(9)
                              Impact = False
                              If ShotInSun Then
                                    If Abs(ScrWidth \ 2 - x) > Scl(20) Then
                                          ShotInSun = False
                                    End If
                              End If
                  End Select

                  'See if bullet hit opponent
                  Select Case PlayerNum
                        Case 0
                              If x > picGorilla(1).Left And x < picGorilla(1).Left + (picGorilla(1).Width - 30) And y > picGorilla(1).Top And y < picGorilla(1).Top + picGorilla(1).Height Then
                                    picBullet.Visible = False
                                    DoEvents
                                    PlayerHit = ExplodeGorilla(x, y)
                                    picGorilla(1).Visible = False
                                    DoEvents
                                    VictoryDance 0
                                    DoEvents
                                    Score0 = Score0 + 1
                                    UpdateScores 0, CStr(Score0)
                                    PlayerNum = 1
                                    Tosser = 1
                                    Cls
                                    picGorilla(0).Visible = False
                                    picGorilla(1).Visible = False
                                    picGorilla(0).Refresh
                                    picGorilla(1).Refresh
                                    DoEvents
                              
                                    MakeCityScape BCoor(), CityNum, 0, 0, 0

                                    PlaceGorillas BCoor()
                              
                                    picGorilla(0).Visible = True
                                    picGorilla(1).Visible = True
                                    picGorilla(0).Refresh
                                    picGorilla(1).Refresh
                                    DoEvents
                                    Angle = 0
                                    Velocity = 0
                                    pointval = 0
                                    txtAngle1 = ""
                                    txtAngle2 = ""
                                    txtVelocity1 = ""
                                    txtVelocity2 = ""
                                    picBullet.Visible = False
                                    DoEvents
                                    PlotShot2 = PlayerHit
                                    DoSun False
                                    picBullet.Visible = False
                                    DoEvents
                                    Exit Function
                              End If
                        Case 1
                              If x > picGorilla(0).Left + 30 And x < picGorilla(0).Left + picGorilla(0).Width And y > picGorilla(0).Top And y < picGorilla(0).Top + picGorilla(0).Height Then
                                    picBullet.Visible = False
                                    DoEvents
                                    PlayerHit = ExplodeGorilla(x, y)
                                    picGorilla(0).Visible = False
                                    DoEvents
                                    VictoryDance 1
                                    DoEvents
                                    Score1 = Score1 + 1
                                    UpdateScores 1, CStr(Score1)
                                    PlayerNum = 0
                                    Tosser = 0
                                    Cls
                                    picGorilla(0).Visible = False
                                    picGorilla(1).Visible = False
                                    picGorilla(0).Refresh
                                    picGorilla(1).Refresh
                                    DoEvents
                              
                                    MakeCityScape BCoor(), CityNum, 0, 0, 0

                                    PlaceGorillas BCoor()
                              
                                    picGorilla(0).Visible = True
                                    picGorilla(1).Visible = True
                                    picGorilla(0).Refresh
                                    picGorilla(1).Refresh
                                    DoEvents
                                    Angle = 0
                                    Velocity = 0
                                    pointval = 0
                                    txtAngle1 = ""
                                    txtAngle2 = ""
                                    txtVelocity1 = ""
                                    txtVelocity2 = ""
                                    picBullet.Visible = False
                                    DoEvents
                                    PlotShot2 = PlayerHit
                                    DoSun False
                                    picBullet.Visible = False
                                    DoEvents
                                    Exit Function
                              End If
                  End Select
            
                  ' Check if hit self
                  If x > picGorilla(PlayerNum).Left And x < picGorilla(PlayerNum).Left + picGorilla(PlayerNum).Width And y > picGorilla(PlayerNum).Top And y < picGorilla(PlayerNum).Top + picGorilla(PlayerNum).Height Then
                        picBullet.Visible = False
                        DoEvents
                        PlayerHit = ExplodeGorilla(x, y)
                        picGorilla(PlayerNum).Visible = False
                        DoEvents
            
                        If PlayerNum = 0 Then
                              VictoryDance 1
                              DoEvents
                              Score1 = Score1 + 1
                              UpdateScores 1, CStr(Score1)
                              PlayerNum = 1
                              Tosser = 1
                        ElseIf PlayerNum = 1 Then
                              VictoryDance 0
                              DoEvents
                              Score0 = Score0 + 1
                              UpdateScores 0, CStr(Score0)
                              PlayerNum = 0
                              Tosser = 0
                        End If
                        
                        Cls
                        picGorilla(0).Visible = False
                        picGorilla(1).Visible = False
                        picGorilla(0).Refresh
                        picGorilla(1).Refresh
                        DoEvents
                
                        MakeCityScape BCoor(), CityNum, 0, 0, 0

                        PlaceGorillas BCoor()
                
                        picGorilla(0).Visible = True
                        picGorilla(1).Visible = True
                        picGorilla(0).Refresh
                        picGorilla(1).Refresh
                        DoEvents
                        Angle = 0
                        Velocity = 0
                        pointval = 0
                        txtAngle1 = ""
                        txtAngle2 = ""
                        txtVelocity1 = ""
                        txtVelocity2 = ""
                        picBullet.Visible = False
                        DoEvents
                        PlotShot2 = PlayerHit
                        DoSun False
                        picBullet.Visible = False
                        DoEvents
                        Exit Function
                  End If

                  If x > picSun.Left And x < picSun.Left + picSun.Width And y > picSun.Top And y < picSun.Top + picSun.Height Then
                        DoSun True
                  End If

                  If Not ShotInSun And Not Impact Then
                        frmGorilla.picBullet.Move x - (frmGorilla.picBullet.Width \ 2), y - (frmGorilla.picBullet.Height \ 2)
                        frmGorilla.picBullet.Visible = True
                        frmGorilla.picBullet.Refresh
                        'rest .05
                        rest MaxSpeed \ 2
                  End If
                    
                  oldx = x
                  oldy = y
                  oldrot = rot
            Else
                  frmGorilla.picBullet.Visible = False
                  DoEvents
            End If
    
            t# = t# + 0.1

      Loop

      If pointval <> QBColor(9) And Impact Then
            frmGorilla.picBullet.Visible = False
            DoExplosion x + adjust, y + adjust
      End If
    
      SetWindSpeed True

      PlotShot2 = PlayerHit
      DoSun False
      frmGorilla.picBullet.Visible = False
      DoEvents

      StartX = 0
      StartY = 0
    
End Function

Private Sub mnuShowSplash_Click()

      mnuShowSplash.Checked = Not mnuShowSplash.Checked
      
      WriteRegistry "Gorillas", "Splash Screen", CStr(mnuShowSplash.Checked)
            
End Sub

Private Sub mnuSpeed_Click()
      frmSetSpeed.Show 1
      WriteRegistry "Gorillas", "Game Speed", CStr(MaxSpeed)
End Sub

Private Sub pnl_Control_DragDrop(Source As Control, x As Single, y As Single)
      'EnableCursor True
End Sub

Private Sub rest(t As Single)
      For ttimer& = 1 To (t * 100000)
      Next ttimer&
End Sub

Private Sub SetSize()
      lbl_GameOver.Left = (ScrWidth \ 2) - (lbl_GameOver.Width) \ 2
      lbl_GameOver.Top = ScrHeight \ 10
End Sub

Private Sub SetWindSpeed(First As Integer)

      'Set Wind speed if 1st time in this routine
      
      If First Then
            wind = Ran(5) - 5
            If Ran(3) = 1 Then
                  If wind > 0 Then
                        wind = wind + Ran(5)
                  Else
                        wind = wind - Ran(5)
                  End If
            End If
      End If

      'Draw Wind speed arrow
      If wind <> 0 Then
            frmGorilla.picWindDirection.Cls
            WindLine = wind * 3 * (ScrWidth \ 320)
            picWindDirection.Line (picWindDirection.ScaleWidth \ 2, picWindDirection.ScaleHeight - 5)-(picWindDirection.ScaleWidth \ 2 + WindLine, picWindDirection.ScaleHeight - 5), QBColor(LEDColor)
            If wind > 0 Then ArrowDir = -2 Else ArrowDir = 2
            picWindDirection.Line (picWindDirection.ScaleWidth / 2 + WindLine, picWindDirection.ScaleHeight - 5)-(picWindDirection.ScaleWidth / 2 + WindLine + ArrowDir, picWindDirection.ScaleHeight - 5 - 2), QBColor(LEDColor)
            picWindDirection.Line (picWindDirection.ScaleWidth / 2 + WindLine, picWindDirection.ScaleHeight - 5)-(picWindDirection.ScaleWidth / 2 + WindLine + ArrowDir, picWindDirection.ScaleHeight - 5 + 2), QBColor(LEDColor)
      End If

End Sub

Private Sub TilePicture(picParent As Form, picTile As PictureBox)

      Const SRCCOPY = &HCC0020
      
      Dim MaximumX As Integer, MaximumY As Integer
      Dim x As Integer, y As Integer
      Dim TileIt As Integer
      Dim TileWidth As Integer, TileHeight As Integer
    
      'Calculate MaxX and MaxY
      MaximumX = picParent.Width + picTile.Width
      MaximumY = picParent.Height + picTile.Height
    
      MaximumX = MaximumX \ Screen.TwipsPerPixelX
      MaximumY = MaximumY \ Screen.TwipsPerPixelY
   
      TileWidth = picTile.Width \ Screen.TwipsPerPixelX
      TileHeight = picTile.Height \ Screen.TwipsPerPixelY
    
      For y = 0 To MaximumY Step TileHeight
            For x = 0 To MaximumX Step TileWidth
                  TileIt = BitBlt(picParent.hDC, x, y, TileWidth, TileHeight, picTile.hDC, 0, 0, SRCCOPY)
            Next x
      Next y

End Sub

Private Sub Timer1_Timer()

      Dim curhWnd As Integer      'Current hWnd
      Dim lpPoint As POINTAPI
      Static LasthWnd As Integer  'Hold previous hWnd

      ' Make sure the program has the input focus:
      If GetActiveWindow() = frmGorilla.hwnd Then
            ' Initialize point structure:
            Call GetCursorPos(lpPoint)
            If lpPoint.y <= 39 Then
                  'EnableCursor True
            End If
      End If

End Sub

Private Sub txtAngle1_GotFocus()

      If GameOver Then
            picSun.Visible = False
            pnlWindDirection.Visible = False
            lblAngle1.Visible = False
            txtAngle1.Visible = False
            lblVelocity1.Visible = False
            txtVelocity1.Visible = False
            Exit Sub
      Else
            If Player1 = "Computer" Then
                  Randomize
                  txtAngle1 = CStr(Int(((50 - 40) + 1) * Rnd + 40))
                  Angle = Val(txtAngle1)
                  Randomize
                  txtVelocity1 = CStr(Int(((100 - 60) + 1) * Rnd + 60))
                  txtVelocity1.SetFocus
                  SendKeys "{ENTER}", True
            End If
      End If

End Sub

Private Sub txtAngle1_KeyPress(KeyAscii As Integer)

      If KeyAscii = 13 Then
            KeyAscii = 0
            Angle = Val(txtAngle1)
            If txtAngle1 = "" Then
                  MsgBox "Please enter a value for the Angle."
                  txtAngle1.SetFocus
                  Exit Sub
            End If
            txtVelocity1.SetFocus
      End If
    
End Sub

Private Sub txtAngle1_LostFocus()

      On Error Resume Next
      Angle = Val(txtAngle1)
      If txtVelocity1.Visible Then txtVelocity1.SetFocus

End Sub

Private Sub txtAngle2_GotFocus()

      If GameOver Then
            picSun.Visible = False
            picWindDirection.Visible = False
            lblAngle1.Visible = False
            txtAngle1.Visible = False
            lblVelocity1.Visible = False
            txtVelocity1.Visible = False
            Exit Sub
      Else
            If Player2 = "Computer" Then
                  Randomize
                  txtAngle2 = CStr(Int((50 - 40 + 1) * Rnd + 40))
                  Angle = Val(txtAngle2)
                  Randomize
                  txtVelocity2 = CStr(Int((100 - 60 + 1) * Rnd + 60))
                  txtVelocity2.SetFocus
                  SendKeys "{ENTER}", True
            End If
      End If

End Sub

Private Sub txtAngle2_KeyPress(KeyAscii As Integer)

      If KeyAscii = 13 Then
            KeyAscii = 0
            Angle = Val(txtAngle2)
            If txtAngle2 = "" Then
                  MsgBox "Please enter a value for the Angle."
                  txtAngle2.SetFocus
                  Exit Sub
            End If
            txtVelocity2.SetFocus
      End If

End Sub

Private Sub txtAngle2_LostFocus()
      Angle = Val(txtAngle2)
End Sub

Private Sub txtVelocity1_KeyPress(KeyAscii As Integer)

      If KeyAscii = 13 Then
            If txtVelocity1 = "" Then
                  MsgBox "Please enter a value for the Velocity."
                  txtVelocity1.SetFocus
                  Exit Sub
            End If
        
            If GameOver Then
                  picSun.Visible = False
                  pnlWindDirection.Visible = False
                  lblAngle1.Enabled = False
                  txtAngle1.Visible = False
                  lblVelocity1.Enabled = False
                  txtVelocity1.Visible = False
                  Exit Sub
            End If
        
            KeyAscii = 0
            Velocity = 0
            Velocity = Val(txtVelocity1)
        
            Tosser = 0
        
            Hit = DoShot(0, GorillaX(1), GorillaY(1), CInt(Angle), CInt(Velocity))

            GorillaX(1) = OrgGorillaX1
            GorillaY(1) = OrgGorillaY1

            PlayerNum = 0
            pnl_Player2Name.Enabled = True
            pnl_Player2Name.BackColor = HiLiteBack
            pnl_Player2Name.ForeColor = HiLiteFore
            pnl_Player1Name.BackColor = Grey
            pnl_Player1Name.ForeColor = Black
            pnl_Player1Name.Enabled = False
            Angle = 0
            Velocity = 0
        
            lblAngle1.Enabled = False
            txtAngle1.Visible = False
            lblVelocity1.Enabled = False
            txtVelocity1.Visible = False
        
            lblAngle2.Enabled = True
            txtAngle2.Visible = True
            lblVelocity2.Enabled = True
            txtVelocity2.Visible = True

            txtAngle2.SetFocus

            txtAngle1 = ""
            txtAngle2 = ""
            txtVelocity1 = ""
            txtVelocity2 = ""
    
      End If
    
End Sub

Private Sub txtVelocity1_LostFocus()
      Velocity = Val(txtVelocity1)
End Sub

Private Sub txtVelocity2_KeyPress(KeyAscii As Integer)

      If KeyAscii = 13 Then
            If txtVelocity2 = "" Then
                  MsgBox "Please enter a value for the Velocity."
                  txtVelocity2.SetFocus
                  Exit Sub
            End If
        
            If GameOver Then
                  picSun.Visible = False
                  pnlWindDirection.Visible = False
                  lblAngle1.Enabled = False
                  txtAngle1.Visible = False
                  lblVelocity1.Enabled = False
                  txtVelocity1.Visible = False
                  Exit Sub
            End If
        
            KeyAscii = 0
        
            Velocity = 0
            Velocity = Val(txtVelocity2)
        
            txtAngle1 = ""
            txtAngle2 = ""
            txtVelocity1 = ""
            txtVelocity2 = ""
        
            Tosser = 1
        
            Hit = DoShot(1, GorillaX(2), GorillaY(2), Angle, Velocity)

            GorillaX(2) = OrgGorillaX2
            GorillaY(2) = OrgGorillaY2
        
            PlayerNum = 1
            pnl_Player1Name.Enabled = True
            pnl_Player1Name.BackColor = HiLiteBack
            pnl_Player1Name.ForeColor = HiLiteFore
            pnl_Player2Name.BackColor = Grey
            pnl_Player2Name.ForeColor = Black
            pnl_Player2Name.Enabled = False
            Angle = 0
            Velocity = 0
        
            lblAngle2.Enabled = False
            txtAngle2.Visible = False
            lblVelocity2.Enabled = False
            txtVelocity2.Visible = False
        
            lblAngle1.Enabled = True
            txtAngle1.Visible = True
            lblVelocity1.Enabled = True
            txtVelocity1.Visible = True

            txtAngle1.SetFocus
      End If

End Sub

Private Sub txtVelocity2_LostFocus()
      Velocity = Val(txtVelocity2)
End Sub

Private Sub UpdateScores(PlayerNum, Sc As String)

   Dim dNum As Integer
   Dim s As Integer
   Dim n As Integer
      
      If Len(Sc) <= 0 Then Exit Sub

      dNum = 2

      Select Case PlayerNum
            Case 0
                  pnlScore(0) = CStr(Score0)
                  If Sc = CStr(NumGames) Then
                        frmGorilla.datHiScores.Recordset.FindFirst "PlayerName = '" & CStr(Player1) & "'"
                        If frmGorilla.datHiScores.Recordset.NoMatch Then
                              frmGorilla.datHiScores.Recordset.AddNew
                              frmGorilla.datHiScores.Recordset!PlayerName = Player1
                              frmGorilla.datHiScores.Recordset!Games = "1"
                              frmGorilla.datHiScores.Recordset!Date = Format$(Now, "mmmm dd, yyyy")
                              frmGorilla.datHiScores.Recordset.Update
                        Else
                              frmGorilla.datHiScores.Recordset.Edit
                              Games = CInt(frmGorilla.datHiScores.Recordset.Fields("Games")) + 1
                              frmGorilla.datHiScores.Recordset!Games = CStr(Games)
                              frmGorilla.datHiScores.Recordset!Date = Format$(Now, "mmmm dd, yyyy")
                              frmGorilla.datHiScores.Recordset.Update
                        End If
      
                        frmWinner.timerWinner.Enabled = True
                        frmWinner.btnWinner.Caption = UCase$(Player1)
                        frmWinner.Show 1
                        lbl_GameOver.Visible = True
                        GameOver = True
                        picSun.Visible = False
                        pnlWindDirection.Visible = False
                        lblAngle1.Visible = False
                        txtAngle1.Visible = False
                        lblVelocity1.Visible = False
                        txtVelocity1.Visible = False
                        Exit Sub
                  End If
                  
            Case 1
                  pnlScore(1).Caption = CStr(Score1)
                  If Sc = CStr(NumGames) Then
                        frmGorilla.datHiScores.Recordset.FindFirst "PlayerName = '" & CStr(Player2) & "'"
                        If frmGorilla.datHiScores.Recordset.NoMatch Then
                              frmGorilla.datHiScores.Recordset.AddNew
                              frmGorilla.datHiScores.Recordset!PlayerName = Player2
                              frmGorilla.datHiScores.Recordset!Games = "1"
                              frmGorilla.datHiScores.Recordset!Date = Format$(Now, "mmmm dd, yyyy")
                              frmGorilla.datHiScores.Recordset.Update
                        Else
                              frmGorilla.datHiScores.Recordset.Edit
                              Games = CInt(frmGorilla.datHiScores.Recordset!Games) + 1
                              frmGorilla.datHiScores.Recordset!Games = CStr(Games)
                              frmGorilla.datHiScores.Recordset!Date = Format$(Now, "mmmm dd, yyyy")
                              frmGorilla.datHiScores.Recordset.Update
                        End If
                  
                        frmWinner.timerWinner.Enabled = True
                        frmWinner.btnWinner.Caption = UCase$(Player2)
                        frmWinner.Show 1
                      
                        lbl_GameOver.Visible = True
                        picSun.Visible = False
                        pnlWindDirection.Visible = False
                        lblAngle1.Visible = False
                        txtAngle1.Visible = False
                        lblVelocity1.Visible = False
                        txtVelocity1.Visible = False
                        GameOver = True
                        Exit Sub
                  End If
      End Select

End Sub

Private Sub VictoryDance(Player)

      For i = 1 To 4
            picGorilla(Player).Picture = frmPics.BossClip1.GraphicCell(1)
            DoEvents
            rest 12
            picGorilla(Player).Picture = frmPics.BossClip1.GraphicCell(2)
            DoEvents
            rest 12
            picGorilla(Player).Picture = frmPics.BossClip1.GraphicCell(1)
            DoEvents
            rest 12
            If Player = 1 Then
                  frmGorilla.picGorilla(Player).Picture = frmPics.BossClip1.GraphicCell(0)
            Else
                  frmGorilla.picGorilla(Player).Picture = frmPics.BossClip2.GraphicCell(0)
            End If
            DoEvents
            rest 12
            picGorilla(Player).Picture = frmPics.BossClip2.GraphicCell(1)
            DoEvents
            rest 12
            picGorilla(Player).Picture = frmPics.BossClip2.GraphicCell(2)
            DoEvents
            rest 12
            picGorilla(Player).Picture = frmPics.BossClip2.GraphicCell(1)
            DoEvents
            rest 12
            If Player = 1 Then
                  frmGorilla.picGorilla(Player).Picture = frmPics.BossClip1.GraphicCell(0)
            Else
                  frmGorilla.picGorilla(Player).Picture = frmPics.BossClip2.GraphicCell(0)
            End If
            DoEvents
            rest 12
      Next i
    
      If Player = 1 Then
            frmGorilla.picGorilla(Player).Picture = frmPics.BossClip1.GraphicCell(0)
      Else
            frmGorilla.picGorilla(Player).Picture = frmPics.BossClip2.GraphicCell(0)
      End If
  
      DoEvents

End Sub
