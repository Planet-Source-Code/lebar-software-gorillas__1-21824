Attribute VB_Name = "GORILLA1"
DefInt A-Z

Global GameInPlay As Integer
Global DBName As String
Global gLIBDB As Database
Global GameOver As Integer
Global gDS As Dynaset
Global GSize As Integer
Global Tosser As Integer
Global LEDColor As Integer
Global Player1 As String
Global Player2 As String
Global Score0 As Integer
Global Score1 As Integer
Global NumGames As Integer
Global MaxSpeed As Single
Global CityNum As Integer

Global OrgGorillaX1 As Integer
Global OrgGorillaY1 As Integer
Global OrgGorillaX2 As Integer
Global OrgGorillaY2 As Integer

Global Const HiLiteBack = &H800000
Global Const HiLiteFore = &HFFFFFF
Global Const Black = &H0&
Global Const Grey = &HC0C0C0

'User-Defined TYPEs
Type XYPoint
      BLeft As Integer
      BTop As Integer
      BWidth As Integer
      BHeight As Integer
      XCoor As Integer
      YCoor As Integer
      WCoor As Integer
      BColor As Integer
End Type

'Constants
Global Const WINDOWCOLOR As Integer = 14

'Global Variables
Global DefBWidth As Integer
Global RandomHeight As Integer
Global CurBuilding As Integer
Global GorillaX(1 To 2)   'Location of the two gorillas
Global GorillaY(1 To 2)   'Location of the two gorillas
Global LastBuilding As Integer
Global BCoor() As XYPoint
Global pi#
Global Angle As Integer
Global Velocity As Integer
Global gravity As Integer
Global wind As Integer
Global PlayerNum As Integer

'Screen Mode Variables
Global ScrHeight As Integer
Global ScrWidth As Integer
Global Mode As Integer

Global SunHit As Integer
Global GHeight As Integer
Global MachSpeed As Single

Global oldx As Integer
Global oldy As Integer
Global oldrot As Integer
Global pointval As Long

Global Const SW_SHOWNOACTIVATE As Integer = 4

' Stores location of cursor
Type POINTAPI
      x As Integer
      y As Integer
End Type

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'CalcDelay:
'  Checks speed of the machine.
Function CalcDelay() As Single

   Dim s As Single, i As Single
         
      s = Timer
      
      Do
            i = i + 1
      Loop Until Timer - s >= 0.5
      
      CalcDelay! = i

End Function

Sub CenterForm(frm As Form)
    
      Dim x, y

      x = (Screen.Width - frm.Width) / 2
      y = (Screen.Height - frm.Height) / 2
      frm.Move x, y
    
End Sub

Sub DisplayHelp(Help$)

   Dim lpPoint As POINTAPI ' Cursor Point variable
   Dim ret As Integer      ' Return value of ShowWindow() API function

      Rem Display Help String
      Rem
      Rem This Function displays the Help$ if Help$ <> "".
      Rem if Help$ = "" then the Help String is removed.
      Rem
      Rem FUNCTION REQUIREMENTS:
      Rem     GetCursorPos()    Windows API function
      Rem     frmHelp           Name of the Help form
      Rem

      If Len(Help$) <> 0 Then  ' Double check help$

            ' Make sure help form is invisible:
            frmHelp.Hide

            ' Change caption of label:

            frmHelp.Label1.Caption = Help$

            ' Get the cursor position so you can calculate where to place the
            ' help form:
            Call GetCursorPos(lpPoint)

            ' Offset the form from the cursor by 18 and 2 pixels (values
            ' chosen to simulate the look of Microsoft Word version 6.0)
            frmHelp.Top = (lpPoint.y + 18) * Screen.TwipsPerPixelY
            frmHelp.Left = (lpPoint.x - 2) * Screen.TwipsPerPixelY

            ' Adjust width of form to label + 4  because 2 are needed for each
            ' pixel of the border and 2 are needed to center the label (the
            ' label is inset by 1 pixel on the form). Also, adjust height of

            ' form to height of label + 2 because 2 ar needed for each pixel
            ' of the border:
            frmHelp.Width = frmHelp.Label1.Width + (4 * Screen.TwipsPerPixelX)
            frmHelp.Height = frmHelp.Label1.Height + 2 * Screen.TwipsPerPixelY

            ' Make sure form is on top:
            frmHelp.ZOrder

            ' Show form without the focus:
            ret = ShowWindow(frmHelp.hwnd, SW_SHOWNOACTIVATE)
      Else
            ' Hide the form:
            frmHelp.Hide
      End If

End Sub

Sub DoSun(Mouth)

      'mouth
      If Mouth Then  'draw "o" mouth
            frmGorilla.picSun.Picture = frmPics.SunClip.GraphicCell(1)
      Else           'draw smile
            frmGorilla.picSun.Picture = frmPics.SunClip.GraphicCell(0)
      End If
  
End Sub

Sub Main()

      On Error Resume Next
      
      Dim x
      GameInPlay = False

      frmGorilla.Visible = True
    
      DBName = App.Path + "\Gorilla.mdb"
      frmGorilla.datHiScores.DatabaseName = DBName
      x% = OpenDB(DBName)
      frmGorilla.datHiScores.RecordSource = "select * from Scores order by Games DESC;"
      frmGorilla.datHiScores.Refresh
    
      Set gDS = frmGorilla.datHiScores.Recordset.Clone()
      gDS.MoveFirst
      
      If ReadRegistry("Gorillas", "Splash Screen") Then
            frmSplash.Show 1
      End If
    
End Sub

Function OpenDB(DBName As String) As Integer

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

Function Ran(x As Integer) As Integer

      Ran = Int(Rnd(1) * x) + 1

End Function

'Scl:
'  Pass the number in to scaling for cga.  If the number is a decimal, then we
'  want to scale down for cga or scale up for ega.  This allows a full range
'  of numbers to be generated for scaling.
'  (i.e. for 3 to get scaled to 1, pass in 2.9)
Function Scl(n!)

      If n! <> Int(n!) Then
            If Mode = 1 Then n! = n! - 1
      End If
      
      If Mode = 1 Then
            Scl = CInt(n! / 2 + 0.1)
      Else
            Scl = CInt(n!)
      End If

End Function
