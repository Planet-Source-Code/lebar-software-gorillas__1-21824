Attribute VB_Name = "functions"
Global stWaveFile As String

Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Registry API prototypes
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_READ = &H20000
Public Const STANDARD_RIGHTS_WRITE = &H20000
Public Const STANDARD_RIGHTS_EXECUTE = &H20000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const REG_DWORD = 4
Public Const REG_BINARY = 3
Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpname As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Public Const REG_NONE = (0)                         'No value type
Public Const REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
Public Const REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
Public Const REG_LINK = (6)                         'Symbolic Link (unicode)
Public Const REG_MULTI_SZ = (7)                     'Multiple Unicode strings
Public Const REG_RESOURCE_LIST = (8)                'Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = (10)
Const READ_CONTROL = &H20000
Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Boolean
End Type
Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
End Type
Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Const LB_ITEMFROMPOINT = &H1A9

Public Sub PlaySound(strFileName As String)
    sndPlaySound strFileName, 1
End Sub


Public Sub GetSettings()
    
      'Get Game Speed
      If ReadRegistry("Gorillas", "Game Speed") <> "" Then
            MaxSpeed = CLng(ReadRegistry("Gorillas", "Game Speed"))
      Else
            MaxSpeed = 20
            CreateRegEntry "Gorillas", "Game Speed", MaxSpeed
      End If
    
      'Get LED Color
      If ReadRegistry("Gorillas", "LED Color") <> "" Then
            LEDColor = CInt(ReadRegistry("Gorillas", "LED Color"))
      Else
            LEDColor = 11
            CreateRegEntry "Gorillas", "LED Color", LEDColor
      End If
      
      frmGorilla.mnu_Colors(0).Checked = False
      frmGorilla.mnu_Colors(1).Checked = False
      frmGorilla.mnu_Colors(2).Checked = False
      frmGorilla.mnu_Colors(3).Checked = False
      
      Select Case LEDColor
            Case 12
                  LEDColor = 12
                  frmGorilla.mnu_Colors(0).Checked = True
            Case 10
                  LEDColor = 10
                  frmGorilla.mnu_Colors(1).Checked = True
            Case 11
                  LEDColor = 11
                  frmGorilla.mnu_Colors(2).Checked = True
            Case 14
                  LEDColor = 14
                  frmGorilla.mnu_Colors(3).Checked = True
      End Select

      frmGorilla.pnlScore(0).ForeColor = QBColor(LEDColor)
      frmGorilla.pnlScore(1).ForeColor = QBColor(LEDColor)

      frmGorilla.txtVelocity1.ForeColor = QBColor(LEDColor)
      frmGorilla.txtVelocity2.ForeColor = QBColor(LEDColor)
      
      'Get Players
      If ReadRegistry("Gorillas", "Player1") <> "" Then
            Player1 = ReadRegistry("Gorillas", "Player1")
      Else
            Player1 = "Player 1"
            CreateRegEntry "Gorillas", "Player1", Player1
      End If
      
      If ReadRegistry("Gorillas", "Player2") <> "" Then
            Player2 = ReadRegistry("Gorillas", "Player2")
      Else
            Player2 = "Player 2"
            CreateRegEntry "Gorillas", "Player2", Player2
      End If
      
      'Get Total Games
      If ReadRegistry("Gorillas", "Total Games") <> "" Then
            NumGames = ReadRegistry("Gorillas", "Total Games")
      Else
            NumGames = 3
            CreateRegEntry "Gorillas", "Total Games", NumGames
      End If
      
      'Get Gravity
      If ReadRegistry("Gorillas", "Gravity") <> "" Then
            gravity = ReadRegistry("Gorillas", "Gravity")
      Else
            gravity = 9.3
            CreateRegEntry "Gorillas", "Gravity", gravity
      End If
      
      'Get Random Height
      If ReadRegistry("Gorillas", "RandomHeight") <> "" Then
            gravity = ReadRegistry("Gorillas", "RandomHeight")
      Else
            RandomHeight = 120
            CreateRegEntry "Gorillas", "RandomHeight", RandomHeight
      End If

End Sub

Public Function ReadRegistry(sPath As String, sValue As String)
'As String

    Dim lKeyHand As Long
    Dim lValueType As Long
    Dim lResult As Long
    Dim sBuff As String
    Dim lDataBufSize As Long
    Dim iZeroPos As Integer
    
    'MsgBox "Software\Deskmate\" + sPath
    
    RegOpenKey HKEY_CURRENT_USER, "Software\Deskmate\" + sPath, lKeyHand                            'open the key we are looking at
    lResult = RegQueryValueEx(lKeyHand, sValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then                                 'is it a Null terminated string?
        sBuff = String$(lDataBufSize, " ")                      'set the Buffer size
        lResult = RegQueryValueEx(lKeyHand, sValue, 0&, 0&, ByVal sBuff, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            iZeroPos = InStr(sBuff, Chr$(0))                    'is there a null in the returned value?
            If Not iZeroPos = 0 Then
                'ReadRegistry = Left(sBuff, iZeroPos - 1)      'yes? then get everything up to the null, and return the value
                ReadRegistry = Mid$(sBuff, 1, iZeroPos - 1)     'yes? then get everything up to the null, and return the value
            Else
                ReadRegistry = sBuff                          'no?  then return the value
            End If
        End If
    End If
    RegCloseKey lKeyHand
    
End Function

Public Sub CreateRegEntry(sPath As String, sKey As String, sValue As Variant)
    
    Dim hKey As Long            ' receives handle to the registry key
    Dim secattr As SECURITY_ATTRIBUTES  ' security settings for the key
    Dim subkey As String        ' name of the subkey to create or open
    Dim neworused As Long       ' receives flag for if the key was created or opened
    Dim stringbuffer As String  ' the string to put into the registry
    Dim retval As Long          ' return value
    
    ' Set the name of the new key and the default security settings
    'subkey = "Software\Deskmate\Settings"
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = 1
    
    ' Create (or open) the registry key.
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Deskmate\" + sPath, 0, "", 0, KEY_WRITE, secattr, hKey, neworused)
    If retval <> 0 Then
        MsgBox "Error opening or creating registry key -- aborting."
        Exit Sub
    End If
    
    ' Write the string to the registry.  Note the use of ByVal in the second-to-last
    ' parameter because we are passing a string.
    stringbuffer = sValue & vbNullChar  ' the terminating null is necessary
    retval = RegSetValueEx(hKey, sKey, 0, REG_SZ, ByVal stringbuffer, Len(stringbuffer))
    
    ' Close the registry key.
    retval = RegCloseKey(hKey)
   
End Sub

Public Sub WriteRegistry(subkey As String, ValueName As String, vNewValue As String)
    Dim Result As Long, retval As Long
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Deskmate\" + subkey, 0, KEY_ALL_ACCESS, Result)
    retval = RegSetValueEx(Result, ValueName, 0, REG_SZ, vNewValue, CLng(Len(vNewValue) + 1))
    RegCloseKey HKEY_CURRENT_USER
    RegCloseKey Result
End Sub


