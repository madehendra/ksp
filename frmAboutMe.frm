VERSION 5.00
Begin VB.Form frmAboutMe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Me.."
   ClientHeight    =   5940
   ClientLeft      =   2340
   ClientTop       =   1860
   ClientWidth     =   7485
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4099.894
   ScaleMode       =   0  'User
   ScaleWidth      =   7028.804
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   1545
      Left            =   975
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmAboutMe.frx":0000
      Top             =   3135
      Width           =   6360
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4485
      Left            =   60
      Picture         =   "frmAboutMe.frx":0008
      ScaleHeight     =   4485
      ScaleWidth      =   885
      TabIndex        =   7
      Top             =   180
      Width           =   885
   End
   Begin VB.CommandButton cmdTectSupport 
      Caption         =   "&Tech Support..."
      Height          =   345
      Left            =   6105
      TabIndex        =   6
      Top             =   2670
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   6105
      TabIndex        =   0
      Top             =   1905
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   6105
      TabIndex        =   1
      Top             =   2295
      Width           =   1245
   End
   Begin VB.Line Line1 
      X1              =   70.429
      X2              =   6958.375
      Y1              =   3323.399
      Y2              =   3323.399
   End
   Begin VB.Label Label1 
      Caption         =   "Label Copyright"
      Height          =   240
      Left            =   1020
      TabIndex        =   8
      Top             =   2880
      Width           =   3780
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   1890
      Left            =   1050
      TabIndex        =   2
      Top             =   870
      Width           =   5370
   End
   Begin VB.Label lblTitle 
      Caption         =   "i-LPD "
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   4545
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   525
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   90
      TabIndex        =   3
      Top             =   4920
      Width           =   7275
   End
End
Attribute VB_Name = "frmAboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mysql As New cMysql

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim dYear As Date
'tahun pertama pengembangan
dYear = DateSerial(2004, 1, 1)

    CenterForm Me
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDisclaimer.Caption = "Warning... This computer program is protected by copyright law and international treaties. Unauthorized reproduction or distribution of this program, or any portion of it, may result in severe civil and criminal penalties, and will be prosecuted to the maximum extent possible under law"
    lblDescription.Caption = _
    "Contact support: " & vbCrLf & _
    " i-Software Development Division " & vbCrLf & _
    " Part of Surya Media Group" & vbCrLf & _
    " Located at BALI ISLAND - INDONESIA" & vbCrLf & _
    " " & vbCrLf & _
    " Technical Support : i made hendra " & vbCrLf & _
    "                     email: made.hendra@gmail.com " & vbCrLf & _
    "                     http://www.ibelog.tk " & vbCrLf & _
    "                     phone: +62 81 338 414 828 "
    If Year(Now) > Year(dYear) Then
      Label1.Caption = "Copyright " & Year(dYear) & "-" & Year(Now) & " i-Software. All Rights Reserved."
    Else
      Label1.Caption = "Copyright " & Year(Now) & " i-Software. All Rights Reserved."
    End If
    GetSysInfo
End Sub

Function CenterForm(bForm As Form, Optional ByVal lZeroTopLeft As Boolean = False)
  If lZeroTopLeft Then
    bForm.left = 0
    bForm.Top = 0
  Else
    bForm.left = (Screen.Width / 2) - (bForm.Width / 2) - 100
    bForm.Top = (Screen.Height / 2) - (bForm.Height / 2) - 750
   End If
End Function

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, Keyname As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, Keyname, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub GetSysInfo()
Dim cIPNumber As String
Dim cDatabase As String
Dim cDSN As String
Dim cPort As String
  
  CenterForm Me
  mysql.connect "localhost", GetRegistry(reg_ServerUID), GetRegistry(reg_ServerPWD)
'  GetIPNumber cIPNumber, cDatabase, cDSN, cPort
  Text1.Text = _
  "Server IP " & vbTab & vbTab & " : " & cIPNumber & vbCrLf & _
  "Database " & vbTab & " : " & cDatabase & vbCrLf & _
  "DSN Name " & vbTab & " : " & cDSN & vbCrLf & _
  "User Login " & vbTab & " : " & GetRegistry(reg_UserName) & vbCrLf & _
  "User Level " & vbTab & " : " & GetRegistry(reg_UserLevel) & vbCrLf & _
  "MySQL Versi " & vbTab & " : " & mysql.get_server_info & vbCrLf & _
  "Client Versi " & vbTab & " : " & mysql.get_client_info & vbCrLf & _
  "Host Info " & vbTab & vbTab & " : " & mysql.get_host_info & vbCrLf & _
  "Server Status " & vbCrLf & _
  mysql.stat
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub


