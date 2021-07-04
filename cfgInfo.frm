VERSION 5.00
Begin VB.Form cfgInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Information"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   615
      Left            =   5085
      Picture         =   "cfgInfo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "cfgInfo.frx":1CFA
      Top             =   60
      Width           =   6525
   End
End
Attribute VB_Name = "cfgInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mysql As New cMysql

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
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
