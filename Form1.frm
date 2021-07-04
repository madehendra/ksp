VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1530
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   780
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   -15
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   3825
   End
   Begin VB.Image Gambar 
      Height          =   1530
      Index           =   11
      Left            =   960
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4485
   End
   Begin VB.Image Gambar 
      Height          =   1530
      Index           =   10
      Left            =   960
      Picture         =   "Form1.frx":166DC
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4485
   End
   Begin VB.Image Gambar 
      Height          =   1530
      Index           =   9
      Left            =   960
      Picture         =   "Form1.frx":2CDB8
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4485
   End
   Begin VB.Image Gambar 
      Height          =   1530
      Index           =   8
      Left            =   960
      Picture         =   "Form1.frx":43494
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4485
   End
   Begin VB.Image Gambar 
      Height          =   1575
      Index           =   7
      Left            =   960
      Picture         =   "Form1.frx":59B70
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4605
   End
   Begin VB.Image Gambar 
      Height          =   1575
      Index           =   6
      Left            =   960
      Picture         =   "Form1.frx":7024C
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4605
   End
   Begin VB.Image Gambar 
      Height          =   1575
      Index           =   5
      Left            =   960
      Picture         =   "Form1.frx":86928
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4605
   End
   Begin VB.Image Gambar 
      Height          =   1575
      Index           =   4
      Left            =   960
      Picture         =   "Form1.frx":9D004
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4605
   End
   Begin VB.Image Gambar 
      Height          =   1575
      Index           =   3
      Left            =   960
      Picture         =   "Form1.frx":B36E0
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4605
   End
   Begin VB.Image Gambar 
      Height          =   1575
      Index           =   2
      Left            =   960
      Picture         =   "Form1.frx":C9DBC
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4605
   End
   Begin VB.Image Gambar 
      Height          =   1575
      Index           =   1
      Left            =   960
      Picture         =   "Form1.frx":E0498
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4620
   End
   Begin VB.Image Gambar 
      Height          =   1575
      Index           =   0
      Left            =   960
      Picture         =   "Form1.frx":F6B74
      Stretch         =   -1  'True
      Top             =   330
      Width           =   4605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer

Private Sub Form_Load()
Dim n As Integer
  
  CenterForm Me
  For n = 0 To 11
    Gambar(n).Visible = False
  Next
  a = 0
End Sub

Private Sub Timer1_Timer()
  If a = 12 Then
    Unload Me
    Exit Sub
  End If
  Image1.Picture = Gambar(a).Picture
  Gambar(a).Visible = False
  a = a + 1
End Sub
