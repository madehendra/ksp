VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FrmPB 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   345
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   6375
   Begin vbalProgBarLib6.vbalProgressBar PB 
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   503
      Picture         =   "FrmPB.frx":0000
      ForeColor       =   0
      BarPicture      =   "FrmPB.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
End
Attribute VB_Name = "FrmPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMaximal As Double

Sub RunPB()
  On Error Resume Next
  PB.Value = IIf(PB.Value >= PB.Max, PB.Max, PB.Value + 1)
  PB.Text = CLng(PB.Percent) & "%"
End Sub

Sub InitPB(ByRef nMax As Long)
  If nMax <= 0 Then
    nMax = 1
  End If
  Screen.MousePointer = vbHourglass
  PB.ShowText = True
  PB.Value = 0
  PB.Max = nMax
  PB.Min = 0
  PB.Visible = True
  nMaximal = nMax
  Me.Show
End Sub

Sub EndPB()
  Screen.MousePointer = vbDefault
  PB.Visible = False
  Unload Me
End Sub

Private Sub Form_Load()
Dim nTinggi As Double
Dim nTinggi1 As Double
Dim nSisa As Double

  CenterForm Me
'  nTinggi = aMainmenu.Height
'  nTinggi1 = Me.Height
'  nSisa = nTinggi - nTinggi1
'  Me.Top = nSisa - 1500
End Sub

Private Function GetPersen(ByVal nNilai As Double) As Double
Dim nPersen As Double
Dim n As Double
Dim va

    va = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
             21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, _
             41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 67, 68, 69, 60, _
             61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, _
             81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100)
    nPersen = Round(nNilai / nMaximal * 100, 2)
    For n = 0 To UBound(va)
      If nPersen >= va(n) Then
        GetPersen = va(n)
      End If
    Next
End Function
