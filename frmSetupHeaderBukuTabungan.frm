VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form frmSetupHeaderBukuTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETUP MARGIN VALIDASI TABUNGAN"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5010
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1650
      Left            =   0
      Top             =   0
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   2910
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSANumberBoxProject.BiSANumberBox nTop1 
         Height          =   330
         Left            =   705
         TabIndex        =   1
         Top             =   690
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TOP MARGIN"
         CaptionWidth    =   1500
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nLeft1 
         Height          =   330
         Left            =   705
         TabIndex        =   2
         Top             =   1080
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LEFT MARGIN"
         CaptionWidth    =   1500
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "mm"
         Height          =   270
         Left            =   3345
         TabIndex        =   4
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Label5 
         Caption         =   "mm"
         Height          =   270
         Left            =   3345
         TabIndex        =   3
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "VALIDASI TABUNGAN"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   0
         Top             =   210
         Width           =   3090
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   1635
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   1138
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   2760
         TabIndex        =   5
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         Caption         =   "    &Save"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "frmSetupHeaderBukuTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   3840
         TabIndex        =   6
         Top             =   90
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   767
         Caption         =   "     &Exit"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "frmSetupHeaderBukuTabungan.frx":0416
      End
   End
End
Attribute VB_Name = "frmSetupHeaderBukuTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  If MsgBox("Data benar-benar sudah valid ?", vbYesNo + vbInformation) = vbYes Then
      UpdCfg msTopValidasiTabungan, nTop1.Value
      UpdCfg msLeftValidasiTabungan, nLeft1.Value
      
      MsgBox "Data telah tersimpan", vbInformation
      Exit Sub
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    cmdSimpan_Click
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  TabIndex nTop1, n
  TabIndex nLeft1, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  nTop1.Value = aCfg(msTopValidasiTabungan)
  nLeft1.Value = aCfg(msLeftValidasiTabungan)
End Sub

