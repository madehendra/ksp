VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form cfgBukuTabungan2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buku SIMPANAN"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3255
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1410
      Left            =   15
      Top             =   -15
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2487
      Caption         =   "Cetak Validasi"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSANumberBoxProject.BiSANumberBox nLeftMargin 
         Height          =   330
         Left            =   525
         TabIndex        =   0
         Top             =   330
         Width           =   2115
         _ExtentX        =   3731
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
         Caption         =   "Batas Kiri"
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
      Begin BiSANumberBoxProject.BiSANumberBox nTopMargin 
         Height          =   330
         Left            =   525
         TabIndex        =   1
         Top             =   690
         Width           =   2115
         _ExtentX        =   3731
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
         Caption         =   "Batas Atas"
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
      Begin VB.Label Label2 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2700
         TabIndex        =   5
         Top             =   750
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2685
         TabIndex        =   4
         Top             =   375
         Width           =   315
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   15
      Top             =   1395
      Width           =   3225
      _ExtentX        =   5689
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
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   1020
         TabIndex        =   2
         Top             =   75
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
         Picture         =   "cfgBukuTabungan2.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   2085
         TabIndex        =   3
         Top             =   75
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
         Picture         =   "cfgBukuTabungan2.frx":0416
      End
   End
End
Attribute VB_Name = "cfgBukuTabungan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  UpdCfg msLeftValidasiTabungan, nLeftMargin.Value
  UpdCfg msTopValidasiTabungan, nTopMargin.Value
  MsgBox "Data telah tesimpan"
  Unload Me
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  nLeftMargin.Value = aCfg(msLeftValidasiTabungan)
  nTopMargin.Value = aCfg(msTopValidasiTabungan)
  TabIndex nLeftMargin, n
  TabIndex nTopMargin, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub
