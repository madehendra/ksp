VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form cfgBukuTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KONFIGURASI BUKU TABUNGAN"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7275
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   2640
      Width           =   7245
      _ExtentX        =   12779
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
         Left            =   4980
         TabIndex        =   23
         Top             =   105
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
         Picture         =   "cfgBukuTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6060
         TabIndex        =   24
         Top             =   105
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
         Picture         =   "cfgBukuTabungan.frx":0416
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   2640
      Left            =   3705
      Top             =   0
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   4657
      Caption         =   "LEBAR KOLOM"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSANumberBoxProject.BiSANumberBox nNomor 
         Height          =   330
         Left            =   345
         TabIndex        =   0
         Top             =   345
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
         Caption         =   "NOMOR"
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
      Begin BiSANumberBoxProject.BiSANumberBox nTgl 
         Height          =   330
         Left            =   345
         TabIndex        =   1
         Top             =   705
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
         Caption         =   "TANGGAL"
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
      Begin BiSANumberBoxProject.BiSANumberBox nSandi 
         Height          =   330
         Left            =   345
         TabIndex        =   2
         Top             =   1065
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
         Caption         =   "SANDI"
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
      Begin BiSANumberBoxProject.BiSANumberBox nDebet 
         Height          =   330
         Left            =   345
         TabIndex        =   3
         Top             =   1425
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
         Caption         =   "DEBET"
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
      Begin BiSANumberBoxProject.BiSANumberBox nKredit 
         Height          =   330
         Left            =   345
         TabIndex        =   4
         Top             =   1785
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
         Caption         =   "KREDIT"
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldo 
         Height          =   330
         Left            =   345
         TabIndex        =   5
         Top             =   2145
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
         Caption         =   "SALDO"
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
      Begin VB.Label Label14 
         Caption         =   "mm"
         Height          =   270
         Left            =   2985
         TabIndex        =   11
         Top             =   2205
         Width           =   360
      End
      Begin VB.Label Label13 
         Caption         =   "mm"
         Height          =   270
         Left            =   2985
         TabIndex        =   10
         Top             =   1890
         Width           =   360
      End
      Begin VB.Label Label12 
         Caption         =   "mm"
         Height          =   270
         Left            =   2985
         TabIndex        =   9
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label Label11 
         Caption         =   "mm"
         Height          =   270
         Left            =   2985
         TabIndex        =   8
         Top             =   1515
         Width           =   360
      End
      Begin VB.Label Label5 
         Caption         =   "mm"
         Height          =   270
         Left            =   2985
         TabIndex        =   7
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label6 
         Caption         =   "mm"
         Height          =   270
         Left            =   2985
         TabIndex        =   6
         Top             =   765
         Width           =   360
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   2640
      Left            =   0
      Top             =   0
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   4657
      Caption         =   "BATAS"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSANumberBoxProject.BiSANumberBox nTop1 
         Height          =   330
         Left            =   675
         TabIndex        =   12
         Top             =   660
         Width           =   2160
         _ExtentX        =   3810
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
         Caption         =   "HALAMAN 1"
         CaptionWidth    =   1200
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
      Begin BiSANumberBoxProject.BiSANumberBox nTop2 
         Height          =   330
         Left            =   675
         TabIndex        =   13
         Top             =   1020
         Width           =   2160
         _ExtentX        =   3810
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
         Caption         =   "HALAMAN 2"
         CaptionWidth    =   1200
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
      Begin BiSANumberBoxProject.BiSANumberBox nLeft 
         Height          =   330
         Left            =   270
         TabIndex        =   14
         Top             =   1440
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
         Caption         =   "BATAS KIRI"
         CaptionWidth    =   1600
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
      Begin BiSANumberBoxProject.BiSANumberBox nWidth 
         Height          =   330
         Left            =   270
         TabIndex        =   15
         Top             =   1800
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
         Caption         =   "LEBAR KERTAS"
         CaptionWidth    =   1600
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
      Begin BiSANumberBoxProject.BiSANumberBox nHeight 
         Height          =   330
         Left            =   270
         TabIndex        =   16
         Top             =   2160
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
         Caption         =   "TINGGI KJERTAS"
         CaptionWidth    =   1600
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
      Begin VB.Label Label10 
         Caption         =   "mm"
         Height          =   270
         Left            =   2940
         TabIndex        =   22
         Top             =   2235
         Width           =   360
      End
      Begin VB.Label Label9 
         Caption         =   "mm"
         Height          =   270
         Left            =   2940
         TabIndex        =   21
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label Label8 
         Caption         =   "mm"
         Height          =   270
         Left            =   2940
         TabIndex        =   20
         Top             =   1860
         Width           =   360
      End
      Begin VB.Label Label7 
         Caption         =   "BATAS ATAS"
         Height          =   330
         Left            =   270
         TabIndex        =   19
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "mm"
         Height          =   270
         Left            =   2940
         TabIndex        =   18
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "mm"
         Height          =   270
         Left            =   2940
         TabIndex        =   17
         Top             =   1080
         Width           =   360
      End
   End
End
Attribute VB_Name = "cfgBukuTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New BiSAMyDLL.data

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  objData.OpenConnection GetDSN
  UpdCfg msTopMargin1, nTop1.Value, objData
  UpdCfg msTopMargin2, nTop2.Value, objData
  UpdCfg msLeftMargin, nLeft.Value, objData
  UpdCfg msLebarNomor, nNomor.Value, objData
  UpdCfg msLebarTgl, nTgl.Value, objData
  UpdCfg msLebarSandi, nSandi.Value, objData
  UpdCfg msLebarMutasi, nDebet.Value, objData
  UpdCfg msLebarKredit, nKredit.Value, objData
  UpdCfg msLebarSaldo, nSaldo.Value, objData
  UpdCfg msPaperHight, nHeight.Value, objData
  UpdCfg msPaperWidth, nWidth.Value, objData
  objData.CloseConnection GetDSN
  MsgBox "Data Sudah Disimpan .......", vbExclamation + vbOKOnly, Me.Caption
End Sub

Private Sub Form_Load()
Dim n As Single
  CenterForm Me
  objData.OpenConnection GetDSN
  nTop1.Value = aCfg(msTopMargin1, 0)
  nTop2.Value = aCfg(msTopMargin2, 0)
  nLeft.Value = aCfg(msLeftMargin, 0)
  nNomor.Value = aCfg(msLebarNomor, 0)
  nTgl.Value = aCfg(msLebarTgl, 0)
  nSandi.Value = aCfg(msLebarSandi, 0)
  nDebet.Value = aCfg(msLebarMutasi, 0)
  nKredit.Value = aCfg(msLebarKredit, 0)
  nSaldo.Value = aCfg(msLebarSaldo, 0)
  nWidth.Value = aCfg(msPaperWidth, 0)
  nHeight.Value = aCfg(msPaperHight, 0)
  
  TabIndex nTop1, n
  TabIndex nTop2, n
  TabIndex nLeft, n
  TabIndex nWidth, n
  TabIndex nHeight, n
  TabIndex nNomor, n
  TabIndex nTgl, n
  TabIndex nSandi, n
  TabIndex nDebet, n
  TabIndex nKredit, n
  TabIndex nSaldo, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

