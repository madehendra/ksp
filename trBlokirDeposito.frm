VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trBlokirDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLOKIR DEPOSITO"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7575
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4035
      Left            =   0
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   7117
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
      Begin BiSAFramProject.BiSAFrame cPesan 
         Height          =   480
         Left            =   4140
         Top             =   3030
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   847
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "REKENING INI SEDANG DIBLOKIR"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   225
            TabIndex        =   0
            Top             =   105
            Width           =   2895
         End
      End
      Begin BiSANumberBoxProject.BiSANumberBox cJangkaWaktu 
         Height          =   330
         Left            =   150
         TabIndex        =   1
         Top             =   2295
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         Caption         =   "Jangka Waktu"
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
      Begin BiSADateProject.BiSADate dValuta 
         Height          =   330
         Left            =   150
         TabIndex        =   2
         Top             =   1200
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Caption         =   "Tgl valuta"
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   465
         Width           =   5160
         _ExtentX        =   9102
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
         FontName        =   "Verdana"
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Nama Deposan"
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   825
         Width           =   5970
         _ExtentX        =   10530
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
         FontName        =   "Verdana"
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Alamat"
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
      Begin BiSATextBoxProject.BiSABrowse cGolonganDeposito 
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   1935
         Width           =   2265
         _ExtentX        =   3995
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
         FontName        =   "Verdana"
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Gol Deposito"
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
      Begin BiSATextBoxProject.BiSATextBox cKetGolDeposito 
         Height          =   330
         Left            =   2430
         TabIndex        =   6
         Top             =   1935
         Width           =   3270
         _ExtentX        =   5768
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
         FontName        =   "Verdana"
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
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
      Begin BiSADateProject.BiSADate dTempo 
         Height          =   330
         Left            =   150
         TabIndex        =   7
         Top             =   1545
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   582
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Caption         =   "Jatuh Tempo"
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame9 
         Height          =   480
         Left            =   1770
         Top             =   3030
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   847
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
         Begin VB.OptionButton optBlokir 
            Caption         =   "&Tidak"
            Height          =   330
            Index           =   1
            Left            =   870
            TabIndex        =   9
            Top             =   75
            Width           =   1020
         End
         Begin VB.OptionButton optBlokir 
            Caption         =   "&Ya"
            Height          =   330
            Index           =   0
            Left            =   105
            TabIndex        =   8
            Top             =   75
            Width           =   1050
         End
      End
      Begin BiSATextBoxProject.BiSATextBox cCabang 
         Height          =   330
         Left            =   150
         TabIndex        =   10
         Top             =   105
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         Text            =   "1234"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         MaxLength       =   4
         Caption         =   "No Rekening"
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   2190
         TabIndex        =   11
         Top             =   105
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         Button          =   -1  'True
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
      Begin BiSATextBoxProject.BiSATextBox cUrut 
         Height          =   330
         Left            =   3030
         TabIndex        =   12
         Top             =   105
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Text            =   "123456"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         MaxLength       =   6
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   3930
         TabIndex        =   13
         Top             =   105
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
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
      Begin BiSANumberBoxProject.BiSANumberBox nNominalDeposito 
         Height          =   330
         Left            =   150
         TabIndex        =   14
         Top             =   2655
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         Caption         =   "Nominal"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   4785
         TabIndex        =   15
         Top             =   90
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Caption         =   "Tgl Blokir"
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   150
         TabIndex        =   16
         Top             =   3555
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   582
         Text            =   "1234"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         Caption         =   "Keterangan"
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
      Begin VB.Label Label4 
         Caption         =   "BLOKIR"
         Height          =   360
         Left            =   150
         TabIndex        =   19
         Top             =   3555
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Bulan"
         Height          =   195
         Left            =   2460
         TabIndex        =   18
         Top             =   2355
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Blokir?"
         Height          =   300
         Left            =   180
         TabIndex        =   17
         Top             =   3180
         Width           =   1515
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   4035
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   1111
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
         Left            =   5250
         TabIndex        =   20
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
         Picture         =   "trBlokirDeposito.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6330
         TabIndex        =   21
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
         Picture         =   "trBlokirDeposito.frx":0416
      End
   End
End
Attribute VB_Name = "trBlokirDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub initvalue()
  dTgl.Value = Date
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
  nNominalDeposito.Value = 0
  cGolonganDeposito.Default
  cKetGolDeposito.Default
  cJangkaWaktu.Default
  cKeterangan.Default
  optBlokir(0).Value = True
  cPesan.Visible = False
End Sub

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.nama,r.Alamat,d.Rekening", "r.Alamat", sisContent, cAlamat.Text, , "r.Nama", _
                              Array("Left Join RegisterNasabah r on r.Kode=d.Kode"))
  cAlamat.Text = cAlamat.Browse(dbData)
  If Not dbData.eof Then
    cCabang.Text = left(GetNull(dbData!Rekening, ""), 2)
    cGolongan.Text = Mid(GetNull(dbData!Rekening, ""), 4, 2)
    cUrut.Text = Mid(GetNull(dbData!Rekening, ""), 7, 6)
    cFrekuensi.Text = Right(GetNull(dbData!Rekening, ""), 2)
    GetData
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "GolonganDeposito", "Kode,Keterangan", "Kode", sisContent, cGolongan.Text)
  cGolongan.Text = cGolongan.Browse(dbData)
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  cGolongan_ButtonClick
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub GetData()
Dim cFields As String
Dim vaJoin
Dim cRekening As String
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cFields = "d.lama as lamaDeposito,d.Rekening,d.Nominaldeposito,d.GolonganDeposito,d.Tgl,d.jthtmp,d.Status,d.StatusBlokir,d.KeteranganBlokir,"
  cFields = cFields & " r.Nama,r.Alamat,r.Telepon,r.Path,"
  cFields = cFields & " b.Keterangan as KeteranganGolDeposito,b.Lama"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode = d.Kode", _
                 "Left Join GolonganDeposito b on b.Kode=d.GolonganDeposito")
  Set dbData = objData.Browse(GetDSN, "Deposito d", cFields, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    cNama.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    dValuta.Value = GetNull(dbData!Tgl, "")
    dTempo.Value = GetNull(dbData!jthtmp, "")
    nNominalDeposito.Value = GetNull(dbData!nominaldeposito, "")
    cGolonganDeposito.Text = GetNull(dbData!GolonganDeposito, "")
    cKetGolDeposito.Text = GetNull(dbData!KeteranganGolDeposito, "")
    cJangkaWaktu.Value = GetNull(dbData!LamaDeposito)
    cKeterangan.Text = GetNull(dbData!KeteranganBlokir, "")
    SetOpt optBlokir, GetNull(dbData!StatusBlokir, "")
    Pesan
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim cRekening As String

  If ValidSaving Then
    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    If MsgBox("Apakah Rekening Benar-benar akan disimpan?", vbYesNo + vbInformation) = vbYes Then
      objData.Edit GetDSN, "Deposito", "Rekening = '" & cRekening & "'", Array("statusBlokir", "KeteranganBlokir"), Array(GetOpt(optBlokir), cKeterangan.Text)
    End If
    initvalue
    cGolongan.SetFocus
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
 
  If Not CheckData(cGolongan.Text, "Invalid kode rekening..!") Then
    ValidSaving = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cUrut.Text, "Invalid kode rekening..!") Then
    ValidSaving = False
    cUrut.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Invalid kode rekening..!") Then
    ValidSaving = False
    cFrekuensi.SetFocus
    Exit Function
  End If
End Function

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cRekening As String
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  Set dbData = objData.Browse(GetDSN, "Deposito", "Rekening,Status", "Rekening", sisAssign, cRekening)
  If Not dbData.eof Then
    If GetNull(dbData!status, "") = "1" Then
      MsgBox "Rekening tersebut sudah Tutup (Cair) !", vbInformation, Me.Caption
      initvalue
      cGolongan.SetFocus
      Exit Sub
    End If
    GetData
    Exit Sub
  End If
  MsgBox "Rekening dengan Nomor : " & cRekening & " Tidak ada. Silahkan Ulangi pengisian.", vbInformation, Me.Caption
  Cancel = True
  cFrekuensi.Default
  cFrekuensi.SetFocus
  Exit Sub
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  initvalue
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex optBlokir(0), n
  TabIndex optBlokir(1), n
  TabIndex cKeterangan, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub optBlokir_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    cKeterangan.Enabled = True
    If Index = 1 Then
      cKeterangan.Default
      cKeterangan.Enabled = False
    End If
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub Pesan()
  cPesan.Visible = False
  If optBlokir(0).Value = True Then
    cPesan.Visible = True
  End If
End Sub
