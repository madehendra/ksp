VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trBlokirTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUKA / BLOKIR SIMPANAN"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8805
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6588
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
      Begin VB.OptionButton optSemua 
         Caption         =   "Buka Blokir"
         Height          =   345
         Index           =   2
         Left            =   4905
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2430
         Width           =   1500
      End
      Begin VB.OptionButton optSemua 
         Caption         =   "Tidak/Sebagian"
         Height          =   345
         Index           =   1
         Left            =   3225
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2430
         Width           =   1500
      End
      Begin VB.OptionButton optSemua 
         Caption         =   "Ya/Semua"
         Height          =   360
         Index           =   0
         Left            =   1980
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2415
         Width           =   1350
      End
      Begin BiSAFramProject.BiSAFrame PesanBlokir 
         Height          =   360
         Left            =   4245
         Top             =   150
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   635
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
            Caption         =   "REKENING INI SEDANG DI BLOKIR"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   165
            TabIndex        =   0
            Top             =   90
            Width           =   4140
         End
      End
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   3825
         TabIndex        =   3
         Top             =   180
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         Text            =   "12"
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   2445
         TabIndex        =   4
         Top             =   180
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
         Text            =   "12"
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
         MaxLength       =   2
         GetPicture      =   1
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
      Begin BiSATextBoxProject.BiSATextBox cCabang 
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   180
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Text            =   "12"
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
         MaxLength       =   2
         Caption         =   "No Rekening"
         CaptionWidth    =   1700
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
         Left            =   2880
         TabIndex        =   6
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   5055
         _ExtentX        =   8916
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
         Caption         =   "Nama Nasabah"
         CaptionWidth    =   1700
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
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Width           =   6015
         _ExtentX        =   10610
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
         Caption         =   "Alamat Nasabah"
         CaptionWidth    =   1700
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
      Begin BiSATextBoxProject.BiSABrowse cNamaGolonganTabungan 
         Height          =   330
         Left            =   2445
         TabIndex        =   9
         Top             =   1650
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         Text            =   "12"
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
      Begin BiSATextBoxProject.BiSATextBox cGolonganTabungan 
         Height          =   330
         Left            =   180
         TabIndex        =   10
         Top             =   1650
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
         Text            =   "12"
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
         MaxLength       =   2
         Appearance      =   0
         Caption         =   "Gol Simpanan"
         CaptionWidth    =   1700
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
      Begin BiSANumberBoxProject.BiSANumberBox nAwal 
         Height          =   330
         Left            =   180
         TabIndex        =   11
         Top             =   2025
         Width           =   3915
         _ExtentX        =   6906
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
         Caption         =   "Saldo Awal"
         CaptionWidth    =   1700
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
      Begin BiSANumberBoxProject.BiSANumberBox nBlokir 
         Height          =   330
         Left            =   180
         TabIndex        =   12
         Top             =   2820
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   582
         MinValue        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Jumlah Blokir"
         CaptionWidth    =   1700
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
         Left            =   180
         TabIndex        =   13
         Top             =   1275
         Width           =   3195
         _ExtentX        =   5636
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
         CaptionWidth    =   1700
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
         Left            =   180
         TabIndex        =   14
         Top             =   3195
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   582
         Text            =   "12"
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
         Caption         =   "Ket Blokir"
         CaptionWidth    =   1700
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
      Begin VB.Label Label1 
         Caption         =   "Blokir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   15
         Top             =   2445
         Width           =   1635
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   3735
      Width           =   8775
      _ExtentX        =   15478
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
         Left            =   6480
         TabIndex        =   16
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
         Picture         =   "trBlokirTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   7560
         TabIndex        =   17
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
         Picture         =   "trBlokirTabungan.frx":0416
      End
   End
End
Attribute VB_Name = "trBlokirTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As Single
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim cDK As String
Dim cKas As String
Dim cNomorRekening As String

Private Sub GetData()
Dim cField As String
Dim vaJoin

  cField = "t.GolonganTabungan,t.JumlahBlokir,t.KeteranganBlokir,r.Nama,r.Alamat,r.Telepon,g.Keterangan as NamaGolongan"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode=t.Kode", _
                 "Left Join golonganTabungan g on g.Kode = t.GolonganTabungan")
  Set dbData = objData.Browse(GetDSN, "Tabungan t", cField, "t.Rekening", sisAssign, cNomorRekening, , , vaJoin)
  If Not dbData.eof Then
    cNama.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    nBlokir.Value = GetNull(dbData!JumlahBlokir)
    cKeterangan.Text = GetNull(dbData!KeteranganBlokir, "")
    nAwal.Value = GetSaldoTabungan1(objData, cNomorRekening)
    cGolonganTabungan.Text = GetNull(dbData!GolonganTabungan, "")
    cNamaGolonganTabungan.Text = GetNull(dbData!NamaGolongan, "")
    If nBlokir.Value > 0 Then
      PesanBlokir.Visible = True
    Else
      PesanBlokir.Visible = False
    End If
   End If
End Sub

Private Sub initvalue()
  dTgl.Value = Date
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
  cGolonganTabungan.Default
  cNamaGolonganTabungan.Default
  nAwal.Value = 0
  nBlokir.Value = 0
  cKeterangan.Default
  PesanBlokir.Visible = False
  cCabang.Text = aCfg(msKodeCabang, "")
  optSemua(0).Value = True
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
 
  If Not CheckData(cGolongan.Text, "Golongan Tabungan Harus Diisi, Silahkan ulangi Pengisian.....!") Then
    ValidSaving = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cUrut.Text, "Nomor Urut Harus Diisi, Silahkan ulangi Pengisian.....!") Then
    ValidSaving = False
    cUrut.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Frekuensi Harus Diisi, Silahkan ulangi Pengisian.....!") Then
    ValidSaving = False
    cFrekuensi.SetFocus
    Exit Function
  End If
End Function

Private Sub cmdSimpan_Click()
Dim nJumlahBlokir As Double
Dim cStatusBlokir As String
Dim Keterangan As String

  
  If optSemua(0).Value = True Then
    cStatusBlokir = "1"
    Keterangan = cKeterangan.Text
  ElseIf optSemua(1).Value = True Then
    cStatusBlokir = "1"
    Keterangan = cKeterangan.Text
  ElseIf optSemua(2).Value = True Then
    cStatusBlokir = "0"
    Keterangan = ""
    nBlokir.Value = 0
  End If
  
  If ValidSaving() Then
    If MsgBox("Data benar-benar akan disimpan ?", vbYesNo + vbInformation) = vbYes Then
      objData.Edit GetDSN, "Tabungan", "Rekening = '" & cNomorRekening & "'", Array("StatusBlokir", "JumlahBlokir", "KeteranganBlokir"), Array(cStatusBlokir, nBlokir.Value, Keterangan)
    End If
  End If
  
  initvalue
  cCabang.SetFocus
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
  cNomorRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  Set dbData = objData.Browse(GetDSN, "tabungan", "Rekening,Close", "Rekening", sisAssign, cNomorRekening)
  If Not dbData.eof Then
    If GetNull(dbData!Close, "") = "1" Then
      MsgBox "Maaf, Nomor Rekening : " & cNomorRekening & " Sudah DITutup. Silahkan mengulangi pengisian!", vbOKOnly, "Blokir Tabungan"
      Cancel = True
      cFrekuensi.Default
      cFrekuensi.SetFocus
      Exit Sub
    End If
    GetData
  Else
    MsgBox "Rekening dengan nomor: " & cNomorRekening & " Tidak ada. Silahkan mengulangi pengisian !", vbOKOnly, "Blokir Tabungan"
    initvalue
    Cancel = True
    cCabang.SetFocus
    Exit Sub
  End If
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
  TabIndex optSemua(0), n
  TabIndex optSemua(1), n
  TabIndex optSemua(2), n
  TabIndex nBlokir, n
  TabIndex cKeterangan, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub nBlokir_Validate(Cancel As Boolean)
  If nBlokir.Value < 0 Then
    nBlokir.Value = 0
    Cancel = True
  End If
End Sub

Private Sub optSemua_Click(Index As Integer)
    If Index = 0 Then
      nBlokir.Value = nAwal.Value
      nBlokir.Enabled = False
    ElseIf Index = 2 Then
      nBlokir.Value = 0
      nBlokir.Enabled = False
    ElseIf Index = 1 Then
      nBlokir.Enabled = True
    End If
End Sub

Private Sub optSemua_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Index = 0 Then
      nBlokir.Value = nAwal.Value
      nBlokir.Enabled = False
      cKeterangan.SetFocus
    ElseIf Index = 1 Then
      nBlokir.Enabled = True
      nBlokir.SetFocus
    ElseIf Index = 2 Then
      nBlokir.Value = 0
      nBlokir.Enabled = False
      cKeterangan.Default
      cKeterangan.SetFocus
    End If
  End If
End Sub

