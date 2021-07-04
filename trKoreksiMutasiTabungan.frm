VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trKoreksiMutasiTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Koreksi Mutasi Simpanan"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7050
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   2490
      Left            =   0
      Top             =   1995
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   4392
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   450
         TabIndex        =   17
         Top             =   1995
         Width           =   6120
         _ExtentX        =   10795
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
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   450
         TabIndex        =   15
         Top             =   1230
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "Jumlah Mutasi"
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
         Left            =   450
         TabIndex        =   14
         Top             =   870
         Width           =   2970
         _ExtentX        =   5239
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
         Caption         =   "Tanggal Mutasi"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaKodeTransaksi 
         Height          =   330
         Left            =   2715
         TabIndex        =   13
         Top             =   510
         Width           =   3780
         _ExtentX        =   6668
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
      Begin BiSATextBoxProject.BiSATextBox cKodeTransaksi 
         Height          =   330
         Left            =   450
         TabIndex        =   12
         Top             =   510
         Width           =   2250
         _ExtentX        =   3969
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
         Caption         =   "Kode Transaksi"
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   450
         TabIndex        =   11
         Top             =   150
         Width           =   4665
         _ExtentX        =   8229
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
         Button          =   -1  'True
         Caption         =   "No Faktur"
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
      Begin BiSATextBoxProject.BiSATextBox cUser 
         Height          =   330
         Left            =   450
         TabIndex        =   16
         Top             =   1605
         Width           =   2940
         _ExtentX        =   5186
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
         Caption         =   "User Name"
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
      Begin BiSATextBoxProject.BiSATextBox cFullName 
         Height          =   330
         Left            =   3405
         TabIndex        =   18
         Top             =   1605
         Width           =   3195
         _ExtentX        =   5636
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1995
      Left            =   0
      Top             =   0
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   3519
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   4200
         TabIndex        =   0
         Top             =   105
         Width           =   435
         _ExtentX        =   767
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
         Left            =   2460
         TabIndex        =   1
         Top             =   105
         Width           =   810
         _ExtentX        =   1429
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
      Begin BiSATextBoxProject.BiSATextBox cCabang 
         Height          =   330
         Left            =   375
         TabIndex        =   2
         Top             =   105
         Width           =   2055
         _ExtentX        =   3625
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
      Begin BiSATextBoxProject.BiSATextBox cUrut 
         Height          =   330
         Left            =   3285
         TabIndex        =   3
         Top             =   105
         Width           =   900
         _ExtentX        =   1588
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
         Left            =   375
         TabIndex        =   4
         Top             =   465
         Width           =   5400
         _ExtentX        =   9525
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
         Left            =   375
         TabIndex        =   5
         Top             =   825
         Width           =   5400
         _ExtentX        =   9525
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
      Begin BiSADateProject.BiSADate dAwal 
         Height          =   330
         Left            =   375
         TabIndex        =   6
         Top             =   1185
         Width           =   2925
         _ExtentX        =   5159
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
         Caption         =   "Antara Tgl"
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
      Begin BiSADateProject.BiSADate dAkhir 
         Height          =   330
         Left            =   3525
         TabIndex        =   7
         Top             =   1185
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   "S.D"
         CaptionWidth    =   500
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
      Begin BiSANumberBoxProject.BiSANumberBox nAkhir 
         Height          =   330
         Left            =   375
         TabIndex        =   8
         Top             =   1560
         Width           =   3465
         _ExtentX        =   6112
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
         Caption         =   "Saldo Akhir"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   4485
      Width           =   7020
      _ExtentX        =   12383
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
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   4725
         TabIndex        =   9
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Caption         =   "  &Save"
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
         Picture         =   "trKoreksiMutasiTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5775
         TabIndex        =   10
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
         Picture         =   "trKoreksiMutasiTabungan.frx":012C
      End
   End
End
Attribute VB_Name = "trKoreksiMutasiTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim xArray As New XArrayDB
Dim vaView As New XArrayDB
Dim cSQL As String
Dim cRekening As String

Private Sub cFaktur_ButtonClick()
Dim cWhere As String
  
  cWhere = "And m.Tgl >= '" & Format(dAwal.Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & "And m.Tgl <='" & Format(dAkhir.Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & "And m.Rekening = '" & cRekening & "'"
  Set dbData = objData.Browse(GetDSN, "MutasiTabungan m", "m.Faktur,m.Tgl,m.KodeTransaksi,m.Keterangan as KeteranganMutasi,k.Keterangan as NamaKodeTransaksi,m.Jumlah,m.UserName,u.Fullname", "m.Faktur", sisContent, cFaktur.Text, cWhere, "m.Tgl,m.ID", _
                              Array("Left join kodetransaksi k on k.Kode=m.KodeTransaksi", _
                                    "Left Join userName u on u.username=m.username"))
  cFaktur.Text = cFaktur.Browse(dbData)
  If Not dbData.eof Then
    cKodeTransaksi.Text = GetNull(dbData!KodeTransaksi)
    cNamaKodeTransaksi.Text = GetNull(dbData!NamaKodeTransaksi)
    dTgl.Value = GetNull(dbData!Tgl)
    nJumlah.Value = GetNull(dbData!Jumlah)
    cUser.Text = GetNull(dbData!UserName)
    cFullName.Text = GetNull(dbData!FullName)
    cKeterangan.Text = GetNull(dbData!KeteranganMutasi)
  End If
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  Set dbData = objData.Browse(GetDSN, "Tabungan t", "t.Rekening,r.Nama,r.Alamat,t.Close", "t.rekening", sisAssign, cRekening, , , _
                              Array("left Join RegisterNasabah r on r.Kode=t.Kode"))
  If dbData.eof Then
    MsgBox "Data tidak ada.", vbInformation
    Cancel = True
    cGolongan.SetFocus
    initvalue
    Exit Sub
  End If
  GetData
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "GolonganTabungan", "Kode,Keterangan", "Kode", sisContent, cGolongan.Text)
  cGolongan.Text = cGolongan.Browse(dbData)
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  cGolongan_ButtonClick
End Sub

Private Sub cmdEdit_Click()
Dim cDK As String
Dim cRekeningJurnal As String
Dim nDebet As Double
Dim nKredit As Double

  If ValidSaving Then
    If MsgBox("Data Benar-benar di Edit/Koreksi ?", vbQuestion + vbYesNo) = vbYes Then
          objData.Delete GetDSN, "MutasiTabungan", "Faktur", sisAssign, cFaktur.Text, "And Rekening='" & cRekening & "'"
          objData.Delete GetDSN, "BukuBesar", "Faktur", sisAssign, cFaktur.Text
          GetRekening cKodeTransaksi.Text, cDK, cRekeningJurnal
          If cDK = "D" Then
            nDebet = nJumlah.Value
            nKredit = 0
          ElseIf cDK = "K" Then
            nKredit = nJumlah.Value
            nDebet = 0
          End If
          UpdMutasiTabungan objData, cKodeTransaksi.Text, cFaktur.Text, dTgl.Value, cRekening, nJumlah.Value, False, cKeterangan.Text, True, cDK, cRekeningJurnal
      MsgBox "Data sudah diEdit/Koreksi", vbInformation
      initvalue
      Exit Sub
    End If
  End If
End Sub

Private Sub GetRekening(ByVal cKode As String, ByRef cDK As String, ByRef cRekeningJurnal As String)
  Set dbData = objData.Browse(GetDSN, "KodeTransaksi", "Rekening,DK", "Kode", sisAssign, cKode)
  If Not dbData.eof Then
    cDK = GetNull(dbData!DK)
    cRekeningJurnal = GetNull(dbData!Rekening)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
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
  
  If Not CheckData(cFaktur.Text, "Faktur harus diisi!") Then
    ValidSaving = False
    cFaktur.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKeterangan.Text, "Keterangan haris diisi..!") Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
  
  If nJumlah.Value <= 0 Then
    MsgBox "Jumlah tidak valid", vbInformation + vbOKOnly
    ValidSaving = False
    nJumlah.SetFocus
    Exit Function
  End If
End Function

Private Sub GetData()
  cNama.Text = GetNull(dbData!nama, "")
  cAlamat.Text = GetNull(dbData!alamat, "")
  nAkhir.Value = GetSaldoTab(objData, cRekening, Date)
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub dAkhir_Validate(Cancel As Boolean)
  If Not IsInPeriod(dAkhir.Value) Then
    Cancel = True
    dAkhir.SetFocus
  End If
End Sub

Private Sub dAwal_Validate(Cancel As Boolean)
  If Not IsInPeriod(dAwal.Value) Then
    Cancel = True
    dAwal.SetFocus
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  Me.Top = 0
  initvalue
  dAwal.Value = Date
  dAkhir.Value = Date
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex dAwal, n
  TabIndex dAkhir, n
  TabIndex cFaktur, n
  TabIndex dTgl, n
  TabIndex nJumlah, n
  TabIndex cKeterangan, n
  TabIndex cmdEdit, n
  TabIndex cmdKeluar, n
End Sub

Private Sub initvalue()
  dAwal.Value = Date
  dAkhir.Value = Date
  cCabang.Text = aCfg(msKodeCabang, "")
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
  nAkhir.Value = 0
  cFaktur.Default
  cKodeTransaksi.Default
  cNamaKodeTransaksi.Default
  dTgl.Value = Date
  nJumlah.Value = 0
  cUser.Default
  cFullName.Default
  cKeterangan.Default
End Sub
