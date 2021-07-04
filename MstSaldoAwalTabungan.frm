VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form MstSaldoAwalTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER SALDO AWAL SIMPANAN"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7770
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   3585
      Left            =   0
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   6324
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   2355
         TabIndex        =   18
         Top             =   120
         Width           =   810
         _ExtentX        =   1429
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
         Left            =   150
         TabIndex        =   0
         Top             =   3135
         Width           =   3735
         _ExtentX        =   6588
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
         Caption         =   "SALDO AWAL"
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
      Begin BiSATextBoxProject.BiSATextBox cRegisterNasabah 
         Height          =   330
         Left            =   150
         TabIndex        =   1
         Top             =   1620
         Width           =   3225
         _ExtentX        =   5689
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
         Caption         =   "NO. REGISTER"
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
         Left            =   150
         TabIndex        =   2
         Top             =   1245
         Width           =   3165
         _ExtentX        =   5583
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
         Caption         =   "TGL PEMBUKAAN"
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   495
         Width           =   5505
         _ExtentX        =   9710
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
         Caption         =   "NAMA NASABAH"
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   4035
         TabIndex        =   4
         Top             =   120
         Width           =   420
         _ExtentX        =   741
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
      Begin BiSATextBoxProject.BiSATextBox cUrut 
         Height          =   330
         Left            =   3165
         TabIndex        =   5
         Top             =   120
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         Text            =   "123456"
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
      Begin BiSATextBoxProject.BiSATextBox cCabang 
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   120
         Width           =   2190
         _ExtentX        =   3863
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
         MaxLength       =   2
         Caption         =   "NO. REKENING"
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
         Left            =   135
         TabIndex        =   7
         Top             =   870
         Width           =   6870
         _ExtentX        =   12118
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
         Caption         =   "ALAMAT NASABAH"
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
      Begin BiSATextBoxProject.BiSATextBox cTelepon 
         Height          =   330
         Left            =   150
         TabIndex        =   8
         Top             =   1995
         Width           =   4605
         _ExtentX        =   8123
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
         Caption         =   "NO. TELEPON"
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
      Begin BiSATextBoxProject.BiSATextBox cGolonganTabungan 
         Height          =   330
         Left            =   150
         TabIndex        =   9
         Top             =   2370
         Width           =   2370
         _ExtentX        =   4180
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
         BackColor       =   12632256
         Enabled         =   0   'False
         MaxLength       =   4
         Appearance      =   0
         Caption         =   "GOL. SIMPANAN"
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
      Begin BiSATextBoxProject.BiSATextBox cKetGolTabungan 
         Height          =   330
         Left            =   2565
         TabIndex        =   10
         Top             =   2370
         Width           =   3255
         _ExtentX        =   5741
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
      Begin BiSADateProject.BiSADate dLastUpdate 
         Height          =   330
         Left            =   150
         TabIndex        =   11
         Top             =   2745
         Width           =   3165
         _ExtentX        =   5583
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
         Caption         =   "TGL AKHIR TRANS"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   3570
      Width           =   7725
      _ExtentX        =   13626
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
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2220
         TabIndex        =   12
         Top             =   105
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "    &Delete"
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
         Picture         =   "MstSaldoAwalTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
         TabIndex        =   13
         Top             =   105
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   767
         Caption         =   ""
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
         Picture         =   "MstSaldoAwalTabungan.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   5355
         TabIndex        =   14
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
         Picture         =   "MstSaldoAwalTabungan.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   15
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Caption         =   "  &Edit"
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
         Picture         =   "MstSaldoAwalTabungan.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   16
         Top             =   105
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         Caption         =   "  &Add"
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
         Picture         =   "MstSaldoAwalTabungan.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6435
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
         Picture         =   "MstSaldoAwalTabungan.frx":0C9A
      End
   End
End
Attribute VB_Name = "MstSaldoAwalTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim cSQL As String
Dim nPos As SisPos
Dim lEdit As Boolean
Dim vaArray As New XArrayDB

Private Sub GetMemory()
Dim cField As String
Dim vaJoin
Dim cRekening As String
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cField = "d.Rekening,d.Tgl,d.GolonganTabungan, d.Kode,d.awal,d.Close,"
  cField = cField & " r.Nama, r.Alamat, r.telepon,"
  cField = cField & " b.Keterangan as KetGolTabungan"
  vaJoin = Array(" Left Join RegisterNasabah r on r.Kode=d.Kode", _
                 " Left Join GolonganTabungan b on b.Kode=d.GolonganTabungan")
  
  Set dbData = objData.Browse(GetDSN, "Tabungan d", cField, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    If GetNull(dbData!Close) = "1" Then
      MsgBox "Rekening Tersebut sudah tutup.", vbOKOnly, "Saldo Awal Tabungan"
      GetEdit False
      initvalue
      Exit Sub
    End If
    
    dTgl.Value = GetNull(dbData!Tgl)
    cRegisterNasabah.Text = GetNull(dbData!Kode)
    cNama.Text = GetNull(dbData!nama)
    cAlamat.Text = GetNull(dbData!alamat)
    cTelepon.Text = GetNull(dbData!Telepon)
    cGolonganTabungan.Text = GetNull(dbData!GolonganTabungan)
    cKetGolTabungan.Text = GetNull(dbData!KetGolTabungan)
    nSaldo.Value = GetSA(cRekening)
  End If
End Sub

Private Function GetSA(ByVal cRek As String) As Double
  GetSA = 0
  Set dbData = objData.Browse(GetDSN, "MutasiTabungan", "Jumlah", "Faktur", sisAssign, "SAT0000001", "And Rekening='" & cRek & "'")
  If Not dbData.eof Then
    GetSA = GetNull(dbData!Jumlah)
  End If
End Function

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cRekening As String

  If cFrekuensi.LastKey = 13 Then
    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    Set dbData = objData.Browse(GetDSN, "Tabungan", "Rekening", "Rekening", sisAssign, cRekening)
    If Not dbData.eof Then
      GetMemory
      If nPos = Delete Then DeleteData cRekening
      Exit Sub
    End If
    MsgBox "Data Tidak Ada. Silahkan Ulangi pengisian.", vbOKOnly, "Saldo Awal Tabungan"
    Cancel = True
    cFrekuensi.Default
    cFrekuensi.SetFocus
    Exit Sub
  End If
End Sub

Private Sub DeleteData(ByVal cRek As String)
  If MsgBox("Akan menghapus saldo awal tabungan ?", vbYesNo) = vbYes Then
    objData.Delete GetDSN, "MutasiTabungan", "Faktur", sisAssign, "SAT0000001", "And Rekening='" & cRek & "'"
    MsgBox "Data sudah terhapus....", vbInformation
  End If
  initvalue
  GetEdit False
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganTabungan", "Kode", cGolongan, "Kode,Keterangan")
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  cCabang.SetFocus
End Sub

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  initvalue
  cCabang.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  initvalue
  cCabang.SetFocus
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    initvalue
    GetEdit False
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue
Dim cRekening As String

  If ValidSaving() Then
    If MsgBox("Apakah Data Benar-benar Sudah Valid dan akan disimpan ?", vbYesNo) = vbYes Then
      cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
      
      'Hapus dulu di Mutasitabungan
      objData.Delete GetDSN, "MutasiTabungan", "Faktur", sisAssign, "SAT0000001", "And Rekening='" & cRekening & "'"
      
      'Simpan di MutasiTabungan
      vaField = Array("Faktur", "Tgl", "Rekening", "Jumlah", "DK", "Keterangan", "UserName", "DateTime")
      vaValue = Array("SAT0000001", dLastUpdate.Value, cRekening, nSaldo.Value, "K", "SALDO AWAL TABUNGAN", cusername, SNow)
      objData.Add GetDSN, "MutasiTabungan", vaField, vaValue
      
      initvalue
      GetEdit False
    Else
      GetEdit True
      nSaldo.SetFocus
      Exit Sub
    End If
  End If
End Sub

Static Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cCabang.Text, "Nomor Rekening tidak valid, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cCabang.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cGolongan.Text, "Nomor Rekening tidak valid, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cUrut.Text, "Nomor Rekening tidak valid, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cUrut.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Frekuensi Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cFrekuensi.SetFocus
    Exit Function
  End If
  
  If nSaldo.Value = 0 Then
    MsgBox "Saldo Tidak Boleh Kosong, Silahkan Mengulangi Pengisian !", vbOKOnly
    ValidSaving = False
    nSaldo.SetFocus
    Exit Function
  End If
End Function

Private Sub initvalue()
  cGolongan.Default
  cFrekuensi.Default
  cUrut.Default
  dTgl.Value = ""
  cRegisterNasabah.Default
  cNama.Default
  cAlamat.Default
  cTelepon.Default
  cGolonganTabungan.Default
  cKetGolTabungan.Default
  nSaldo.Value = 0
  cCabang.Text = aCfg(msKodeCabang)
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Tabungan t", "r.Nama,r.Alamat,t.Rekening", "r.Nama", sisContent, cNama.Text, "And t.Close <>'1' And r.Alamat like '" & cAlamat.Text & "%'", "r.nama,r.Alamat", Array("Left Join RegisterNasabah r on r.Kode = t.Kode"))
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.eof Then
    GetRekening GetNull(dbData!Rekening)
    GetMemory
  End If
End Sub

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Tabungan t", "r.Alamat,r.Nama,t.Rekening", "r.Alamat", sisContent, cAlamat.Text, "And t.Close <>'1' And r.Nama like '" & cNama.Text & "%'", "r.nama,r.Alamat", Array("Left Join RegisterNasabah r on r.Kode = t.Kode"))
  cAlamat.Text = cAlamat.Browse(dbData)
  If Not dbData.eof Then
    GetRekening GetNull(dbData!Rekening)
    GetMemory
  End If
End Sub

Private Sub GetRekening(ByVal cRekening As String)
  cCabang.Text = left(cRekening, 2)
  cGolongan.Text = Mid(cRekening, 4, 2)
  cUrut.Text = Mid(cRekening, 7, 6)
  cFrekuensi.Text = Right(cRekening, 2)
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  initvalue
  GetEdit False
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex dLastUpdate, n
  TabIndex nSaldo, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub
