VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form frmTutupBuku 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TUTUP BUKU"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8820
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1245
      Left            =   0
      Top             =   4035
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2196
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
      Begin BiSADateProject.BiSADate dAwal 
         Height          =   330
         Left            =   165
         TabIndex        =   0
         Top             =   450
         Width           =   3975
         _ExtentX        =   7011
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Caption         =   "ANTARA TGL"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cPeriode 
         Height          =   330
         Left            =   150
         TabIndex        =   1
         Top             =   90
         Width           =   3225
         _ExtentX        =   5689
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   4
         Caption         =   "Periode Yg Ditutup"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaPeriode 
         Height          =   330
         Left            =   3405
         TabIndex        =   2
         Top             =   90
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
         Text            =   "1234567890123456789012345678901234567890"
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   40
         CaptionWidth    =   1300
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
         Left            =   4155
         TabIndex        =   3
         Top             =   450
         Width           =   1980
         _ExtentX        =   3493
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         Enabled         =   0   'False
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
      Begin BiSATextBoxProject.BiSATextBox cProses 
         Height          =   330
         Left            =   165
         TabIndex        =   7
         Top             =   810
         Width           =   3975
         _ExtentX        =   7011
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
         MaxLength       =   6
         Caption         =   "Ketik ""PROSES"" utk Lanjut"
         CaptionWidth    =   2500
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4020
      Left            =   0
      Top             =   15
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7091
      Caption         =   "PERHATIAN"
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
      Begin VB.TextBox cPerhatian 
         Height          =   3720
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmTutupBuku.frx":0000
         Top             =   210
         Width           =   8580
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5265
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   7605
         TabIndex        =   4
         Top             =   120
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
         Picture         =   "frmTutupBuku.frx":0006
      End
      Begin BiSAButtonProject.BiSAButton cmdProses 
         Height          =   435
         Left            =   6525
         TabIndex        =   5
         Top             =   120
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         Caption         =   "    &Proses"
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
         Picture         =   "frmTutupBuku.frx":00AC
      End
   End
End
Attribute VB_Name = "frmTutupBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdProses_Click()
  If MsgBox("Proses Dilanjutkan ?", vbQuestion + vbYesNo) = vbYes Then
    If ValidProses() Then
      objData.Edit GetDSN, "Periode", "Kode = '" & cPeriode.Text & "'", Array("Status"), Array("1")
      MsgBox "Proses Sudah Selesai, Transaksi Harian bisa Dilanjutkan", vbInformation
    End If
  End If
End Sub

Private Function ValidProses() As Boolean
  ValidProses = True
  
  If cProses.Text <> "PROSES" Then
    MsgBox "Kata Kunci Salah, Proses Tidak Bisa Dilanjutkan", vbExclamation
    ValidProses = False
    Exit Function
  End If
End Function

Private Sub cPerhatian_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub cPerhatian_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Form_Activate()
  ' Periksa apakah ada Periode yang aktif
  ' Jika Tidak ada maka Proses Tidak Bisa Dilanjutkan
  Set dbData = objData.Browse(GetDSN, "Periode", , "Status", sisAssign, "0", , "Kode", , 0, 1)
  If Not dbData.eof Then
    cPeriode.Text = dbData!Kode
    cNamaPeriode.Text = dbData!Keterangan
    dAwal.Value = dbData!Awal
    dAkhir.Value = dbData!akhir
  Else
    MsgBox "Tidak Ada Periode Akuntansi yang Aktif, Proses Tutup Buku Tidak Bisa Dilanjutkan ... !", vbExclamation
    Unload Me
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
Dim cMsg As String

  CenterForm Me, True
  
  TabIndex cProses, n
  TabIndex cmdProses, n
  TabIndex cmdKeluar, n
  
  cMsg = "Perhatian : " & vbCrLf & vbCrLf
  cMsg = cMsg & "Proses Tutup Buku adalah Proses dimana system akan menutup transaksi Periode Berjalan. "
  cMsg = cMsg & "Setelah dilakukannya proses tutup buku bisa menyebabkan hal-hal sbb :" & vbCrLf
  cMsg = cMsg & "1. Tanggal Transaksi akan di Tutup." & vbCrLf
  cMsg = cMsg & "2. Transaksi yang terjadi pada periode yang telah ditutup tidak dapat dilakukan : " & vbCrLf
  cMsg = cMsg & "    a. Tambah " & vbCrLf
  cMsg = cMsg & "    b. Edit " & vbCrLf
  cMsg = cMsg & "    c. Hapus " & vbCrLf
  cMsg = cMsg & "3. Apabila Terjadi kesalahan Harus dikoreksi periode berikutnya." & vbCrLf & vbCrLf
  cMsg = cMsg & "Sebelum melanjutkan Proses Tutup Buku pastikan anda telah melakukan hal-hal sbb :" & vbCrLf
  cMsg = cMsg & "1. Pastikan Semua Transaksi hari ini telah masuk semua." & vbCrLf
  cMsg = cMsg & "2. Periksa kebenaran Transaksi untuk masalah penyimpanan dan Pengakuannya." & vbCrLf
  cMsg = cMsg & "3. Lanjutkan Proses ini jika 2 Langkat diatas sudah anda lakukan." & vbCrLf
  
  cPerhatian.Text = cMsg
End Sub

