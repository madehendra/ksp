VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KONFIGURASI SIMPANAN"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   9330
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2475
      Left            =   0
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   4366
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
      Begin BiSATextBoxProject.BiSATextBox cNamaPajak 
         Height          =   330
         Left            =   4605
         TabIndex        =   0
         Top             =   135
         Width           =   4620
         _ExtentX        =   8149
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
      Begin BiSATextBoxProject.BiSABrowse cKodePajak 
         Height          =   330
         Left            =   345
         TabIndex        =   1
         Top             =   135
         Width           =   4260
         _ExtentX        =   7514
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
         Caption         =   "KODE TRANSAKSI PAJAK BUNGA"
         CaptionWidth    =   3000
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
      Begin BiSATextBoxProject.BiSABrowse cKodeSetoranTunai 
         Height          =   330
         Left            =   345
         TabIndex        =   2
         Top             =   510
         Width           =   4260
         _ExtentX        =   7514
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
         Caption         =   "KODE SETORAN TUNAI"
         CaptionWidth    =   3000
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
      Begin BiSATextBoxProject.BiSATextBox cNamaSetoranTunai 
         Height          =   330
         Left            =   4605
         TabIndex        =   3
         Top             =   510
         Width           =   4620
         _ExtentX        =   8149
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
      Begin BiSATextBoxProject.BiSABrowse cKodePenarikanTunai 
         Height          =   330
         Left            =   345
         TabIndex        =   4
         Top             =   885
         Width           =   4260
         _ExtentX        =   7514
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
         Caption         =   "KODE PENARIKAN TUNAI"
         CaptionWidth    =   3000
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
      Begin BiSATextBoxProject.BiSATextBox cNamaPenarikanTunai 
         Height          =   330
         Left            =   4605
         TabIndex        =   5
         Top             =   885
         Width           =   4620
         _ExtentX        =   8149
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
      Begin BiSATextBoxProject.BiSABrowse cKodePembulatankas 
         Height          =   330
         Left            =   345
         TabIndex        =   6
         Top             =   1245
         Width           =   4260
         _ExtentX        =   7514
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
         Caption         =   "KODE PEMBULATAN KAS"
         CaptionWidth    =   3000
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
      Begin BiSATextBoxProject.BiSATextBox cNamaPembulatanKas 
         Height          =   330
         Left            =   4605
         TabIndex        =   7
         Top             =   1245
         Width           =   4620
         _ExtentX        =   8149
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
      Begin BiSATextBoxProject.BiSABrowse cKodeAdministrasi 
         Height          =   330
         Left            =   345
         TabIndex        =   8
         Top             =   1620
         Width           =   4260
         _ExtentX        =   7514
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
         Caption         =   "KODE ADM. TUTUP SIMPANAN"
         CaptionWidth    =   3000
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
      Begin BiSATextBoxProject.BiSATextBox cNamaKodeAdministrasi 
         Height          =   330
         Left            =   4605
         TabIndex        =   9
         Top             =   1620
         Width           =   4620
         _ExtentX        =   8149
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
      Begin BiSATextBoxProject.BiSABrowse cKodeBunga 
         Height          =   330
         Left            =   345
         TabIndex        =   10
         Top             =   1995
         Width           =   4260
         _ExtentX        =   7514
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
         Caption         =   "KODE TRANSAKSI BUNGA"
         CaptionWidth    =   3000
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
      Begin BiSATextBoxProject.BiSATextBox cNamaBunga 
         Height          =   330
         Left            =   4605
         TabIndex        =   11
         Top             =   1995
         Width           =   4620
         _ExtentX        =   8149
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   2460
      Width           =   9330
      _ExtentX        =   16457
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
         Left            =   7020
         TabIndex        =   12
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
         Picture         =   "cfgTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   8100
         TabIndex        =   13
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
         Picture         =   "cfgTabungan.frx":0416
      End
   End
End
Attribute VB_Name = "cfgTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub cKodeAdministrasi_ButtonClick()
  GetKodeTransaksi cNamaKodeAdministrasi, cKodeAdministrasi
End Sub

Private Sub cKodeAdministrasi_Validate(Cancel As Boolean)
  cKodeAdministrasi_ButtonClick
End Sub

Private Function GetKodeTransaksi(cNM As Object, cKD As Object)
  Set dbData = objData.Pick(GetDSN, "KodeTransaksi", "Kode", cKD, "Kode,Keterangan", " and DK <> 'M'")
  If Not dbData.eof Then
    cNM.Text = GetNull(dbData!Keterangan)
  End If
End Function

Private Sub cKodePajak_ButtonClick()
  GetKodeTransaksi cNamaPajak, cKodePajak
End Sub

Private Sub cKodePajak_Validate(Cancel As Boolean)
  If cKodePajak.LastKey = 13 Then
    cKodePajak_ButtonClick
  End If
End Sub

Private Sub cKodePembulatankas_ButtonClick()
  GetKodeTransaksi cNamaPembulatanKas, cKodePembulatankas
End Sub

Private Sub cKodePembulatankas_Validate(Cancel As Boolean)
  If cKodePembulatankas.LastKey = 13 Then
    cKodePembulatankas_ButtonClick
  End If
End Sub


Private Sub cKodeSetoranTunai_ButtonClick()
  GetKodeTransaksi cNamaSetoranTunai, cKodeSetoranTunai
End Sub

Private Sub cKodesetoranTunai_Validate(Cancel As Boolean)
  If cKodeSetoranTunai.LastKey = 13 Then
    cKodeSetoranTunai_ButtonClick
  End If
End Sub

Private Sub cKodePenarikanTunai_ButtonClick()
  GetKodeTransaksi cNamaPenarikanTunai, cKodePenarikanTunai
End Sub

Private Sub cKodePenarikanTunai_Validate(Cancel As Boolean)
  If cKodePenarikanTunai.LastKey = 13 Then
    cKodePenarikanTunai_ButtonClick
  End If
End Sub


Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cKodeBunga_ButtonClick()
  GetKodeTransaksi cNamaBunga, cKodeBunga
End Sub

Private Sub cKodeBunga_Validate(Cancel As Boolean)
  If cKodeBunga.LastKey = 13 Then
    cKodeBunga_ButtonClick
  End If
End Sub

Private Sub cmdSimpan_Click()
  UpdCfg msKodeAdministrasi, cKodeAdministrasi.Text, objData
  UpdCfg msKodePembulatankas, cKodePembulatankas.Text, objData
  UpdCfg msKodePenarikanTunai, cKodePenarikanTunai.Text, objData
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  
  TabIndex cKodePajak, n
  TabIndex cKodeSetoranTunai, n
  TabIndex cKodePenarikanTunai, n
  TabIndex cKodePembulatankas, n
  TabIndex cKodeAdministrasi, n
  TabIndex cKodeBunga, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  cKodePajak.Text = aCfg(msKodePajakBagiHasil, "")
  cNamaPajak.Text = GetKeterangan(cKodePajak.Text)
  cKodeSetoranTunai.Text = aCfg(msKodeSetoranTunai, "")
  cNamaSetoranTunai.Text = GetKeterangan(cKodeSetoranTunai.Text)
  cKodePenarikanTunai.Text = aCfg(msKodePenarikanTunai, "")
  cNamaPenarikanTunai.Text = GetKeterangan(cKodePenarikanTunai.Text)
  cKodePembulatankas.Text = aCfg(msKodePembulatankas, "")
  cNamaPembulatanKas.Text = GetKeterangan(cKodePembulatankas.Text)
  cKodeAdministrasi.Text = aCfg(msKodeAdministrasi, "")
  cNamaKodeAdministrasi.Text = GetKeterangan(cKodeAdministrasi.Text)
  cKodeBunga.Text = aCfg(mskodebagihasil)
  cNamaBunga.Text = GetKeterangan(cKodeBunga.Text)
End Sub

Private Function GetKeterangan(ByVal cKode As String) As String
  Set dbData = objData.Browse(GetDSN, "KodeTransaksi", "Keterangan", "Kode", sisAssign, cKode)
  If Not dbData.eof Then
    GetKeterangan = GetNull(dbData!Keterangan, "")
  End If
End Function
