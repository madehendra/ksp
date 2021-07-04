VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form trAdministrasiSimpanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrasi Simpanan"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5880
   Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   2130
      Width           =   2520
      _ExtentX        =   4445
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
      Caption         =   "Jumlah"
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
   Begin BiSATextBoxProject.BiSATextBox cRekDebet 
      Height          =   330
      Left            =   2595
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   840
      Width           =   2100
      _ExtentX        =   3704
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
      Enabled         =   0   'False
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
   Begin BiSATextBoxProject.BiSABrowse cGolonganSimpanan 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   840
      Width           =   2475
      _ExtentX        =   4366
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
      Caption         =   "Golongan"
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
   Begin VB.OptionButton optProses 
      Caption         =   "Saldo +"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2535
      TabIndex        =   7
      Top             =   2565
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.OptionButton optProses 
      Caption         =   "Semuanya"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   1215
      TabIndex        =   6
      Top             =   2565
      Visible         =   0   'False
      Width           =   1215
   End
   Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   2910
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   556
      Picture         =   "trAdministrasiSimpanan.frx":0000
      ForeColor       =   0
      BarPicture      =   "trAdministrasiSimpanan.frx":001C
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
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   465
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Tanggal"
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
   Begin BiSATextBoxProject.BiSATextBox cKode 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   2430
      _ExtentX        =   4286
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
      MaxLength       =   8
      Caption         =   "ID"
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
   Begin BiSAButtonProject.BiSAButton cmdSimpan 
      Height          =   435
      Left            =   3525
      TabIndex        =   2
      Top             =   2790
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
      Picture         =   "trAdministrasiSimpanan.frx":0038
   End
   Begin BiSAButtonProject.BiSAButton cmdKeluar 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4605
      TabIndex        =   3
      Top             =   2790
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
      Picture         =   "trAdministrasiSimpanan.frx":05D2
   End
   Begin BiSATextBoxProject.BiSABrowse cKodeTransaksi 
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   1215
      Width           =   1980
      _ExtentX        =   3493
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
      Caption         =   "Golongan"
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
   Begin BiSATextBoxProject.BiSATextBox cRekTransaksi 
      Height          =   330
      Left            =   2115
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1215
      Width           =   1905
      _ExtentX        =   3360
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
      Enabled         =   0   'False
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
   Begin BiSATextBoxProject.BiSATextBox cDK 
      Height          =   330
      Left            =   2115
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1575
      Width           =   945
      _ExtentX        =   1667
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
      Enabled         =   0   'False
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
   Begin BiSATextBoxProject.BiSATextBox cKas 
      Height          =   330
      Left            =   3090
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1575
      Width           =   930
      _ExtentX        =   1640
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
      Enabled         =   0   'False
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
      Caption         =   "Proses?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   5
      Top             =   2550
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "trAdministrasiSimpanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As New CodeSuiteLibrary.data

Private Sub cGolonganSimpanan_ButtonClick()
Dim db As New ADODB.Recordset
  Set db = obj.Browse(GetDSN, "golongantabungan", "kode,keterangan,rekening")
  If Not db.eof Then
    cGolonganSimpanan.Text = cGolonganSimpanan.Browse(db)
    cGolonganSimpanan.Text = GetNull(db!Kode)
    cRekDebet.Text = GetNull(db!Rekening)
  End If
End Sub

Private Sub cKodeTransaksi_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = obj.Browse(GetDSN, "kodetransaksi")
  If Not db.eof Then
    cKodeTransaksi.Text = cKodeTransaksi.Browse(db)
    cKodeTransaksi.Text = GetNull(db!Kode)
    cRekTransaksi.Text = GetNull(db!Rekening)
    cDK.Text = GetNull(db!DK)
    cKas.Text = GetNull(db!Kas)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim db As New ADODB.Recordset
Dim nSaldoTabungan As Double

  If isValidSimpan Then
    
    Set db = obj.Browse(GetDSN, "tabungan", , "golongantabungan", sisAssign, cGolonganSimpanan.Text)
    If Not db.eof Then
      ProgressBar.Visible = True
      ProgressBar.Max = db.RecordCount
      ProgressBar.Value = 0
      obj.Delete GetDSN, "administrasisimpanan", "kode", sisAssign, cKode.Text
      obj.Delete GetDSN, "mutasitabungan", "faktur", sisAssign, cKode.Text
      obj.Delete GetDSN, "bukubesar", "faktur", sisAssign, cKode.Text
      
      Do While Not db.eof
        ProgressBar.Value = ProgressBar.Value + 1
        
        nSaldoTabungan = GetSaldoTab(obj, GetNull(db!Rekening), dTgl.Value)
        If nSaldoTabungan > 0 Then
          If nSaldoTabungan >= nJumlah.Value Then
            SaveAdministrasi obj, cKode.Text, GetNull(db!Rekening), cRekDebet.Text, cRekTransaksi.Text, nJumlah.Value
          Else
            SaveAdministrasi obj, cKode.Text, GetNull(db!Rekening), cRekDebet.Text, cRekTransaksi.Text, nSaldoTabungan
          End If
'        If -nJumlah.Value < 0 Then
'          If optProses(0).Value = True Then
'            'proses
'            SaveAdministrasi obj, cKode.Text, GetNull(db!Rekening), cRekDebet.Text, cRekTransaksi.Text, nJumlah.Value
'          End If
'        Else
'          'proses
'          SaveAdministrasi obj, cKode.Text, GetNull(db!Rekening), cRekDebet.Text, cRekTransaksi.Text, nJumlah.Value
        End If
        
        db.MoveNext
      Loop
      ProgressBar.Visible = False
    End If
    
  End If
End Sub

Private Sub SaveAdministrasi(ByVal obj As CodeSuiteLibrary.data, ByVal Kode As String, ByVal Rek As String, ByVal RekDebet As String, ByVal RekKredit As String, ByVal Jumlah As Double)
Dim cKodeCabang As String

  cKodeCabang = aCfg(msKodeCabang)
  'simpan di table administrasi
  obj.Add GetDSN, "administrasisimpanan", Array("kode", "rekening", "jumlah", "rekeningsimpanan", "rekeningpendapatan", "username", "datetime", "tgl"), Array(Kode, Rek, Jumlah, RekDebet, RekKredit, GetRegistry(reg_UserName), SNow, Format(dTgl.Value, "yyyy-MM-dd"))
  UpdMutasiTabungan obj, cKodeTransaksi.Text, Kode, Format(dTgl.Value, "yyyy-MM-dd"), Rek, Jumlah, False, "Administrasi Bulanan " & Kode, False, cDK.Text, cRekTransaksi.Text
  
  'Update bukubesar
  UpdKodeTr obj, msTabungan, cKodeCabang, Kode, Format(dTgl.Value, "yyyy-MM-dd"), RekDebet, "Administrasi Bulanan " & Kode, Jumlah, , "N"
    UpdKodeTr obj, msTabungan, cKodeCabang, Kode, Format(dTgl.Value, "yyyy-MM-dd"), RekKredit, "Administrasi Bulanan " & Kode, , Jumlah, "N"
  
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  optProses(0).Value = True
  dTgl.Value = Date
  cGolonganSimpanan.Default
  cRekDebet.Default
  cRekTransaksi.Default
  cDK.Default
  cKas.Default
  nJumlah.Default
  
  TabIndex cKode, n
  TabIndex dTgl, n
  TabIndex cGolonganSimpanan, n
  TabIndex cKodeTransaksi, n
  TabIndex nJumlah, n
  TabIndex optProses(0), n
  TabIndex optProses(1), n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Function isValidSimpan() As Boolean
Dim db As New ADODB.Recordset
isValidSimpan = True
  
  If Trim(cKode.Text) = "" Then
    MsgBox "Form Kode tidak boleh kosong", , "error"
    isValidSimpan = False
    Exit Function
  End If
    
  Set db = obj.Browse(GetDSN, "administrasisimpanan", , "kode", sisAssign, cKode.Text)
  If Not db.eof Then
    MsgBox "Kode " & cKode.Text & " sudah pernah dipakai" & vbCrLf & "Tolong gunakan yang lain", , "error"
    db.Close
    Set db = Nothing
    isValidSimpan = False
    Exit Function
  End If
  
db.Close
Set db = Nothing
End Function

Private Sub optProses_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub
