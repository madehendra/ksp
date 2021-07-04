VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form TrBatalBilyet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMBATALAN NOMOR BILYET DEPOSITO"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   7920
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2235
      Left            =   0
      Top             =   0
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   3942
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
      Begin BiSADateProject.BiSADate dTanggal 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   1470
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   556
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
         Caption         =   "Tgl cetak"
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
         TabIndex        =   1
         Top             =   435
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
         Enabled         =   0   'False
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
         TabIndex        =   2
         Top             =   780
         Width           =   6240
         _ExtentX        =   11007
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
      Begin BiSATextBoxProject.BiSATextBox cRekening 
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   1125
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   582
         Text            =   "12345678901234567890"
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
         MaxLength       =   20
         Appearance      =   0
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
      Begin BiSATextBoxProject.BiSATextBox cNomorBilyet 
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   90
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   582
         Text            =   "1234567890"
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
         MaxLength       =   10
         Caption         =   "No Bilyet"
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   1800
         Width           =   7680
         _ExtentX        =   13547
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   2220
      Width           =   7920
      _ExtentX        =   13970
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
         Left            =   5655
         TabIndex        =   6
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
         Picture         =   "TrBatalBilyet.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6735
         TabIndex        =   7
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
         Picture         =   "TrBatalBilyet.frx":0416
      End
   End
End
Attribute VB_Name = "TrBatalBilyet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub cmdSimpan_Click()
  If ValidSaving() Then
    If MsgBox("Apakah Data Benar-benar sudah Valid ?", vbYesNo + vbInformation) = vbYes Then
    
       objData.Edit GetDSN, "DEPOSITO", "Rekening = '" & cRekening.Text & "'", Array("NoBilyet", "NoSeri"), Array("", "")
       objData.Edit GetDSN, "NomorBilyet", "NomorBilyet='" & cNomorBilyet.Text & "'", Array("Status"), Array("1")
    End If
    
  End If
  initvalue
  cNomorBilyet.SetFocus
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cNomorBilyet.Text, "Nomor Bilyet Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cNomorBilyet.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKeterangan.Text, "Keterangan Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
End Function

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.Nama,r.ALamat,d.Rekening,n.Tanggal,d.noBilyet", "r.Nama", sisContent, cNama.Text, "And d.Status <> '1'", "r.Nama", _
                              Array("left Join RegisterNasabah r on r.Kode=d.Kode", _
                                    "Left Join NomorBilyet n on NomorBilyet=d.NoBilyet"))
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.eof Then
    GetRegister
  End If
End Sub

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.ALamat,r.Nama,d.Rekening,n.Tanggal,n.NomorBilyet", "r.Alamat", sisContent, cAlamat.Text, "And d.Status <> '1'", "r.Nama", _
                              Array("left Join RegisterNasabah r on r.Kode=d.Kode", _
                                    "Left Join NomorBilyet n on NomorBilyet=d.NoBilyet"))
  cAlamat.Text = cAlamat.Browse(dbData)
  If Not dbData.eof Then
    GetRegister
  End If
End Sub

Private Sub cNomorBilyet_Validate(Cancel As Boolean)
Dim vaJoin
  
  If cNomorBilyet.LastKey = 13 Then
     cNomorBilyet.Text = GetBilyetDeposito("01", cNomorBilyet.Text)
     If cNomorBilyet.Text <> "" Then
        vaJoin = Array("Left Join RegisterNasabah r on r.Kode=d.Kode", _
                       "Left Join NomorBilyet n on NomorBilyet=d.NoBilyet")
        Set dbData = objData.Browse(GetDSN, "Deposito d", "d.Rekening,d.NoBilyet,n.Tanggal,r.Nama,r.Alamat", "NoBilyet", sisAssign, cNomorBilyet.Text, , , vaJoin)
        If Not dbData.eof Then
          GetRegister
        Else
           MsgBox "Data Tidak Ada.", vbOKOnly + vbInformation, "Pembatalan Nomor Bilyet Deposito"
           Cancel = True
           initvalue
           cNomorBilyet.SetFocus
           Exit Sub
        End If
      Else
        MsgBox "Inputan tidak boleh kosong", vbInformation
        Cancel = True
        cNomorBilyet.SetFocus
        Exit Sub
      End If
  End If
End Sub

Function GetBilyetDeposito(ByVal Cabang As String, ByVal cNomorFakturDeposito As String) As String
Dim cNoBilyet As String

  cNoBilyet = Cabang & Padl(Trim(cNomorFakturDeposito), 8, "0")
  GetBilyetDeposito = cNoBilyet
End Function

  
Private Sub GetRegister()
  cNomorBilyet.Text = GetNull(dbData!Nobilyet, "")
  cRekening.Text = GetNull(dbData!Rekening, "")
  dTanggal.Value = GetNull(dbData!Tanggal, "")
  cNama.Text = GetNull(dbData!nama, "")
  cAlamat.Text = GetNull(dbData!alamat, "")
End Sub

Private Sub initvalue()
  cNomorBilyet.Default
  cRekening.Default
  cNama.Default
  cAlamat.Default
  cKeterangan.Default
  dTanggal.Value = Date
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  initvalue
  
  TabIndex cNomorBilyet, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cKeterangan, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub


