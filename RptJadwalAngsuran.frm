VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptJadwalAngsuran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN JADWAL ANGSURAN"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7230
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   1590
      Left            =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2805
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   300
         Left            =   135
         TabIndex        =   0
         Top             =   1080
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   529
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
         Caption         =   "TGL REALISASI"
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   300
         Left            =   3735
         TabIndex        =   1
         Top             =   90
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   529
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
         Height          =   300
         Left            =   2175
         TabIndex        =   2
         Top             =   90
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   529
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
         Height          =   300
         Left            =   135
         TabIndex        =   3
         Top             =   90
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
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
         Height          =   300
         Left            =   2925
         TabIndex        =   4
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   300
         Left            =   135
         TabIndex        =   5
         Top             =   420
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   529
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
         Caption         =   "NAMA"
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
         Height          =   300
         Left            =   135
         TabIndex        =   6
         Top             =   750
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   529
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
         Caption         =   "ALAMAT"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1575
      Width           =   7215
      _ExtentX        =   12726
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Height          =   435
         Left            =   5970
         TabIndex        =   7
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "RptJadwalAngsuran.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4800
         TabIndex        =   8
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Preview"
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
         Picture         =   "RptJadwalAngsuran.frx":00A6
      End
   End
End
Attribute VB_Name = "RptJadwalAngsuran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim dbData1 As New ADODB.Recordset
Dim xArray As New XArrayDB
Dim nLama As Integer
Dim cNoSPK As String
Dim dJatuhTempo As Date
Dim nPlafond As Double
Dim nPersBunga As Double
Dim cStatus As String
Dim nTotalBunga As Double

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cRekening As String

  If cFrekuensi.LastKey = 13 Then
    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    Set dbData = objData.Browse(GetDSN, "Debitur", "Rekening,StatusPencairan", "Rekening", sisAssign, cRekening)
    If Not dbData.eof Then
      GetMemory
    Else
      MsgBox "No. Rekening Tidak Ditemukan, Ulangi Pengisian", vbExclamation + vbOKOnly
      Cancel = True
      cFrekuensi.SetFocus
    End If
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganKredit", "Kode", cGolongan, "Kode,Keterangan")
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cmdKeluar_Click()
   Unload Me
End Sub

Private Sub GetMemory()
Dim n As Integer
Dim vaJoin
Dim cField As String
Dim cRekening As String

  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cField = "d.*,r.Nama, r.Alamat"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode = d.Kode")
  Set dbData = objData.Browse(GetDSN, "Debitur d", cField, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    dTgl.Value = GetNull(dbData!Tgl)
    cNama.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cNoSPK = GetNull(dbData!NoSPK, "")
    nPersBunga = GetNull(dbData!SukuBunga)
    nPlafond = GetNull(dbData!plafond)
    nLama = GetNull(dbData!Lama)
    dJatuhTempo = GetNull(dbData!JatuhTempo)
    cStatus = GetNull(dbData!caraperhitungan)
    nTotalBunga = GetNull(dbData!totalBunga)
  End If
End Sub

Private Sub cmdPreview_Click()
  If cStatus = "1" Then
    GetJadwalMenurun
  Else
    GetJadwalFlat
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub Form_Load()
Dim n As Single

  cStatus = ""
  
  CenterForm Me
  initvalue
  
  cCabang.Text = aCfg(msKodeCabang)
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub initvalue()
  dTgl.Value = Date
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
  nTotalBunga = 0
End Sub

Private Sub GetJadwalMenurun()
Dim n As Single
Dim AngsPokok As Double
Dim dTanggal As Date
Dim nSukuBungaPerBulan As Double
Dim nKe As Integer
  
  xArray.ReDim 0, nLama, 0, 6
  dTanggal = (DateAdd("m", 1, dTgl.Value))
  nSukuBungaPerBulan = Round(nPersBunga / 12, 2)
  xArray(0, 5) = Round(nPlafond * nPersBunga / 100 / 12 * nLama, 0)
  xArray(0, 6) = nPlafond
  nKe = 1
  For n = 1 To nLama
    xArray(n, 0) = n
    xArray(n, 1) = DateAdd("m", n, dTgl.Value)
    xArray(n, 2) = GetBungaReguler(xArray(n - 1, 6), nSukuBungaPerBulan)
    xArray(n, 3) = nPlafond / (nLama)
    xArray(n, 4) = xArray(n, 2) + xArray(n, 3)
    xArray(n, 5) = xArray(n - 1, 5) - xArray(n, 2)
    xArray(n, 6) = xArray(n - 1, 6) - xArray(n, 3)
    dTanggal = (DateAdd("m", 1, xArray(n, 1)))
  Next
  rpt
End Sub

Private Sub GetJadwalFlat()
Dim n As Single
Dim dTanggal As Date
Dim nSukuBungaPerBulan As Double
Dim nKe As Integer

  xArray.ReDim 0, nLama, 0, 6
  dTanggal = (DateAdd("m", 1, dTgl.Value))
  nSukuBungaPerBulan = Round(nPersBunga / 12, 2)
  xArray(0, 5) = nTotalBunga 'nBunga.Value
  xArray(0, 6) = nPlafond 'nPlafond.Value
  nKe = 1
  For n = 1 To nLama
    xArray(n, 0) = n
    xArray(n, 1) = dTanggal
    xArray(n, 2) = Devide(nTotalBunga, nLama)   'GetBungaReguler(xArray(n - 1, 6), nSukuBungaPerBulan)
    xArray(n, 3) = nPlafond / (nLama)
    xArray(n, 4) = xArray(n, 2) + xArray(n, 3)
    xArray(n, 5) = xArray(n - 1, 5) - xArray(n, 2)
    xArray(n, 6) = xArray(n - 1, 6) - xArray(n, 3)
    dTanggal = (DateAdd("m", 1, xArray(n, 1)))
  Next
  rpt
End Sub

Private Function GetBungaReguler(ByVal nSisaPokok As Double, ByVal nBunga As Double) As Double
  GetBungaReguler = nSisaPokok * (nBunga / 100)
  GetBungaReguler = Mod50(GetBungaReguler)
End Function

Private Sub rpt()
  With FrmRPT
    .AddPageHeader "JADWAL ANGSURAN", tdbHalignCenter, , , , dbArial, 12, True, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddPageHeader "NO. REKENING", , , 15, True, , , True
    .AddPageHeader ": " & SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text), , , 50
    
    .AddPageHeader "PLAFOND", , , 15, , , , True
    .AddPageHeader ": ", , , 2, , , , True
    .AddPageHeader Format(nPlafond, "###,###,###,###,###,###"), tdbHalignRight, , 12
    
    .AddPageHeader "NAMA DEBITUR", , , 15, True, , , True
    .AddPageHeader ": " & cNama.Text, , , 50
    
    .AddPageHeader "LAMA ANGS.", , , 15, , , , True
    .AddPageHeader ": " & nLama & " Bulan"
    
    .AddPageHeader "ALAMAT", , , 15, True, , , True
    .AddPageHeader ": " & cAlamat.Text, , , 50
    
    .AddPageHeader "JATUH TEMPO", , , 15, , , , True
    .AddPageHeader ": " & Format(dJatuhTempo, "dd-MM-yyyy")
        
    .AddPageHeader "No. SPK", , , 15, True, , , True
    .AddPageHeader ": " & cNoSPK, , , 50
    
    .AddPageHeader "TGL REALISASI", , , 15, True, , , True
    .AddPageHeader ": " & Format(dTgl.Value, "dd-MM-yyyy"), , , 50
        
    .AddTableHeader "KE", , , , 8, , , , , , True, , , tdbMergeOnText
    .AddTableHeader "JTHTMP", , , , 12, , , , , , , , , tdbMergeOnText
    .AddTableHeader "BUNGA"
    .AddTableHeader "POKOK"
    .AddTableHeader "ANGSURAN"
    .AddTableHeader "SISA BUNGA"
    .AddTableHeader "SISA POKOK"
    
    .AddTableBody
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableFooter "TOTAL", , tdbHalignCenter, , , , , , , , , , , , 2
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    .AddTableFooter
    
    .Preview xArray
  End With
End Sub

