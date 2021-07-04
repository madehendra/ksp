VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptMutasiTabunganHarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN MUTASI SIMPANAN HARIAN"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   7725
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1935
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   3413
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
      Begin VB.OptionButton Option1 
         Caption         =   "&Tidak"
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
         Index           =   1
         Left            =   2775
         TabIndex        =   11
         Top             =   1230
         Width           =   840
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Ya"
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
         Index           =   0
         Left            =   2115
         TabIndex        =   10
         Top             =   1230
         Value           =   -1  'True
         Width           =   525
      End
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Left            =   315
         TabIndex        =   0
         Top             =   150
         Width           =   3270
         _ExtentX        =   5768
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
         Caption         =   "TANGGAL"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGolongan 
         Height          =   330
         Left            =   3060
         TabIndex        =   1
         Top             =   510
         Width           =   4095
         _ExtentX        =   7223
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
         BackColor       =   -2147483633
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   315
         TabIndex        =   2
         Top             =   510
         Width           =   2745
         _ExtentX        =   4842
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
         Caption         =   "GOLONGAN"
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
      Begin BiSATextBoxProject.BiSABrowse cPdl 
         Height          =   330
         Index           =   0
         Left            =   315
         TabIndex        =   5
         Top             =   870
         Width           =   3030
         _ExtentX        =   5345
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
         Caption         =   "ANTARA PDL"
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
      Begin BiSATextBoxProject.BiSABrowse cPdl 
         Height          =   330
         Index           =   1
         Left            =   3390
         TabIndex        =   6
         Top             =   855
         Width           =   1800
         _ExtentX        =   3175
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
      Begin BiSATextBoxProject.BiSABrowse cJenisTransaksi 
         Height          =   330
         Left            =   315
         TabIndex        =   7
         Top             =   1485
         Width           =   2730
         _ExtentX        =   4815
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
         Caption         =   "JENIS TRANSAKSI"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaJenisTransaksi 
         Height          =   330
         Left            =   3045
         TabIndex        =   8
         Top             =   1485
         Width           =   4095
         _ExtentX        =   7223
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
         BackColor       =   -2147483633
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
      Begin BiSADateProject.BiSADate dDate2 
         Height          =   330
         Left            =   3675
         TabIndex        =   12
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
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
         Caption         =   "sd"
         CaptionWidth    =   0
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
         Caption         =   "FILTER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   405
         TabIndex        =   9
         Top             =   1245
         Width           =   1095
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1920
      Width           =   7710
      _ExtentX        =   13600
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
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6420
         TabIndex        =   3
         Top             =   105
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
         Picture         =   "RptMutasiTabunganHarian.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5250
         TabIndex        =   4
         Top             =   105
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
         Picture         =   "RptMutasiTabunganHarian.frx":00A6
      End
   End
End
Attribute VB_Name = "RptMutasiTabunganHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganTabungan", "Kode", cGolongan, "Kode,Keterangan")
  If Not dbData.eof Then
    cNamaGolongan.Text = dbData!Keterangan
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Or Trim(cGolongan.Text) <> "" Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cJenisTransaksi_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "KodeTransaksi", "Kode", cJenisTransaksi, "Kode,Keterangan")
  If Not dbData.eof Then
    cNamaJenisTransaksi.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cJenisTransaksi_Validate(Cancel As Boolean)
  cJenisTransaksi_ButtonClick
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub GetRpt()
    With FrmRPT
    .AddPageHeader UCase("Daftar Mutasi Simpanan Harian"), tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "TANGGAL  : " & Format(dDate.Value, "dd MMMM yyyy"), tdbHalignCenter, , , True, , 12, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableGroupHeader True, "[]", , , , 10
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "NO. TRANSAKSI", , , , 14
    .AddTableHeader "NO. REKENING", , , , 10
    .AddTableHeader "NO. REG PDL", , , , 8
    .AddTableHeader "NAMA NASABAH", , , , 20
    .AddTableHeader "SD", , , , 3
    .AddTableHeader "KETERANGAN"
    .AddTableHeader "DEBET", , , , 10
    .AddTableHeader "KREDIT", , , , 10
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody , tdbHalignCenter
    .AddTableBody
    .AddTableBody , tdbHalignCenter
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "SUB TOTAL", , tdbHalignRight, , , , , , , , , , , , 6
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&sum", Sis_Rpt_Number2
    .AddTableGroupFooter "&sum", Sis_Rpt_Number2
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "GRAND TOTAL", , tdbHalignRight, , , , , , , , , , , , 6
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True, , True
  End With
End Sub

Private Sub GetSQL()
Dim vaJoin
Dim cWhere As String
Dim cField As String

  cField = "t.PDL,a.Nama as NamaPDL, m.Faktur,m.Rekening,t.NoPDL,r.Nama,"
  cField = cField & "m.KodeTransaksi,m.Keterangan,if(m.DK='D',m.Jumlah,0),if(m.DK='K',m.Jumlah,0)"
'  cWhere = " and m.Tgl >= '" & Format(dDate.Value, "yyyy-mm-dd") & "'"
'  cWhere = " and m.Tgl <= '" & Format(dDate2.Value, "yyyy-mm-dd") & "'"
  cWhere = " and m.Tgl = '" & Format(dDate.Value, "yyyy-mm-dd") & "'"

'  cWhere = " and m.Tgl >= '" & Format(dDate.Value, "yyyy-mm-dd") & "'"
'  cWhere = cWhere & " and m.Tgl <= '" & Format(dDate2.Value, "yyyy-mm-dd") & "'"

  cWhere = cWhere & " and t.GolonganTabungan='" & cGolongan.Text & "'"
'  cWhere = cWhere & " and t.pdl >='" & cPDL(0).Text & "'"
'  cWhere = cWhere & " and t.pdl <='" & cPDL(1).Text & "'"
  If Trim(cJenisTransaksi.Text) <> "" Then
    cWhere = cWhere & " and m.KodeTransaksi = '" & cJenisTransaksi.Text & "'"
  End If
  vaJoin = Array("Left Join Tabungan t On m.Rekening = t.Rekening", _
                 "Left Join RegisterNasabah r On t.Kode = r.Kode", _
                 "Left Join AO a on a.Kode = t.PDL")
  Set dbData = objData.Browse(GetDSN, "MutasiTabungan m", cField, _
                              "m.Posting", sisAssign, "0", cWhere, "m.Rekening,m.Tgl,m.ID", vaJoin)
  If Not dbData.eof Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 0, XORDER_ASCEND, XTYPE_DEFAULT, 2, XORDER_ASCEND, XTYPE_DEFAULT
    GetRpt
  Else
    MsgBox "Data Tidak Ada,..", vbInformation, Me.Caption
  End If
End Sub

Private Sub cPdl_ButtonClick(Index As Integer)
  Set dbData = objData.Pick(GetDSN, "PDL", "Kode", cPDL(Index), "Kode,Keterangan")
End Sub

Private Sub cPdl_Validate(Index As Integer, Cancel As Boolean)
  If cPDL(Index).LastKey = 13 Then
    cPdl_ButtonClick (Index)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  dDate.Value = Date
  dDate2.Value = Date
  GetMinMax "PDL", cPDL, "Kode"
      
  TabIndex dDate, n
  TabIndex dDate2, n
  TabIndex cGolongan, n
  TabIndex cPDL(0), n
  TabIndex cPDL(1), n
  TabIndex Option1(0), n
  TabIndex Option1(1), n
  TabIndex cJenisTransaksi, n
  TabIndex cNamaJenisTransaksi, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
  cJenisTransaksi.Enabled = True
End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub Option1_LostFocus(Index As Integer)
If Option1(0).Value = True Then
    cJenisTransaksi.Enabled = True
  Else
    cJenisTransaksi.Enabled = False
    cJenisTransaksi.Default
    cNamaJenisTransaksi.Default
  End If
End Sub
