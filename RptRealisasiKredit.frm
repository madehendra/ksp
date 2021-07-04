VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptRealisasiKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN REALISASI"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7710
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2250
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   3969
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
         Height          =   270
         Index           =   2
         Left            =   4860
         TabIndex        =   8
         Top             =   960
         Width           =   1245
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Belum"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   3405
         TabIndex        =   7
         Top             =   960
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sudah"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   2295
         TabIndex        =   6
         Top             =   960
         Value           =   -1  'True
         Width           =   1080
      End
      Begin BiSATextBoxProject.BiSATextBox cNamaGolongan 
         Height          =   330
         Left            =   3060
         TabIndex        =   0
         Top             =   585
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   315
         TabIndex        =   1
         Top             =   585
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
         Appearance      =   0
         Button          =   -1  'True
         Caption         =   "GOL KREDIT"
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   315
         TabIndex        =   2
         Top             =   210
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   582
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Caption         =   "ANTARA TANGGAL"
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   3495
         TabIndex        =   3
         Top             =   210
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   582
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   465
         Left            =   2070
         Top             =   1665
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   820
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
         Begin VB.OptionButton optJenisKelamin 
            Caption         =   "Laki"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   12
            Top             =   120
            Width           =   720
         End
         Begin VB.OptionButton optJenisKelamin 
            Caption         =   "Perempuan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1350
            TabIndex        =   11
            Top             =   120
            Width           =   1275
         End
         Begin VB.OptionButton optJenisKelamin 
            Caption         =   "Semuanya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2805
            TabIndex        =   10
            Top             =   120
            Width           =   1140
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   435
         Left            =   2070
         Top             =   1245
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   767
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
         Begin VB.OptionButton optAnggota 
            Caption         =   "Anggota"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   15
            Top             =   105
            Width           =   975
         End
         Begin VB.OptionButton optAnggota 
            Caption         =   "Calon Anggota"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1350
            TabIndex        =   14
            Top             =   105
            Width           =   1395
         End
         Begin VB.OptionButton optAnggota 
            Caption         =   "Semuanya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2790
            TabIndex        =   13
            Top             =   105
            Width           =   1065
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Dicairkan?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   375
         TabIndex        =   9
         Top             =   990
         Width           =   1410
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   2235
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
         Left            =   6435
         TabIndex        =   4
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
         Picture         =   "RptRealisasiKredit.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5265
         TabIndex        =   5
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
         Picture         =   "RptRealisasiKredit.frx":00A6
      End
   End
End
Attribute VB_Name = "RptRealisasiKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim CekAktif As Byte
Dim n As Single, M As Single

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganKredit", "Kode", cGolongan, "Kode,Keterangan")
  If Not dbData.eof Then
    cNamaGolongan.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  dDate(0).Value = BOM(Date)
  dDate(1).Value = EOM(Date)
  
  optAnggota(0).Value = True
  optJenisKelamin(0).Value = True
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cGolongan, n
  TabIndex Option1(0), n
  TabIndex Option1(1), n
  TabIndex Option1(2), n
  TabIndex optAnggota(0), n
  TabIndex optAnggota(1), n
  TabIndex optAnggota(2), n
  TabIndex optJenisKelamin(0), n
  TabIndex optJenisKelamin(1), n
  TabIndex optJenisKelamin(2), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader "LAPORAN REALISASI PINJAMAN", tdbHalignCenter, , , , , 12, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy") & " s/d " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "Tanggal", , , , 8
    .AddTableHeader "No. Rekening", , , , 12
    .AddTableHeader "Nama"
    .AddTableHeader "Alamat"
    .AddTableHeader "Plafond", , , , 12
    .AddTableHeader "Lama", , , , 5
    .AddTableHeader "Suku Bunga", , , , 7
    .AddTableHeader "Jth Tmp", , , , 8
    
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody
    
    .AddTableFooter "Total Realisasi", , tdbHalignCenter, , , , , , , , , , , , 4
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    
    .Preview vaArray, True
  End With
End Sub

Private Sub GetData()
Dim nRow As Double
Dim cField As String
Dim cWhere As String

  vaArray.ReDim 0, -1, 0, 7
  cField = "d.Tgl,d.rekening,d.Lama,d.plafond,d.SukuBunga,d.jatuhtempo,r.Nama,r.Alamat,r.jenisanggota,r.kelamin"
  cWhere = ""
  If Option1(0).Value = True Then
    cWhere = cWhere & " and d.statuspencairan = '1'"
  End If
  If Option1(1).Value = True Then
    cWhere = cWhere & " and d.statuspencairan <> '1'"
  End If
  If optAnggota(0).Value = True Then
    cWhere = cWhere & " and r.jenisanggota = '1'"
  End If
  If optAnggota(1).Value = True Then
    cWhere = cWhere & " and r.jenisanggota = '2'"
  End If
  If optJenisKelamin(0).Value = True Then
    cWhere = cWhere & " and r.kelamin = 'L'"
  End If
  If optJenisKelamin(1).Value = True Then
    cWhere = " and r.kelamin = 'P'"
  End If
  
  Set dbData = objData.Browse(GetDSN, "debitur d", cField, "Tgl", sisGTEqual, Format(dDate(0).Value, "yyyy-mm-dd"), " and Tgl <= '" & Format(dDate(1).Value, "yyyy-mm-dd") & "' And d.golonganKredit ='" & cGolongan.Text & "'" & cWhere, "d.tgl,d.Rekening", _
               Array("Left Join registernasabah r on d.Kode = r.kode"))
  
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount + 1
    dbData.MoveFirst
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      nRow = vaArray.UpperBound(1)
      
      vaArray(nRow, 0) = (dbData!Tgl)
      vaArray(nRow, 1) = (dbData!Rekening)
      vaArray(nRow, 2) = (dbData!nama)
      vaArray(nRow, 3) = (dbData!alamat)
      vaArray(nRow, 4) = (dbData!plafond)
      vaArray(nRow, 5) = (dbData!Lama)
      vaArray(nRow, 6) = (dbData!SukuBunga)
      vaArray(nRow, 7) = (dbData!JatuhTempo)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    rpt
  End If
End Sub

Private Sub optAnggota_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub


Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optJenisKelamin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub


