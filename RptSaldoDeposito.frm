VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptSaldoDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN SALDO DEPOSITO"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7725
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2325
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   4101
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
         Button          =   -1  'True
         Caption         =   "GOL DEPOSITO"
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
         Left            =   315
         TabIndex        =   2
         Top             =   210
         Width           =   3300
         _ExtentX        =   5821
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
         ForeColor       =   -2147483640
         Caption         =   "SAMPAI TANGGAL"
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   435
         Left            =   2100
         Top             =   945
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   105
            Width           =   1065
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   465
         Left            =   2100
         Top             =   1365
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
            Left            =   285
            TabIndex        =   10
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
            TabIndex        =   9
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
            Left            =   2835
            TabIndex        =   8
            Top             =   120
            Width           =   1140
         End
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   2280
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6435
         TabIndex        =   3
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
         Picture         =   "RptSaldoDeposito.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5265
         TabIndex        =   4
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
         Picture         =   "RptSaldoDeposito.frx":00A6
      End
   End
End
Attribute VB_Name = "RptSaldoDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganDeposito", "Kode", cGolongan, "Kode,Keterangan")
  If dbData.RecordCount > 0 Then
    cNamaGolongan.Text = dbData!Keterangan
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Or Trim(cGolongan.Text) <> "" Then
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
  optAnggota(2).Value = True
  optJenisKelamin(2).Value = True
  TabIndex dDate, n
  TabIndex cGolongan, n
  TabIndex optAnggota(0), n
  TabIndex optAnggota(1), n
  TabIndex optAnggota(2), n
  TabIndex optJenisKelamin(0), n
  TabIndex optJenisKelamin(1), n
  TabIndex optJenisKelamin(2), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim n As Integer
Dim dTanggalCair As Date
Dim cField As String
Dim vaJoin
Dim cWhere As String
Dim nTotal As Double

  vaArray.ReDim 0, -1, 0, 11
 
  cField = "d.lama as lamaDeposito,d.Rekening, d.Nominaldeposito,d.Tgl, d.SukuBunga, d.JthTmp,d.Status,d.TglCair,g.PajakBunga,d.pdl,p.keterangan as namapdl,"
  cField = cField & "r.Nama, g.Lama"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode=d.Kode", _
                 "Left Join GolonganDeposito g on g.Kode = d.GolonganDeposito", _
                 "left join pdl p on p.kode = d.pdl")
  cWhere = "And d.Tgl <='" & Format(dDate.Value, "yyyy-MM-dd") & "'"
  
  'Filter anggota/calon anggota
  If optAnggota(0).Value = True Then
    cWhere = cWhere & " and r.jenisanggota = '1'"
  ElseIf optAnggota(1).Value = True Then
    cWhere = cWhere & " and r.jenisanggota = '2'"
  End If
  
  If optJenisKelamin(0).Value = True Then 'laki
    cWhere = cWhere & " and r.kelamin = 'L'"
  ElseIf optJenisKelamin(1).Value = True Then
    cWhere = cWhere & " and r.kelamin = 'P'"
  End If
  
  Set dbData = objData.Browse(GetDSN, "Deposito d", cField, "d.GolonganDeposito", sisAssign, cGolongan.Text, cWhere, "d.Golongandeposito,d.Rekening", vaJoin)
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      dTanggalCair = IIf(GetNull(dbData!TglCair, 0) = 0, dbData!Tgl, dbData!TglCair)
      If dTanggalCair <= dDate.Value And dbData!status = "1" Then
      Else
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = GetNull(dbData!Rekening, "")
        vaArray(n, 1) = GetNull(dbData!nama, "")
        vaArray(n, 2) = GetNull(dbData!LamaDeposito)
        vaArray(n, 3) = GetNull(dbData!SukuBunga)
        vaArray(n, 4) = GetNull(dbData!Tgl)
        vaArray(n, 5) = GetNull(dbData!jthtmp)
        vaArray(n, 6) = GetNull(dbData!nominaldeposito)
        vaArray(n, 7) = Round(vaArray(n, 10) * Day(EOM(dDate.Value)) * GetNull(dbData!SukuBunga) / 365 / 100, 0)
        If vaArray(n, 6) > 7500000 Then
          vaArray(n, 8) = Round(vaArray(n, 7) * GetNull(dbData!pajakbunga) / 100, 2)
        Else
          vaArray(n, 8) = 0
        End If
        vaArray(n, 9) = vaArray(n, 7) - vaArray(n, 8)
        vaArray(n, 10) = vaArray(n, 6) + vaArray(n, 9)
        vaArray(n, 11) = GetNull(dbData!namapdl)
      End If
      dbData.MoveNext
    Loop
    rpt
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Sub
  End If
End Sub

Private Sub rpt()
    With FrmRPT
      .AddPageHeader UCase("Laporan Saldo Deposito"), tdbHalignCenter, , , , , 12, True
      .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 12
      .AddPageHeader "Sampai Tanggal : " & Format(dDate.Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 9, True
      .AddPageHeader " ", , , , True
      .AddPageHeader " ", , , , True
      
      .AddTableHeader "No. Rekening", , , , 10
      .AddTableHeader "Nama"
      .AddTableHeader "Lama", , , , 4
      .AddTableHeader "Rate", , , , 4
      .AddTableHeader "Tgl. Valuta", , , , 7
      .AddTableHeader "Jatuh Tempo", , , , 7
      .AddTableHeader "Nominal", , , , 8
      .AddTableHeader "Bunga Kotor", , , , 7
      .AddTableHeader "Pajak", , , , 7
      .AddTableHeader "Bunga Bersih", , , , 7
      .AddTableHeader "Total Saldo", , , , 8
      .AddTableHeader "PDL", , , , 9
      
      .AddTableBody
      .AddTableBody
      .AddTableBody , tdbHalignRight
      .AddTableBody , tdbHalignRight
      .AddTableBody Sis_Rpt_dd_MM_yyyy
      .AddTableBody Sis_Rpt_dd_MM_yyyy
      .AddTableBody Sis_Rpt_Number
      .AddTableBody Sis_Rpt_Number
      .AddTableBody Sis_Rpt_Number
      .AddTableBody Sis_Rpt_Number
      .AddTableBody Sis_Rpt_Number
      .AddTableBody
      
      .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 5
      .AddTableFooter ""
      .AddTableFooter ""
      .AddTableFooter ""
      .AddTableFooter ""
      .AddTableFooter ""
      .AddTableFooter "&Sum", Sis_Rpt_Number
      .AddTableFooter "&Sum", Sis_Rpt_Number
      .AddTableFooter "&Sum", Sis_Rpt_Number
      .AddTableFooter "&Sum", Sis_Rpt_Number
      .AddTableFooter "&Sum", Sis_Rpt_Number
      .AddTableFooter ""
      
      .Preview vaArray, True, , True
    End With
End Sub

Private Sub optAnggota_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optJenisKelamin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub
