VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptDebtRealisasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debt Collector - Realisasi"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   7710
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1500
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2646
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
      Begin VB.CheckBox Check1 
         Caption         =   "Pilih Kolektor"
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
         Left            =   3015
         TabIndex        =   5
         Top             =   645
         Width           =   1860
      End
      Begin BiSATextBoxProject.BiSABrowse cKolektor 
         Height          =   330
         Left            =   1875
         TabIndex        =   4
         Top             =   870
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "Kolektor"
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
         Left            =   1200
         TabIndex        =   0
         Top             =   210
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   582
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
         Left            =   4380
         TabIndex        =   1
         Top             =   210
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   582
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1485
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
         TabIndex        =   2
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
         Picture         =   "rptDebtRealisasi.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5265
         TabIndex        =   3
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
         Picture         =   "rptDebtRealisasi.frx":00A6
      End
   End
End
Attribute VB_Name = "rptDebtRealisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub cKolektor_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "ao", "kode,nama", "nama", sisContent, cKolektor.Text)
  If Not dbData.eof Then
    cKolektor.Text = cKolektor.Browse(dbData)
    cKolektor.Text = GetNull(dbData!Kode, "")
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  Check1.Value = 0
  dDate(0).Value = BOM(Date)
  dDate(1).Value = EOM(Date)
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex Check1, n
  TabIndex cKolektor, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSQL()
Dim cSQL As String
Dim n As Integer

  cSQL = " select d.tgl,d.rekening,r.nama,d.ao,a.nama as namaao,d.plafond from debitur d"
  cSQL = cSQL & " left join registernasabah r on r.kode = d.kode"
  cSQL = cSQL & " left join ao a on a.kode = d.ao"
  cSQL = cSQL & " where tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' and tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  If Check1.Value = 1 Then
    cSQL = cSQL & " and d.ao = '" & cKolektor.Text & "'"
  End If
  cSQL = cSQL & " order by d.ao,d.tgl"
  
  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!AO)
      vaArray(n, 1) = GetNull(dbData!namaao)
      vaArray(n, 2) = GetNull(dbData!Tgl)
      vaArray(n, 3) = GetNull(dbData!Rekening)
      vaArray(n, 4) = GetNull(dbData!nama)
      vaArray(n, 5) = GetNull(dbData!plafond)
      dbData.MoveNext
    Loop
    GetRpt
  End If
End Sub

Private Sub GetRpt()
    With FrmRPT
    .AddPageHeader "LAPORAN REALISASI PER KOLEKTOR", tdbHalignCenter, , , , , 10, True
    .AddPageHeader "TANGGAL  : " & Format(dDate(0).Value, "dd MMMM yyyy") & " sd " & Format(dDate(1).Value, "dd MMMM yyyy"), tdbHalignCenter, , , True, , 10, True
    .AddPageHeader aCfg(msNama), tdbHalignCenter, , , True, , 14, True
    
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableGroupHeader True, "[]", , , , 10
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "TGL", , , , 12
    .AddTableHeader "REKENING", , , , 15
    .AddTableHeader "NAMA"
    .AddTableHeader "PLAFOND", , , , 15
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "SUB TOTAL", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&sum", Sis_Rpt_Number2
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "GRAND TOTAL", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub
