VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptJatuhTempo1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN DEPOSITO JATUH TEMPO"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7725
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1590
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
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
      Begin VB.CheckBox chkARO 
         Caption         =   "Non ARO"
         Height          =   300
         Index           =   1
         Left            =   2145
         TabIndex        =   7
         Top             =   1200
         Width           =   1320
      End
      Begin VB.CheckBox chkARO 
         Caption         =   "ARO"
         Height          =   300
         Index           =   0
         Left            =   2145
         TabIndex        =   6
         Top             =   930
         Width           =   840
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
         Index           =   0
         Left            =   315
         TabIndex        =   2
         Top             =   210
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
         Caption         =   "TGL JATUH TEMPO"
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
      Top             =   1590
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
         Picture         =   "RptJatuhTempo.frx":0000
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
         Picture         =   "RptJatuhTempo.frx":00A6
      End
   End
End
Attribute VB_Name = "RptJatuhTempo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim cSQL As String

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganDeposito", "Kode", cGolongan, "Kode,Keterangan")
  If Not dbData.eof Then
    cNamaGolongan.Text = dbData!Keterangan
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Or Trim(cGolongan.Text) <> "" Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub chkARO_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
    Exit Sub
  End If
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cGolongan, n
  TabIndex chkARO(0), n
  TabIndex chkARO(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim n As Integer
Dim dTanggalCair As Date
Dim cField As String
Dim vaJoin
Dim cWhere As String

  vaArray.ReDim 0, -1, 0, 6
  cField = "d.lama as LamaDeposito,d.Rekening, d.Nominaldeposito, d.Tgl, d.SukuBunga, d.JthTmp,d.Status,r.Nama, g.Lama"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode=d.Kode", _
                 "Left Join GolonganDeposito g on g.Kode = d.golonganDeposito")
  cWhere = "And d.JthTmp >='" & Format(dDate(0).Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & "And d.JthTmp <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " And d.Status <> '1'"
  If chkARO(0).Value = 1 And chkARO(1).Value = 0 Then
    cWhere = cWhere & " and d.SistemARO = 'Y'"
  ElseIf chkARO(0).Value = 0 And chkARO(1).Value = 1 Then
    cWhere = cWhere & " and d.SistemARO ='T'"
  End If
  Set dbData = objData.Browse(GetDSN, "Deposito d", cField, "d.GolonganDeposito", sisAssign, cGolongan.Text, cWhere, "d.Golongandeposito,d.rekening", vaJoin)
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = (dbData!Rekening)
      vaArray(n, 1) = (dbData!nama)
      vaArray(n, 2) = (dbData!LamaDeposito)
      vaArray(n, 3) = (dbData!SukuBunga)
      vaArray(n, 4) = (dbData!Tgl)
      vaArray(n, 5) = (dbData!nominaldeposito)
      vaArray(n, 6) = (dbData!jthtmp)
      dbData.MoveNext
    Loop
    rpt
  End If
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader UCase("Laporan Deposito Jatuh Tempo"), tdbHalignCenter, , , , , 12, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 10
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy") & " s.d " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 10
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "No. Rekening", , , , 13
    .AddTableHeader "Nama"
    .AddTableHeader "Lama", , , , 7
    .AddTableHeader "Suku Bunga", , , , 7
    .AddTableHeader "Tgl. Valuta", , , , 10
    .AddTableHeader "Nominal", , , , 13
    .AddTableHeader "Jatuh Tempo", , , , 10
    
    .AddTableBody
    .AddTableBody
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignCenter
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    
    .AddTableFooter "TOTAL", , tdbHalignRight, , , , , , , , , , , , 5
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter ""
    
    .Preview vaArray, True
  End With
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub


