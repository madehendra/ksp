VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptDaftarDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN DAFTAR DEPOSAN"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7710
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1560
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2752
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
         Caption         =   "ARO"
         Height          =   300
         Index           =   0
         Left            =   2145
         TabIndex        =   7
         Top             =   915
         Width           =   840
      End
      Begin VB.CheckBox chkARO 
         Caption         =   "Non ARO"
         Height          =   300
         Index           =   1
         Left            =   2145
         TabIndex        =   6
         Top             =   1155
         Width           =   1320
      End
      Begin BiSATextBoxProject.BiSATextBox cNamaGolongan 
         Height          =   330
         Left            =   3060
         TabIndex        =   0
         Top             =   570
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
         Top             =   570
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
         Caption         =   "TGL VALUTA"
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
         TabIndex        =   5
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
      Top             =   1560
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
         Picture         =   "RptDaftarDeposito.frx":0000
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
         Picture         =   "RptDaftarDeposito.frx":00A6
      End
   End
End
Attribute VB_Name = "RptDaftarDeposito"
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

Private Sub chkARO_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
    Exit Sub
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader UCase("Laporan Daftar Deposan"), tdbHalignCenter, , , , , 12, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "Tanggal Valuta : " & dDate(0).Value & " s/d " & dDate(1).Value, tdbHalignCenter, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
        
    .AddTableGroupHeader True, "[]", , , , 8
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
    .AddTableHeader "Tgl Valuta", , , , 7
    .AddTableHeader "No. Rekening", , , , 9
    .AddTableHeader "Nama", , , , 20
    .AddTableHeader "Alamat"
    .AddTableHeader "Suku Bunga", , , , 5
    .AddTableHeader "JthTmp", , , , 7
    .AddTableHeader "No. Bilyet", , , , 10
    .AddTableHeader "Nominal", , , , 13
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody , tdbHalignRight
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "Sub Total", , tdbHalignRight, , , , , , , , , , , , 7
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 7
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True, , True
  End With
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub GetData()
Dim n As Integer
Dim cField As String
Dim vaJoin
Dim cWhere As String
    
  cField = "d.GolonganDeposito,g.Keterangan,d.Tgl,d.Rekening,r.Nama,r.Alamat,d.SukuBunga,d.JthTmp,d.NoBilyet,d.NominalDeposito"
  vaJoin = Array("Left Join RegisterNasabah r On r.Kode = d.Kode", _
                 "Left Join GolonganDeposito g on g.Kode = d.golonganDeposito")
  cWhere = "And d.tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' and d.tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  If chkARO(0).Value = 1 And chkARO(1).Value = 0 Then
    cWhere = cWhere & " and d.SistemARO = 'Y'"
  ElseIf chkARO(0).Value = 0 And chkARO(1).Value = 1 Then
    cWhere = cWhere & " and d.SistemARO ='T'"
  End If
  Set dbData = objData.Browse(GetDSN, "Deposito d", cField, "d.GolonganDeposito", sisAssign, cGolongan.Text, cWhere, "d.GolonganDeposito,d.Tgl", vaJoin)
  If Not dbData.eof Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    rpt
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  dDate(0).Value = BOM(Date)
  dDate(1).Value = Date
      
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cGolongan, n
  TabIndex chkARO(0), n
  TabIndex chkARO(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub


