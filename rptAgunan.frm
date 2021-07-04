VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form rptAgunan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agunan"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4170
   Begin BiSATextBoxProject.BiSABrowse cJaminan 
      Height          =   330
      Left            =   345
      TabIndex        =   2
      Top             =   585
      Width           =   3195
      _ExtentX        =   5636
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
      Caption         =   "Jaminan"
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
   Begin VB.CheckBox Check1 
      Caption         =   "Pilih Jaminan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   675
      TabIndex        =   1
      Top             =   225
      Width           =   1335
   End
   Begin BiSAButtonProject.BiSAButton cmdPreview 
      Height          =   555
      Left            =   2310
      TabIndex        =   0
      Top             =   1170
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   979
      Caption         =   "Preview"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
   End
End
Attribute VB_Name = "rptAgunan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub GetSQL()
Dim cSQL As String
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 5
  cSQL = "select a.kode,g.keterangan,a.rekening,r.nama,d.plafond,a.nilaijaminan from agunan a"
  cSQL = cSQL & " left join debitur d on d.rekening = a.rekening"
  cSQL = cSQL & " left join registernasabah r on r.kode = d.kode"
  cSQL = cSQL & " left join gagunan g on g.kode = a.kode"
  cSQL = cSQL & " Where d.status <> 1"
  If Check1.Value = 1 Then
    cSQL = cSQL & " and a.kode = '" & cJaminan.Text & "'"
  End If
  cSQL = cSQL & " order by a.kode,a.rekening;"

  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Kode, "")
      vaArray(n, 1) = GetNull(dbData!Keterangan, "")
      vaArray(n, 2) = GetNull(dbData!Rekening, "")
      vaArray(n, 3) = GetNull(dbData!nama, "")
      vaArray(n, 4) = GetNull(dbData!plafond, "")
      vaArray(n, 5) = GetNull(dbData!nilaijaminan, "")
      dbData.MoveNext
    Loop
    
    With FrmRPT
      .AddPageHeader "Nilai Agunan", tdbHalignCenter, , , , , 10, True
      .AddPageHeader aCfg(msNama), tdbHalignCenter, , , True, , 14, True
      .AddPageHeader " ", , , , True
      .AddPageHeader " ", , , , True
      
      .AddTableGroupHeader True, "[]", , , , 12
      .AddTableGroupHeader
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Rekening", , , , 18, , , , , , True, tdbTableHeaderSect
      .AddTableHeader "Nama"
      .AddTableHeader "Plafond", , , , 15
      .AddTableHeader "Nilai Jaminan", , , , 15
      
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter "Total", , tdbHalignRight, , , , , , , , , , , , 2
      .AddTableGroupFooter
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2

      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 2
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      .Preview vaArray, True
      End With
    
  End If
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub cJaminan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gagunan", "kode,keterangan")
  If Not dbData.eof Then
    cJaminan.Text = cJaminan.Browse(dbData)
    cJaminan.Text = GetNull(dbData!Kode, "")
  End If
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  TabIndex Check1, n
  TabIndex cJaminan, n
  TabIndex cmdPreview, n
End Sub
