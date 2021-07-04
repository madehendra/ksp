VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptArusKas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN ARUS KAS"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   9885
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1200
      Left            =   0
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   2117
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekening 
         Height          =   330
         Left            =   4725
         TabIndex        =   0
         Top             =   210
         Width           =   5010
         _ExtentX        =   8837
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Left            =   285
         TabIndex        =   1
         Top             =   210
         Width           =   4440
         _ExtentX        =   7832
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
         Caption         =   "REKENING"
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
         Left            =   285
         TabIndex        =   2
         Top             =   675
         Width           =   3180
         _ExtentX        =   5609
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
         Left            =   3870
         TabIndex        =   3
         Top             =   690
         Width           =   1995
         _ExtentX        =   3519
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
      Top             =   1185
      Width           =   9840
      _ExtentX        =   17357
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
         Left            =   8580
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
         Picture         =   "rptArusKas.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   7410
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
         Picture         =   "rptArusKas.frx":00A6
      End
   End
End
Attribute VB_Name = "rptArusKas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim db As New ADODB.Recordset
Dim vaArray As New XArrayDB

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  dDate(0).Value = Date
  dDate(1).Value = Date
  
  TabIndex cRekening, n
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cRekening_ButtonClick()
  Set db = objData.Pick(GetDSN, "Rekening", "Kode", cRekening, "Kode,Keterangan")
  If Not db.eof Then
    cNamaRekening.Text = GetNull(db!Keterangan, "")
  End If
End Sub

Private Sub cRekening_Validate(Cancel As Boolean)
  If cRekening.LastKey = 13 Then
    cRekening_ButtonClick
  End If
End Sub

Private Sub GetSQL()
Dim cSQL As String
Dim n As Double
Dim nDebet As Double
Dim nKredit As Double

  vaArray.ReDim 0, 0, 0, 5
  nDebet = 0
  nKredit = 0
  vaArray(0, 2) = "SALDO AWAL"
  cSQL = ""
  cSQL = "Select Sum(Awal) as Awal From SaldoRekening where Rekening = '" & cRekening.Text & "'"
  cSQL = cSQL & " union "
  cSQL = cSQL & "Select Sum(Debet-Kredit) as Awal From BukuBesar Where Tgl < '" & Format(dDate(0).Value, "yyyy-mm-dd") & "' and Rekening = '" & cRekening.Text & "'"
  
'  cSQL = cSQL & "Select Sum(Debet-Kredit) as Awal From BukuBesar Where Tgl < '" & Format(dDate(0).Value, "yyyy-mm-dd") & "' and rekening = '" & cRekening.Text & "' "

  Set dbData = objData.SQL(GetDSN, cSQL)
  vaArray(0, 5) = 0
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray(0, 5) = GetNull(vaArray(0, 5)) + GetNull((dbData!Awal))
      dbData.MoveNext
    Loop
  End If
    
  cSQL = " select `STATUS`,sum(debet) as debet,sum(KREDIT) as kredit from bukubesar"
  cSQL = cSQL & " where rekening like '" & cRekening.Text & "' and TGL >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " GROUP BY `STATUS`"
  
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.eof
      nDebet = 0
      nKredit = 0
      FrmPB.RunPB
      n = n + 1
      vaArray.InsertRows n
      vaArray(n, 0) = n
      vaArray(n, 1) = GetNull(dbData!status)
      vaArray(n, 2) = GetNamaKasRpt(vaArray(n, 1))
      vaArray(n, 3) = GetNull(dbData!Debet)
      vaArray(n, 4) = GetNull(dbData!Kredit)
      vaArray(n, 5) = GetNull(vaArray(n - 1, 5)) + GetNull(vaArray(n, 3)) - GetNull(vaArray(n, 4))
      nDebet = nDebet + vaArray(n, 3)
      nKredit = nKredit + vaArray(n, 4)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  rpt
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader "LAPORAN ARUS KAS", tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddPageHeader "REKENING", , , 15, True, , , , , True, , tdbPageHeaderSect
    .AddPageHeader " : " & " [ " & cRekening.Text & " ] " & cNamaRekening.Text
    .AddPageHeader "ANTARA TANGGAL", , , 15, True
    .AddPageHeader " : " & Format(dDate(0).Value, "dd-MM-yyyy") & " S.D " & Format(dDate(1).Value, "dd-MM-yyyy")
    
    .AddTableHeader "NO", , , , 9, , , , , , , , , , , , , , , False
    .AddTableHeader "KODE KAS", , , , 9, , , , , , , , , , , , , , , False
    .AddTableHeader "KETERANGAN"
    .AddTableHeader "DEBET", , , , 14
    .AddTableHeader "KREDIT", , , , 14
    .AddTableHeader "SALDO", , , , 14
    

    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
        
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Total", , tdbHalignCenter, , , , , , , , , , , , 1
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
        
    .Preview vaArray, True
  End With
End Sub

Private Function GetNamaKasRpt(ByVal n As Integer) As String
  Select Case n
    Case 1:
      GetNamaKasRpt = "Tabungan"
    Case 2:
      GetNamaKasRpt = "Realisasi Kredit"
    Case 3:
      GetNamaKasRpt = "Pembayaran Angsuran"
    Case 7:
      GetNamaKasRpt = "Deposito"
    Case 9:
      GetNamaKasRpt = "Voucher/Journal"
  End Select
End Function


