VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptMutasiharianKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN MUTASI HARIAN PINJAMAN"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   7725
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1155
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2037
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
         Caption         =   "GOL"
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
      Top             =   1140
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
         Picture         =   "RptMutasiharianKredit.frx":0000
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
         Picture         =   "RptMutasiharianKredit.frx":00A6
      End
   End
End
Attribute VB_Name = "RptMutasiharianKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim dbData1 As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

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
  dDate(0).Value = Date
  dDate(1).Value = Date
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cGolongan, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSAwal(ByVal cRekening As String, ByRef nSPokok As Double, ByRef nSBunga As Double)
  Set dbData1 = objData.Browse(GetDSN, "Angsuran", "sum(Pokok) as Pokok ,sum(Bunga) as Bunga", "Rekening", sisAssign, cRekening, "And Tgl < '" & Format(dDate(0).Value, "yyyy-mm-dd") & "' Group By Rekening", , "Rekening")
  If Not dbData1.eof Then
    nSPokok = GetNull(dbData1!pokok)
    nSBunga = GetNull(dbData1!bunga)
  End If
End Sub

Private Sub GetData()
Dim cSQL As String
Dim n As Integer
Dim nSPokok As Double
Dim nSBunga As Double

    vaArray.ReDim 0, -1, 0, 8
    cSQL = "Select a.Faktur,d.Rekening,r.Nama,a.Tgl,d.Plafond,d.TotalBunga,a.Pokok,a.Bunga"
    cSQL = cSQL & " From Angsuran a"
    cSQL = cSQL & " Left Join Debitur d on d.Rekening = a.Rekening"
    cSQL = cSQL & " Left Join RegisterNasabah r on r.Kode = d.Kode"
    cSQL = cSQL & " Where d.GolonganKredit ='" & cGolongan.Text & "'"
    cSQL = cSQL & " And a.Tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "'"
    cSQL = cSQL & " And a.Tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
    cSQL = cSQL & " Order By a.Faktur"
    Set dbData = objData.SQL(GetDSN, cSQL)
    If Not dbData.eof Then
      dbData.MoveFirst
      FrmPB.InitPB dbData.RecordCount + 1
      Do While Not dbData.eof
        FrmPB.RunPB
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        
        vaArray(n, 0) = GetNull(dbData!Rekening, "")
        vaArray(n, 1) = GetNull(dbData!nama, "")
        vaArray(n, 2) = GetNull(dbData!plafond, "")
        GetSAwal vaArray(n, 0), nSPokok, nSBunga
        vaArray(n, 3) = nSPokok
        vaArray(n, 4) = nSBunga
        vaArray(n, 5) = (dbData!pokok)
        vaArray(n, 6) = (dbData!bunga)
        vaArray(n, 7) = vaArray(n, 2) - (vaArray(n, 3) + vaArray(n, 5))
        vaArray(n, 8) = (dbData!totalBunga) - vaArray(n, 4) + vaArray(n, 6)
        dbData.MoveNext
      Loop
      FrmPB.EndPB
      rpt
    End If
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader "LAPORAN MUTASI HARIAN PINJAMAN", tdbHalignCenter, , , , , 12, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy") & " s/d " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "No. Rekening", , , , 12, , , , , , True, tdbTableHeaderSect, , tdbMergeOnText
    .AddTableHeader "Nama", , , , 20, , , , , , , , , tdbMergeOnText
    .AddTableHeader "Plafond", , , , 11, , , , , , , , , tdbMergeOnText
    .AddTableHeader "Saldo Awal", , , , 11, , , , , , , , , , 2
    .AddTableHeader "", , , , 11
    .AddTableHeader "Mutasi", , , , 11, , , , , , , , , , 2
    .AddTableHeader "", , , , 11
    .AddTableHeader "Saldo Akhir", , , , 11, , , , , , , , , , 2
    .AddTableHeader "", , , , 11
    
    .AddTableHeader "No. Rekening", , , , 12, , , , , , True, tdbTableHeaderSect, , tdbMergeOnText
    .AddTableHeader "Nama", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "Plafond", , , , 10, , , , , , , , , tdbMergeOnText
    .AddTableHeader "Pokok", , , , 11
    .AddTableHeader "Bunga", , , , 11
    .AddTableHeader "Pokok", , , , 11
    .AddTableHeader "Bunga", , , , 11
    .AddTableHeader "Pokok", , , , 11
    .AddTableHeader "Bunga", , , , 11

    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    
    .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 2
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .Preview vaArray, True, , True
  End With
End Sub
