VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form RptBungaTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN BUNGA SIMPANAN"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5205
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   645
      Left            =   0
      Top             =   1020
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   1138
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
         Left            =   3840
         TabIndex        =   3
         Top             =   120
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
         Picture         =   "RptBungaTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   2670
         TabIndex        =   4
         Top             =   120
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
         Picture         =   "RptBungaTabungan.frx":00A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1020
      Left            =   0
      Top             =   0
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   1799
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Left            =   1245
         TabIndex        =   0
         Top             =   165
         Width           =   2715
         _ExtentX        =   4789
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
         Caption         =   "PERIODE"
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
      Begin BiSANumberBoxProject.BiSANumberBox nBulan 
         Height          =   330
         Left            =   1245
         TabIndex        =   1
         Top             =   540
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         Decimals        =   0
         DecimalPoint    =   ""
         Separator       =   ""
         MaxValue        =   12
         MinValue        =   1
         Enabled         =   0   'False
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
         Caption         =   "BULAN / TAHUN"
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
      Begin BiSANumberBoxProject.BiSANumberBox nTahun 
         Height          =   330
         Left            =   3345
         TabIndex        =   2
         Top             =   540
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   582
         Decimals        =   0
         DecimalPoint    =   ""
         Separator       =   ""
         MaxValue        =   9999
         MinValue        =   1
         Enabled         =   0   'False
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
End
Attribute VB_Name = "RptBungaTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cKode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "PeriodeBungaTabungan", "Kode,Bulan,Tahun", "Kode", sisContent, cKode.Text, "And Status = '1'", "Kode")
  cKode.Text = cKode.Browse(dbData)
  If Not dbData.eof Then
    nBulan.Value = GetNull(dbData!Bulan)
    nTahun.Value = GetNull(dbData!Tahun)
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
  nBulan.Value = 0
  nTahun.Value = 0
  cKode.Default
  vaArray.ReDim 0, -1, 0, 6
  
  TabIndex cKode, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetPreview()
  With FrmRPT
    .AddPageHeader "LAPORAN BUNGA SIMPANAN", tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader aCfg(msNama), tdbHalignCenter, , , True
    .AddPageHeader "Periode Bulan/Tahun : " & nBulan.Value & "/" & nTahun.Value, tdbHalignCenter, , , True
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
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "REKENING", , , , 13
    .AddTableHeader "TGL POSTING", , , , 8
    .AddTableHeader "NAMA NASABAH"
    .AddTableHeader "SALDO MIN.", , , , 12
    .AddTableHeader "BUNGA", , , , 12
    .AddTableHeader "PAJAK BUNGA", , , , 12
    .AddTableHeader "TOTAL BUNGA", , , , 12
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "Sub Total", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&SUM", Sis_Rpt_Number2, tdbHalignRight
    .AddTableGroupFooter "&SUM", Sis_Rpt_Number2, tdbHalignRight
    .AddTableGroupFooter "&SUM", Sis_Rpt_Number2, tdbHalignRight
    .AddTableGroupFooter "&SUM", Sis_Rpt_Number2, tdbHalignRight
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "TOTAL", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&sum", Sis_Rpt_Number2
    .AddTableFooter "&sum", Sis_Rpt_Number2
    .AddTableFooter "&sum", Sis_Rpt_Number2
    .AddTableFooter "&sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub

Private Sub GetData()
Dim n As Integer
Dim cField As String
  
  vaArray.ReDim 0, -1, 0, 8
  
  cField = "t.Pdl,p.Rekening,p.Tanggal,p.SaldoMinimal,p.Bunga,p.Pajak,p.TotalBunga,r.Nama,d.Keterangan as NamaPDL"
  Set dbData = objData.Browse(GetDSN, "PostingBungaTabungan p", cField, "p.Kode", sisAssign, cKode.Text, "And t.Close <>'1'", "t.Pdl,p.Rekening", _
                              Array("Left Join Tabungan t on t.Rekening=p.Rekening", _
                                    "Left Join registernasabah r on r.Kode = t.Kode", _
                                    "Left join Pdl d on d.Kode=t.Pdl"))
                                    
  If Not dbData.eof Then
    dbData.MoveFirst
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull(dbData!PDL, "")
      vaArray(n, 1) = GetNull(dbData!namapdl)
      vaArray(n, 2) = GetNull(dbData!Rekening, "")
      vaArray(n, 3) = GetNull(dbData!Tanggal)
      vaArray(n, 4) = GetNull(dbData!nama, "")
      vaArray(n, 5) = GetNull(dbData!SaldoMinimal)
      vaArray(n, 6) = GetNull(dbData!bunga)
      vaArray(n, 7) = GetNull(dbData!Pajak)
      vaArray(n, 8) = GetNull(dbData!totalBunga)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    GetPreview
  End If
  
End Sub
