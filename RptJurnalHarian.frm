VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptJurnalHarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN JURNAL HARIAN"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5940
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   750
      Left            =   0
      Top             =   0
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   1323
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   195
         Width           =   3180
         _ExtentX        =   5609
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
         Left            =   3465
         TabIndex        =   1
         Top             =   210
         Width           =   1995
         _ExtentX        =   3519
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
      Top             =   735
      Width           =   5910
      _ExtentX        =   10425
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
         Left            =   4665
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
         Picture         =   "RptJurnalHarian.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   3495
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
         Picture         =   "RptJurnalHarian.frx":00A6
      End
   End
End
Attribute VB_Name = "RptJurnalHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub SQL()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 6
  Set dbData = objData.Browse(GetDSN, "BukuBesar b", "b.Rekening,r.Keterangan as namaRekening,b.Faktur,b.Tgl,b.Keterangan as KeteranganBukuBesar,b.Debet,b.Kredit", "b.Tgl", sisGTEqual, Format(dDate(0).Value, "yyyy-mm-dd"), _
                            " and b.Tgl <= '" & Format(dDate(1).Value, "yyyy-mm-dd") & "'", "b.Rekening,b.Tgl,b.ID", _
                            Array("Left Join Rekening r on b.Rekening = r.Kode"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Rekening, "")
      vaArray(n, 1) = GetNull(dbData!NamaRekening, "")
      vaArray(n, 2) = GetNull(dbData!Faktur, "")
      vaArray(n, 3) = GetNull(dbData!Tgl, "")
      vaArray(n, 4) = GetNull(dbData!KeteranganBukuBesar, "")
      vaArray(n, 5) = GetNull(dbData!Debet)
      vaArray(n, 6) = GetNull(dbData!Kredit)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    GetRpt
  Else
    MsgBox "Data tidak ada", vbInformation
  End If
End Sub

Private Sub GetRpt()
  With FrmRPT
    .AddPageHeader "LAPORAN JURNAL HARIAN", tdbHalignCenter, , , , , 12, True
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy") & " S.D " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableGroupHeader True, "[]", , , , 12
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "NO. FAKTUR", , , , 18, , , , , , True, tdbTableHeaderSect
    .AddTableHeader "TANGGAL", , , , 8
    .AddTableHeader "KETERANGAN"
    .AddTableHeader "DEBET", , , , 15
    .AddTableHeader "KREDIT", , , , 15
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody , , , , , , , , , , , , , False
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "TOTAL", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2

    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "GRAND TOTAL", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
    End With
End Sub

Private Sub cmdPreview_Click()
  SQL
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  dDate(0).Value = Date
  dDate(1).Value = Date
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub


