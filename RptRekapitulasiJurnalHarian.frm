VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptRekapitulasiJurnalHarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REKAPITULASI JURNAL HARIAN"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5895
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   990
      Left            =   0
      Top             =   0
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   1746
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
         Left            =   195
         TabIndex        =   0
         Top             =   195
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
         Left            =   3480
         TabIndex        =   1
         Top             =   210
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
      Top             =   975
      Width           =   5865
      _ExtentX        =   10345
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
         Left            =   4620
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
         Picture         =   "RptRekapitulasiJurnalHarian.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   3450
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
         Picture         =   "RptRekapitulasiJurnalHarian.frx":00A6
      End
   End
End
Attribute VB_Name = "RptRekapitulasiJurnalHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
  rpt
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  dDate(0).Value = BOM(Date)
  dDate(1).Value = EOM(Date)
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSQL()
Dim n As Integer
Dim vaField As String
Dim cWhere As String

  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 5
  
  vaField = "b.Rekening,r.Keterangan,sum(b.Debet) as Debet,sum(b.Kredit) as Kredit"
'  Set dbData = objData.Browse(GetDSN, "BukuBesar b", vaField, "b.Tgl", sisGTEqual, Format(dDate(0).Value, "yyyy-mm-dd"), "And b.Tgl <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "' Group by b.Rekening", "b.Rekening", _
               Array("Left Join Rekening r on b.Rekening = r.Kode", _
                     "Left Join SaldoRekening s on s.Rekening = b.Rekening"))
  Set dbData = objData.Browse(GetDSN, "BukuBesar b", vaField, "b.Tgl", sisGTEqual, Format(dDate(0).Value, "yyyy-mm-dd"), "And b.Tgl <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "' Group by b.Rekening", "b.Rekening", _
               Array("Left Join Rekening r on b.Rekening = r.Kode", _
                     "Left Join SaldoRekening s on s.Rekening = b.Rekening"))
  
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Rekening, "")
      vaArray(n, 1) = GetNull(dbData!Keterangan, "")
      vaArray(n, 2) = GetMutasi(vaArray(n, 0))
      vaArray(n, 3) = GetNull(dbData!Debet)
      vaArray(n, 4) = GetNull(dbData!Kredit)
      vaArray(n, 5) = SumRekening(vaArray(n, 0), vaArray(n, 2), vaArray(n, 3), vaArray(n, 4))
      dbData.MoveNext
     Loop
     FrmPB.EndPB
    Else
      MsgBox "Maaf, mutasi tgl " & Format(dDate(0).Value, "dd/MM/yy") & " - " & Format(dDate(1).Value, "dd/MM/yy" & " tidak ada"), vbInformation
      Exit Sub
    End If
    
End Sub

Private Function SumRekening(ByVal cRekening As String, ByVal nAwal As Double, ByVal nDebet As Double, ByVal nKredit As Double) As Double
  If left(cRekening, 1) = "1" Or left(cRekening, 1) = "5" Then
    SumRekening = nAwal + nDebet - nKredit
  Else
    SumRekening = nAwal - nDebet + nKredit
  End If
End Function

Private Sub rpt()
  With FrmRPT
    .AddPageHeader UCase("Laporan Rekapitulasi Jurnal Harian"), tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy") & " s.d " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , , , , , , , , , , 10
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "REKENING", , , , 11, , , , , , True, tdbTableHeaderSect, , tdbMergeOnText, , , , 5
    .AddTableHeader "KETERANGAN", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "SALDO AWAL", , , , 15, , , , , , , , , tdbMergeOnText
    .AddTableHeader "MUTASI", , , , 15, , , , , , , , , , 2
    .AddTableHeader "", , , , 15
    .AddTableHeader "SALDO AKHIR", , , , 15, , , , , , , , , tdbMergeOnText
    
    .AddTableHeader "REKENING", , , , 10, , , , , , True, tdbTableHeaderSect, , tdbMergeOnText
    .AddTableHeader "KETERANGAN", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "SALDO AWAL", , , , 15, , , , , , , , , tdbMergeOnText
    .AddTableHeader "DEBET", , , , 15
    .AddTableHeader "KREDIT", , , , 15
    .AddTableHeader "SALDO AKHIR", , , , 15, , , , , , , , , tdbMergeOnText
    
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .AddTableFooter "TOTAL", , tdbHalignCenter, , , , , , , , , , , , 2
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    
    .Preview vaArray, True
  End With
End Sub

Private Function GetMutasi(ByVal cRekening As String) As Double
Dim dbData1 As New ADODB.Recordset
Dim vaField As String

  If left(cRekening, 1) < 4 Then
    vaField = "sum(b.Debet) as debet,sum(b.Kredit)as kredit"
    Set dbData1 = objData.Browse(GetDSN, "BukuBesar b", vaField, "b.Rekening", sisAssign, cRekening, "And b.Tgl < '" & Format(dDate(0).Value, "yyyy-mm-dd") & "' Group by b.Rekening", "b.Rekening")
    If Not dbData1.eof Then
      GetMutasi = (dbData1!Debet) - (dbData1!Kredit)
    Else
      GetMutasi = 0
    End If
  Else
    vaField = "sum(b.Debet) as debet,sum(b.Kredit)as kredit"
    Set dbData1 = objData.Browse(GetDSN, "BukuBesar b", vaField, "b.Rekening", sisAssign, cRekening, " And b.Tgl < '" & Format(dDate(0).Value, "yyyy-mm-dd") & "'  and b.tgl >= '" & Format(BOY(dDate(0).Value), "yyyy-MM-dd") & "' Group by b.Rekening", "b.Rekening")
    If Not dbData1.eof Then
      GetMutasi = (dbData1!Debet) - (dbData1!Kredit)
    Else
      GetMutasi = 0
    End If
  End If
  
  Set dbData1 = objData.Browse(GetDSN, "SaldoRekening", "Sum(Awal) as Awal", "Rekening", sisAssign, cRekening)
  If Not dbData1.eof Then
    GetMutasi = GetMutasi + GetNull(dbData1!Awal)
  End If
  
  If Not (left(cRekening, 1) = "1" Or left(cRekening, 1) = "5") Then
    GetMutasi = -GetMutasi
  End If
  
End Function



