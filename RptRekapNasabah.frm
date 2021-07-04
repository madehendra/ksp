VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptRekapNasabah 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN REKAPITULASI NASABAH"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   6795
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   960
      Left            =   0
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   1693
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
         Left            =   390
         TabIndex        =   0
         Top             =   285
         Width           =   3270
         _ExtentX        =   5768
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
         Left            =   3675
         TabIndex        =   1
         Top             =   285
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
      Top             =   960
      Width           =   6780
      _ExtentX        =   11959
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
         Left            =   5520
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
         Picture         =   "RptRekapNasabah.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4350
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
         Picture         =   "RptRekapNasabah.frx":00A6
      End
   End
End
Attribute VB_Name = "RptRekapNasabah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim dbTabungan As New ADODB.Recordset
Dim dbDeposito As New ADODB.Recordset
Dim dbKredit As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
  dDate(0).Value = BOM(Date)
  dDate(1).Value = EOM(Date)
End Sub

Private Sub GetSQL()
Dim cWhere As String
Dim cSQL As String
Dim n, i As Integer
Dim nTabAktif, nTabNonAKtif As Integer
Dim nDepAktif, nDepNonAktif As Integer
Dim nPembAktif, nPembNonAktif As Integer

  
  vaArray.ReDim 0, 0, 0, 12
  Set dbData = objData.Browse(GetDSN, "RegisterNasabah", "Kode,Nama,TglRegister", "TglRegister", sisGTEqual, Format(dDate(0).Value, "yyyy-MM-dd"), "And Tglregister <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "'", "Kode,Tglregister")
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount + 1
    
    n = 0
    nTabAktif = 0
    nTabNonAKtif = 0
    nDepAktif = 0
    nDepNonAktif = 0
    nPembAktif = 0
    nPembNonAktif = 0
     
    vaArray.ReDim 0, dbData.RecordCount - 1, 0, 12
    dbData.MoveFirst
    Do While Not dbData.eof
        FrmPB.RunPB
        Set dbTabungan = objData.Browse(GetDSN, "Tabungan", "Kode,Close", "Kode", sisAssign, dbData!Kode)
        Do While Not dbTabungan.eof
          If dbTabungan!Close = "" Then
            nTabAktif = nTabAktif + 1
          ElseIf dbTabungan!Close = "1" Then
            nTabNonAKtif = nTabNonAKtif + 1
          End If
          dbTabungan.MoveNext
         Loop
        
        Set dbDeposito = objData.Browse(GetDSN, "Deposito", "Kode,Status", "Kode", sisAssign, dbData!Kode)
        Do While Not dbDeposito.eof
          If dbDeposito!status = "" Then
            nDepAktif = nDepAktif + 1
          ElseIf dbDeposito!status = "1" Then
            nDepNonAktif = nDepNonAktif + 1
          End If
          dbDeposito.MoveNext
        Loop
        
        Set dbKredit = objData.SQL(GetDSN, "Select Kode,Status From Debitur Where Kode='" & dbData!Kode & "'")
        Do While Not dbKredit.eof
          If dbKredit!status = "" Then
            nPembAktif = nPembAktif + 1
          ElseIf dbKredit!status = "1" Then
            nPembNonAktif = nPembNonAktif + 1
          End If
          dbKredit.MoveNext
        Loop
      
      vaArray(n, 0) = (dbData!Kode)
      vaArray(n, 1) = (dbData!nama)
      vaArray(n, 2) = GetNull(dbTabungan.RecordCount, 0)
      vaArray(n, 3) = GetNull(nTabAktif, 0)
      vaArray(n, 4) = GetNull(nTabNonAKtif, 0)
      vaArray(n, 5) = GetNull(dbDeposito.RecordCount, 0)
      vaArray(n, 6) = GetNull(nDepAktif, 0)
      vaArray(n, 7) = GetNull(nDepNonAktif, 0)
      vaArray(n, 8) = GetNull(dbKredit.RecordCount, 0)
      vaArray(n, 9) = GetNull(nPembAktif, 0)
      vaArray(n, 10) = GetNull(nPembNonAktif, 0)
      
      nTabAktif = 0
      nTabNonAKtif = 0
      nDepAktif = 0
      nDepNonAktif = 0
      nPembAktif = 0
      nPembNonAktif = 0
      
     dbData.MoveNext
     n = n + 1
    Loop
    FrmPB.EndPB
    rpt
  End If
  
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader UCase("Laporan Rekapitulasi Nasabah"), tdbHalignCenter, , , , , 14, True, True
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy") & " s.d " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , , , , , , , , , , 10
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "No. Register", , , , 10, , , , , , True, tdbTableHeaderSect, , tdbMergeOnText, , , , 5
    .AddTableHeader "Nama Nasabah", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "TABUNGAN", , , , 6, , , , , , , , , , 3
    .AddTableHeader "", , , , 6
    .AddTableHeader "", , , , 6
    .AddTableHeader "DEPOSITO", , , , 6, , , , , , , , , , 3
    .AddTableHeader "", , , , 6
    .AddTableHeader "", , , , 6
    .AddTableHeader "KREDIT", , , , 6, , , , , , , , , , 3
    .AddTableHeader "", , , , 6
    .AddTableHeader "", , , , 6
    
    .AddTableHeader "No. Register", , , , 10, , , , , , True, tdbTableHeaderSect, , tdbMergeOnText
    .AddTableHeader "Nama Nasabah", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "Jumlah", , , , 6
    .AddTableHeader "Aktif", , , , 6
    .AddTableHeader "Non Aktif", , , , 6
    .AddTableHeader "Jumlah", , , , 6
    .AddTableHeader "Aktif", , , , 6
    .AddTableHeader "Non Aktif", , , , 6
    .AddTableHeader "Jumlah", , , , 6
    .AddTableHeader "Aktif", , , , 6
    .AddTableHeader "Non Aktif", , , , 6
    
    .AddTableBody
    .AddTableBody
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignRight
    .AddTableBody , tdbHalignRight
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    
    .AddTableFooter "Total", , tdbHalignCenter, , , , , , , , , , , , 2
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    
    .Preview vaArray, True
  End With
End Sub
