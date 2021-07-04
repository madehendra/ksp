VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Begin VB.Form rptDaftarAnggota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Anggota"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   6735
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   0
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
      Begin VB.OptionButton optAnggota 
         Caption         =   "Seluruhnya"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   2895
         TabIndex        =   4
         Top             =   180
         Width           =   1305
      End
      Begin VB.OptionButton optAnggota 
         Caption         =   "Non Anggota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1455
         TabIndex        =   3
         Top             =   180
         Width           =   1305
      End
      Begin VB.OptionButton optAnggota 
         Caption         =   "Anggota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   390
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5520
         TabIndex        =   0
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
         Picture         =   "rptDaftarAnggota.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4350
         TabIndex        =   1
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
         Picture         =   "rptDaftarAnggota.frx":00A6
      End
   End
End
Attribute VB_Name = "rptDaftarAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim dbData1 As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  TabIndex optAnggota(0), n
  TabIndex optAnggota(1), n
  TabIndex optAnggota(0), n
  TabIndex optAnggota(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
  End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub GetSQL()
Dim cField As String
Dim vaJoin
Dim cWhere As String
Dim n As Integer

  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 3
  cField = "r.Kode,r.Nama,p.Keterangan,r.Alamat,r.Telepon"
  vaJoin = Array("Left Join Pekerjaan p on p.Kode = r.Pekerjaan")
  
  cWhere = ""
  If optAnggota(0).Value = True Then
    cWhere = " jenisanggota = '1'"
  ElseIf optAnggota(1).Value = True Then
    cWhere = " jenisanggota = '2'"
  End If
  
  Set dbData = objData.Browse(GetDSN, "RegisterNasabah r", cField, , , , cWhere, "r.Kode", vaJoin)
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull(dbData!Kode, "")
      vaArray(n, 1) = GetNull(dbData!nama, "")
      vaArray(n, 2) = GetNull(dbData!alamat, "")
      vaArray(n, 3) = GetNull(dbData!Telepon, "")
      dbData.MoveNext
    Loop
    GetRpt
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Sub
  End If
End Sub

Private Sub GetRpt()
Dim cHeader As String

  With FrmRPT
    If optAnggota(0).Value = True Then
      cHeader = "Daftar Anggota"
    ElseIf optAnggota(1).Value = True Then
      cHeader = "Daftar Non Anggota"
    Else
      cHeader = "Daftar Seluruh Anggota"
    End If
    
    .AddPageHeader UCase(cHeader), tdbHalignCenter, , , , , 11, True
    .AddPageHeader aCfg(msNama), tdbHalignCenter, , , True, , 14, True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "Kode", , , , 10, , , , , , , , , , , , , 5
    .AddTableHeader "Nama Nasabah", , , , 25
    .AddTableHeader "Alamat"
    .AddTableHeader "Telepon", , , , 15
    
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    
    .Preview vaArray, True
  End With
End Sub



