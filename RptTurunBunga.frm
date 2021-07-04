VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptTurunBunga 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN TURUN BUNGA"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   7695
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1500
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2646
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
         Caption         =   "GOLONGAN"
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
         Caption         =   "TANGGAL"
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
      Begin BiSATextBoxProject.BiSABrowse cAO 
         Height          =   330
         Index           =   0
         Left            =   315
         TabIndex        =   5
         Top             =   960
         Width           =   3060
         _ExtentX        =   5398
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
         Caption         =   "ANTARA AO"
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
      Begin BiSATextBoxProject.BiSABrowse cAO 
         Height          =   330
         Index           =   1
         Left            =   3525
         TabIndex        =   6
         Top             =   960
         Width           =   1845
         _ExtentX        =   3254
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
      Top             =   1485
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
         Picture         =   "RptTurunBunga.frx":0000
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
         Picture         =   "RptTurunBunga.frx":00A6
      End
   End
End
Attribute VB_Name = "RptTurunBunga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cAO_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "AO", "Kode,Nama", "Kode", sisContent, cAO(Index).Text, , "Kode")
  cAO(Index).Text = cAO(Index).Browse(dbData)
End Sub

Private Sub cAO_Validate(Index As Integer, Cancel As Boolean)
  If cAO(Index).LastKey = 13 Or cAO(Index).LastKey = 40 Then
    cAO_ButtonClick (Index)
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganKredit", "Kode", cGolongan, "Kode,Keterangan")
  If Not dbData.eof Then
    cNamaGolongan.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Or cGolongan.LastKey = 40 Then
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
  dDate.Value = Date
  GetMinMax "AO", cAO, "Kode"
  
  TabIndex dDate, n
  TabIndex cGolongan, n
  TabIndex cAO(0), n
  TabIndex cAO(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim n As Double
Dim cField As String
Dim cWhere As String
Dim vaJoin

  vaArray.ReDim 0, -1, 0, 10
  cField = "d.Tgl,d.rekening,d.Lama,d.plafond,d.SukuBunga,d.CaraAngsuran,d.AO,r.Nama as NamaDebitur,a.Nama as NamaAO"
  cWhere = " And d.statuspencairan = '1'"
  cWhere = cWhere & " And d.AO >= '" & cAO(0).Text & "'"
  cWhere = cWhere & " And d.AO <= '" & cAO(1).Text & "'"
  cWhere = cWhere & " And d.Status <> '1'"
  cWhere = cWhere & " And d.Tgl < '" & Format(dDate.Value, "yyyy-mm-dd") & "'"
  vaJoin = Array("Left Join RegisterNasabah r on d.Kode = r.Kode", _
                 "Left Join AO a on a.Kode = d.AO")
  Set dbData = objData.Browse(GetDSN, "Debitur d", cField, "GolonganKredit", sisAssign, cGolongan.Text, cWhere, "d.Rekening", vaJoin)
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount + 1
    dbData.MoveFirst
    Do While Not dbData.eof
      FrmPB.RunPB
      If GetTempo(GetNull(dbData!Tgl), GetNull(dbData!Lama), GetNull(dbData!CaraAngsuran)) Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = GetNull(dbData!AO)
        vaArray(n, 1) = GetNull(dbData!namaao)
        vaArray(n, 2) = GetNull(dbData!Rekening)
        vaArray(n, 3) = GetNull(dbData!NamaDebitur)
        vaArray(n, 4) = GetNull(dbData!Tgl)
        vaArray(n, 5) = GetNull(dbData!plafond)
        vaArray(n, 6) = GetNull(dbData!Lama)
        vaArray(n, 7) = Round(GetNull(dbData!SukuBunga) / 12, 2)
        vaArray(n, 8) = IIf(GetNull(dbData!CaraAngsuran) = "H", "Harian", "Bulanan")
        vaArray(n, 9) = GetBakiDebet(objData, vaArray(n, 2), vaArray(n, 5), dDate.Value)
        vaArray(n, 10) = GetBungaReguler(vaArray(n, 9), Round(vaArray(n, 7) / 12, 2))
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    rpt
  Else
    MsgBox "Data tidak ada..", vbInformation
    Exit Sub
  End If
End Sub

Private Function GetBungaReguler(ByVal nSisaPokok As Double, ByVal nBunga As Double) As Double
  GetBungaReguler = nSisaPokok * (nBunga / 100)
  GetBungaReguler = Mod50(GetBungaReguler)
End Function

Private Sub rpt()
  With FrmRPT
    .AddPageHeader "LAPORAN TURUN BUNGA PINJAMAN", tdbHalignCenter, , , , , 12, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "Tanggal : " & Format(dDate.Value, "dd-MM-yyyy"), tdbHalignCenter, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableGroupHeader True, "[]", , , 10
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
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
    .AddTableHeader "No. Rekening", , , , 10
    .AddTableHeader "Nama Debitur"
    .AddTableHeader "Tgl Real.", , , , 8
    .AddTableHeader "Plafond", , , , 10
    .AddTableHeader "Lama", , , , 5
    .AddTableHeader "Suku Bunga/Bln", , , , 10
    .AddTableHeader "Harian/Bulanan", , , , 10
    .AddTableHeader "Baki Debet", , , , 10
    .AddTableHeader "Turun Bunga", , , , 10
        
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .Preview vaArray, True, , True
  End With
End Sub

Private Function GetTempo(ByVal dTglRealisasi As Date, ByVal nLama As Integer, ByVal cCaraAngsuran As String) As Boolean
Dim n As Single
Dim dTanggal As Date
Dim xArray As New XArrayDB
  
  GetTempo = False
  xArray.ReDim 0, nLama, 0, 0
  dTanggal = DateAdd("m", 1, dTglRealisasi)
  If cCaraAngsuran = "H" Then
    dTanggal = DateAdd("d", 1, dTanggal)
  End If
  
  For n = 1 To nLama
    xArray(n, 0) = dTanggal
    If dTanggal = dDate.Value Then
      GetTempo = True
      Exit For
    End If
    dTanggal = DateAdd("m", 1, xArray(n, 0))
  Next
End Function
