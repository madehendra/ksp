VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptJatuhtempoKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN PINJAMAN JATUH TEMPO"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7725
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1590
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2805
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
         Caption         =   "TGL JATUH TEMPO"
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
         Visible         =   0   'False
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
      Begin BiSATextBoxProject.BiSABrowse cAO 
         Height          =   330
         Index           =   0
         Left            =   315
         TabIndex        =   6
         Top             =   960
         Width           =   3030
         _ExtentX        =   5345
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
         Caption         =   "AO"
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
         Left            =   3405
         TabIndex        =   7
         Top             =   975
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
      Top             =   1590
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
         Picture         =   "RptJatuhtempoKredit.frx":0000
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
         Picture         =   "RptJatuhtempoKredit.frx":00A6
      End
   End
End
Attribute VB_Name = "RptJatuhtempoKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim cSQL As String

Private Sub cAO_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "ao", "kode,nama")
  If Not dbData.eof Then
    cAO(Index).Text = cAO(Index).Browse(dbData)
  End If
End Sub

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

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  GetMinMax "AO", cAO, "Kode"
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cGolongan, n
  TabIndex cAO(0), n
  TabIndex cAO(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Function isSudahBayar(ByVal Rekening As String, ByVal dTglCek As Date, ByRef dTglBayarTerkhir As Date, ByVal dTglRealisasi As Date) As Boolean
Dim db As New ADODB.Recordset

  isSudahBayar = False
  Set db = objData.Browse(GetDSN, "angsuran", , "rekening", sisAssign, Rekening, , "rekening,tgl desc", , 0, 1)
  If Not db.eof Then
    'jika tgl yg besangkutan ada dalam rentang bulan dari tgl yg dipilih maka artinya sudah lunas
    If GetNull(db!Tgl) <= Format(dTglCek, "yyyy-MM-dd") And GetNull(db!Tgl) >= BOM(Format(dTglCek, "yyyy-MM-dd")) Then
      isSudahBayar = True
    End If
    dTglBayarTerkhir = GetNull(db!Tgl)
  Else
    If Format(DateAdd("d", 31, dTglRealisasi), "yyyy-MM-dd") > Format(dTglCek, "yyyy-MM-dd") Then
      isSudahBayar = True
    Else
      dTglBayarTerkhir = dTglRealisasi
    End If
  End If
End Function

Private Sub GetData()
Dim n As Integer
Dim dTanggalCair As Date
Dim cField As String
Dim vaJoin
Dim cWhere As String
Dim dTglLasAngsur As Date
Dim dTglReal As Date

  vaArray.ReDim 0, -1, 0, 9
  cField = "d.AO,d.Rekening, d.Plafond, d.Tgl, d.SukuBunga, d.Jatuhtempo,d.Status,r.Nama,d.lama,a.Nama as NamaAO"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode=d.Kode", _
                 "Left Join AO a on a.Kode = d.AO")
                                
'  cWhere = "And d.JatuhTempo >='" & Format(dDate(0).Value, "yyyy-MM-dd") & "'"
'  cWhere = cWhere & "And d.JatuhTempo <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  
  cWhere = cWhere & " And d.Status <> '1' and d.statuspencairan = '1'"
  If Trim(cAO(0).Text) <> "" Then
    cWhere = cWhere & " and d.AO = '" & cAO(0).Text & "'"
  End If
  Set dbData = objData.Browse(GetDSN, "Debitur d", cField, "d.GolonganKredit", sisAssign, cGolongan.Text, cWhere, "d.ao,d.rekening", vaJoin)
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = (dbData!AO)
      vaArray(n, 1) = (dbData!namaao)
      vaArray(n, 2) = (dbData!Rekening)
      vaArray(n, 3) = (dbData!nama)
      vaArray(n, 4) = (dbData!plafond)
      vaArray(n, 5) = (dbData!Lama)
      vaArray(n, 6) = (dbData!SukuBunga)
      vaArray(n, 7) = GetBK(vaArray(n, 2), vaArray(n, 4))
      vaArray(n, 8) = Format((dbData!Tgl), "dd/MM/yyyy")
      
'      vaArray(n, 9) = (dbData!JatuhTempo)
      
      If isSudahBayar(vaArray(n, 2), dDate(0).Value, dTglLasAngsur, GetNull(dbData!Tgl)) = True Then
        vaArray.DeleteRows n
      Else
        vaArray(n, 9) = Format(dTglLasAngsur, "dd-MM-yyyy")
      End If
      
      dbData.MoveNext
    Loop
    rpt
  End If
End Sub

Private Function GetBK(ByVal cRek As String, ByVal nPlafond As Double) As Double
Dim dbBK As New ADODB.Recordset
  GetBK = nPlafond
  Set dbBK = objData.Browse(GetDSN, "Angsuran", "Sum(Pokok) as Pokok", "Rekening", sisAssign, cRek, "And Tgl <='" & Format(dDate(1).Value, "yyyy-mm-dd") & "' Group By Rekening", "Rekening")
  If Not dbBK.eof Then
    GetBK = nPlafond - GetNull(dbBK!pokok)
  End If
End Function

Private Sub rpt()
  With FrmRPT
    .AddPageHeader "LAPORAN PINJAMAN JATUH TEMPO", tdbHalignCenter, , , , , 12, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 10
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy") & " s.d " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 10
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
    .AddTableHeader "No. Rekening", , , , 13
    .AddTableHeader "Nama"
    .AddTableHeader "Plafond", , , , 13
    .AddTableHeader "Lama", , , , 7
    .AddTableHeader "Suku Bunga", , , , 7
    .AddTableHeader "Baki Debet", , , , 13
    .AddTableHeader "Tgl. Realisasi", , , , 10
    .AddTableHeader "Ang. Terakhir", , , , 10
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody , tdbHalignCenter
    .AddTableBody Sis_Rpt_Number, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Total", , tdbHalignCenter, , , , , , , , , , , , 5
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    .AddTableFooter
    
'    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
'    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
'    .AddTableGroupFooter "Sub Total", , tdbHalignRight, , , , , , , , , , , , 5
'    .AddTableGroupFooter
'    .AddTableGroupFooter
'    .AddTableGroupFooter
'    .AddTableGroupFooter
'    .AddTableGroupFooter "&SUM", Sis_Rpt_Number2, tdbHalignRight
'    .AddTableGroupFooter
'    .AddTableGroupFooter
    
    .Preview vaArray, True
  End With
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub
