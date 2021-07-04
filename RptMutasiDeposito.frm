VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptMutasiDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN MUTASI DEPOSITO"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   7710
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
         Caption         =   "GOL DEPOSITO"
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
         Picture         =   "RptMutasiDeposito.frx":0000
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
         Picture         =   "RptMutasiDeposito.frx":00A6
      End
   End
End
Attribute VB_Name = "RptMutasiDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganDeposito", "Kode", cGolongan, "Kode,Keterangan")
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

Private Sub GetData()
Dim cField As String
Dim cWhere  As String
Dim vaJoin
Dim nCari As Integer
Dim n As Integer

  cField = "m.Faktur,m.Rekening,r.nama,r.alamat,m.kodemutasi, m.jumlah"
  vaJoin = Array("Left Join deposito d on d.Rekening = m.Rekening", _
                 "Left Join registerNasabah r on r.kode = d.kode")
  cWhere = " and m.Tgl <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "' "
  cWhere = cWhere & "and d.GolonganDeposito = '" & cGolongan.Text & "'"
  Set dbData = objData.Browse(GetDSN, "MutasiDeposito m", cField, "m.tgl", sisGTEqual, Format(dDate(0).Value, "yyyy-MM-dd"), cWhere, "m.Tgl,m.Rekening,m.ID", vaJoin)
  If Not dbData.eof Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    For n = 0 To vaArray.UpperBound(1)
      nCari = vaArray.Find(0, 4, vaArray(n, 4))
      If nCari >= 0 Then
        Select Case vaArray(n, 4)
          Case 1
            vaArray(n, 4) = "Pembukaan Deposito"
          Case 2
            vaArray(n, 4) = "Pencairan Pokok Deposito"
          Case 3
            vaArray(n, 4) = "Pencairan Bunga Deposito"
          Case 4
            vaArray(n, 4) = "Pinalty Deposito"
        End Select
      End If
    Next
    rpt
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cGolongan, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub rpt()
    With FrmRPT
      .AddPageHeader UCase("Laporan Mutasi Harian Deposito"), tdbHalignCenter, , , True, , 12, True
      .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 10
      .AddPageHeader "Antara Tanggal : " & dDate(0).Value & " s.d " & dDate(1).Value, tdbHalignCenter, , , True
      .AddPageHeader " ", , , , True
      .AddPageHeader " ", , , , True
      
      .AddTableHeader "No. Faktur", , , , 18, , , , , , True, tdbTableHeaderSect
      .AddTableHeader "No. Rekening", , , , 12
      .AddTableHeader "Nama", , , , 15
      .AddTableHeader "Alamat"
      .AddTableHeader "Keterangan Mutasi", , , , 20
      .AddTableHeader "Jml Mutasi", , , , 12
      
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody , , , , , , , , , , , , , False
      
      .AddTableFooter "Grand Total Mutasi", , tdbHalignRight, , , , , , , , , , , , 5
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
        
      .Preview vaArray, True
    End With
End Sub
