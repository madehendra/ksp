VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptRekapBungaTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN REKAPITULASI BUNGA TABUNGAN"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   7725
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1065
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   1879
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   315
         TabIndex        =   0
         Top             =   150
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGolongan 
         Height          =   330
         Left            =   3045
         TabIndex        =   1
         Top             =   525
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
         TabIndex        =   2
         Top             =   525
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
         Caption         =   "GOL TABUNGAN"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   3780
         TabIndex        =   3
         Top             =   150
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
      Top             =   1050
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6390
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
         Picture         =   "RptRekapBungaTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5220
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
         Picture         =   "RptRekapBungaTabungan.frx":00A6
      End
   End
End
Attribute VB_Name = "RptRekapBungaTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New BiSAMyDLL.data
Dim vaArray As New XArrayDB
Dim nBesarPajakBunga As Double

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganTabungan", "Kode", cGolongan, "Kode,Keterangan,PajakBunga")
  If Not dbData.eof Then
    cNamaGolongan.Text = dbData!Keterangan
    nBesarPajakBunga = dbData!PajakBunga
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Or Trim(cGolongan.Text) <> "" Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader UCase("Laporan Bunga Tabungan") & cNamaGolongan.Text, tdbHalignCenter, , , , , 12, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "Antara Tanggal :" & Format(dTgl(0).Value, "dd-MM-yyyy") & " s.d " & Format(dTgl(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 9, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "REKENING", , , , 12
    .AddTableHeader "NAMA NASABAH"
    .AddTableHeader "SALDO TAB.", , , , 14
    .AddTableHeader "BUNGA", , , , 12
    .AddTableHeader "PAJAK BUNGA", , , , 10
    .AddTableHeader "SALDO AKHIR", , , , 14

    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2

    .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 2
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2

    .Preview vaArray, True
  End With

End Sub

Private Sub GetSQL()
Dim cWhere As String
Dim n As Double
Dim dTanggalTutup
Dim nTotalBunga As Double
Dim nTotalPajak As Double
Dim nSaldoTabungan As Double
Dim i As Integer

  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 5
  objData.OpenConnection GetDSN
  
  cWhere = cWhere & " t.GolonganTabungan = '" & cGolongan.Text & "' "
  cWhere = cWhere & " and t.Close <>'1'" ' and t.rekening = '02.11.000002.01'"
  Set dbData = objData.Browse(GetDSN, "Tabungan t", "t.Rekening,r.Nama,r.Alamat,t.Awal,t.Tgl,t.TglPenutupan,t.Close", , , , cWhere, "t.Rekening", _
               Array("Left Join RegisterNasabah r on t.Kode = r.Kode"))
                     
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = (dbData!Rekening)
      vaArray(n, 1) = (dbData!Nama)
      
      GetBungaHarian objData, dbData!Rekening, dTgl(0).Value, dTgl(1).Value, nTotalBunga, nTotalPajak, nSaldoTabungan
      
      vaArray(n, 2) = nSaldoTabungan
      vaArray(n, 3) = nTotalBunga
      vaArray(n, 4) = nTotalPajak
      vaArray(n, 5) = vaArray(n, 2) + vaArray(n, 3) - vaArray(n, 4)
      dbData.MoveNext
     Loop
     FrmPB.EndPB
     rpt
    Else
      MsgBox "Data tidak ada", vbInformation
      Exit Sub
    End If
    objData.CloseConnection GetDSN
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = EOM(Date)
  
      
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cGolongan, n
  TabIndex cmdPreview, n
End Sub




