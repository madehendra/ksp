VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptTagihanKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN DAFTAR TAGIHAN KREDIT"
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
      BorderStyle     =   4
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
         Caption         =   "SAMPAI TANGGAL"
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
      BorderStyle     =   4
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
         Picture         =   "RptTagihanKredit.frx":0000
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
         Picture         =   "RptTagihanKredit.frx":00A6
      End
   End
End
Attribute VB_Name = "RptTagihanKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New BisaMyDLL.data
Dim dbData As New ADODB.Recordset
Dim nKe As Single
Dim cField As String
Dim vaJoin
Dim vaarray As New XArrayDB
Dim i As Double
Dim nBunga As Double
Dim nPokok As Double

Private Sub GetDataAngsuran()
Dim dbAngsuran As New ADODB.Recordset

  Set dbAngsuran = objData.Browse(GetDSN, "Angsuran", "Count(Rekening) as nKe", "rekening", sisAssign, dbData!Rekening, " and (Pokok > 0 or bunga > 0)")
  If Not dbAngsuran.eof Then
    nKe = GetNull(dbAngsuran!nKe) + 1
  Else
    nKe = 1
  End If
  nBunga = GetAngsBunga(GetNull(dbData!Plafond), GetNull(dbData!SukuBunga), GetNull(dbData!Lama), nKe, GetNull(dbData!Plafond) - GetNull(dbData!PelunasanPokok))
  nPokok = GetAngsuranPokok(GetNull(dbData!Plafond), GetNull(dbData!SukuBunga), GetNull(dbData!Lama), nKe, GetNull(dbData!Plafond) - GetNull(dbData!PelunasanPokok))
End Sub

Private Function GetAngsBunga(ByVal nPlafond As Double, ByVal nSukuBunga As Double, Optional ByVal nLama As Double = 1, Optional ByVal nKe As Double = 0, Optional ByVal nAkhir As Double = 0, Optional ByVal lMod50 As Boolean = True, Optional ByVal dTgl As Date) As Double
Dim nRetval As Double
Dim n As Single
Dim nPersBunga As Double
Dim nBungaEfektif As Double
Dim nLamaHari As Double
Dim nBungaSliding As Double
  
  dTgl = IIf(GetNull(dTgl, 0) = 0, Date, dTgl)
  nBungaSliding = Round(nSukuBunga / 100, 2)
  nPersBunga = nSukuBunga
  nSukuBunga = Round(nPersBunga / 12 / 100, 2)
  
  nLamaHari = GetLamaHari(BOM(dTgl), EOM(dTgl)) + 1
  nLamaHari = 30
  nRetval = nAkhir * nBungaSliding * nLamaHari / 360
  If lMod50 Then
    nRetval = Mod50(nRetval)
  End If
  GetAngsBunga = Round(nRetval, 0)
End Function

Function GetAngsuranPokok(ByVal nPlafond As Double, ByVal nSukuBunga As Double, _
                  ByVal nLama As Single, ByVal nKe As Single, Optional ByVal nAkhir As Double = 0, _
                  Optional ByVal lMod50 As Boolean = False) As Double
Dim n As Integer
Dim nRetval As Double

  nRetval = Mod50(Round(nPlafond / nLama, 0))
  If nKe = nLama Then
    If nAkhir = 0 Then
      nRetval = nPlafond - (nRetval * (nLama - 1))
    Else
      nRetval = nPlafond - nAkhir
    End If
  End If
  GetAngsuranPokok = Round(nRetval, 0)
End Function

Private Sub GetData()
Dim nBakiDebet As Double
Dim n As Single

    i = 0
    cField = "d.*,r.Nama,r.Alamat"
    vaJoin = Array("Left Join RegisterNasabah r on d.Kode = r.Kode")
    Set dbData = objData.Browse(GetDSN, "Debitur d", cField, "d.golonganKredit", sisGTEqual, cGolongan.Text, , "d.Rekening", vaJoin)
    vaarray.ReDim 0, -1, 0, 8
    If Not dbData.eof Then
      n = 0
      Do While Not dbData.eof
        GetDataAngsuran
        nBakiDebet = GetBakiDebet(objData, dbData!Rekening, dbData!Plafond, dDate.Value)
        If nBakiDebet > 0 Then
          vaarray.InsertRows vaarray.UpperBound(1) + 1
          i = vaarray.UpperBound(1)
'          n = 0
          vaarray(i, 0) = (dbData!KelompokDebitur)
          vaarray(i, 1) = (dbData!NamaKelompok)
          vaarray(i, 2) = (dbData!Rekening)
          vaarray(i, 3) = (dbData!Nama)
          vaarray(i, 4) = nBakiDebet
          vaarray(i, 5) = nPokok
          vaarray(i, 6) = nBunga
          vaarray(i, 7) = vaarray(i, 5) + vaarray(i, 6) '+ vaArray(i, 8)
          vaarray(i, 8) = (nKe)
        End If
        dbData.MoveNext
      Loop
      rpt
    End If
End Sub

Private Function TabN(n As Single) As Single
  TabN = n
  If n = vaarray.UpperBound(2) Then
    TabN = 0
  End If
    n = n + 1
End Function

Private Sub rpt()

  With FrmRPT
    .AddPageHeader "Tagihan Kredit", tdbHalignCenter, , , , , 12, True
    .AddPageHeader "Bulan " & GetMonth(Month(dDate.Value)) & " " & Year(dDate.Value), tdbHalignCenter, , , True, , 9, True
    .AddPageHeader "", , , , True
    
    .AddTableGroupHeader True, "[]", , , , 7
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
'    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
'    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
'    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
'    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Rekening", , tdbHalignCenter, , 12
    .AddTableHeader "Nama"
'    .AddTableHeader "Bagian / Instansi", , , , 10
    .AddTableHeader "Baki Debet", , , , 12
    .AddTableHeader "Pokok", , , , 11
    .AddTableHeader "Bunga", , , , 11
'    .AddTableHeader "Tabungan", , , , 8
    .AddTableHeader "Total", , , , 12
'    .AddTableHeader "Nomor Tabungan", , , , 10
'    .AddTableHeader "Komisi", , , , 8
    .AddTableHeader "Ke", , , , 3
        
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
'    .AddTableBody
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
'    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
'    .AddTableBody
'    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "Total", , tdbHalignRight, , , , , , , , , , , , 2
    .AddTableGroupFooter
'    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
'    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
'    .AddTableGroupFooter
'    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    .AddTableGroupFooter
    
    .Preview vaarray, True, True
  End With
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

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  TabIndex dDate, n
  TabIndex cGolongan, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub


