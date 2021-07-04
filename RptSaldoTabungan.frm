VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptSaldoTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN SALDO SIMPANAN"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7725
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2940
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   5186
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   465
         Left            =   2085
         Top             =   2370
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   820
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
         Begin VB.OptionButton optJenisKelamin 
            Caption         =   "Semuanya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2835
            TabIndex        =   17
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton optJenisKelamin 
            Caption         =   "Perempuan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1350
            TabIndex        =   16
            Top             =   120
            Width           =   1275
         End
         Begin VB.OptionButton optJenisKelamin 
            Caption         =   "Laki"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   285
            TabIndex        =   15
            Top             =   120
            Width           =   720
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   435
         Left            =   2100
         Top             =   1935
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   767
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
            Caption         =   "Semuanya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2790
            TabIndex        =   14
            Top             =   105
            Width           =   1065
         End
         Begin VB.OptionButton optAnggota 
            Caption         =   "Calon Anggota"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1350
            TabIndex        =   13
            Top             =   105
            Width           =   1395
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
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   12
            Top             =   105
            Width           =   975
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame3 
         Height          =   435
         Left            =   2115
         Top             =   1500
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   767
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
         Begin VB.OptionButton optTampil 
            Caption         =   "Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   150
            Width           =   720
         End
         Begin VB.OptionButton optTampil 
            Caption         =   "Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            Top             =   135
            Width           =   720
         End
      End
      Begin BiSANumberBoxProject.BiSANumberBox nSaldo 
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   1155
         Width           =   3735
         _ExtentX        =   6588
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
         Caption         =   "ANTARA SALDO"
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
         Left            =   2880
         TabIndex        =   1
         Top             =   465
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
         Left            =   135
         TabIndex        =   2
         Top             =   465
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
         Left            =   135
         TabIndex        =   3
         Top             =   120
         Width           =   3300
         _ExtentX        =   5821
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldo 
         Height          =   330
         Index           =   1
         Left            =   4170
         TabIndex        =   4
         Top             =   1170
         Width           =   2340
         _ExtentX        =   4128
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
      Begin BiSATextBoxProject.BiSABrowse cPdl 
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   810
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
         Caption         =   "ANTARA PDL"
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
      Begin BiSATextBoxProject.BiSABrowse cPdl 
         Height          =   330
         Index           =   1
         Left            =   3225
         TabIndex        =   8
         Top             =   810
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
      Begin VB.Label Label1 
         Caption         =   "TAMPILKAN SALDO 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   9
         Top             =   1560
         Width           =   1995
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   2910
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
         TabIndex        =   5
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
         Picture         =   "RptSaldoTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5265
         TabIndex        =   6
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
         Picture         =   "RptSaldoTabungan.frx":00A6
      End
   End
End
Attribute VB_Name = "RptSaldoTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim lEmpty As Boolean

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganTabungan", "Kode", cGolongan, "Kode,Keterangan")
  If Not dbData.eof Then
    cNamaGolongan.Text = dbData!Keterangan
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

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub GetSQL()
Dim cWhere As String
Dim n As Double
Dim dTanggalTutup
Dim cField As String
  
  lEmpty = False
  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 5
  
  cWhere = cWhere & " t.GolonganTabungan = '" & cGolongan.Text & "' "
'  cWhere = cWhere & " and t.Tgl <= '" & Format(dDate.Value, "yyyy-mm-dd") & "'"
'  cWhere = cWhere & " and ucase(t.PDL) >= '" & UCase(cPDL(0).Text) & "'"
'  cWhere = cWhere & " and ucase(t.PDL) <= '" & UCase(cPDL(1).Text) & "'"

'
'  If optAnggota(0).Value = True Then
'    cWhere = cWhere & " and r.jenisanggota = '1'"
'  ElseIf optAnggota(1).Value = True Then
'    cWhere = cWhere & " and r.jenisanggota = '2'"
'  End If
'
'  If optJenisKelamin(0).Value = True Then 'laki
'    cWhere = cWhere & " and r.kelamin = 'L'"
'  ElseIf optJenisKelamin(1).Value = True Then
'    cWhere = cWhere & " and r.kelamin = 'P'"
'  End If


'  cWhere = cWhere & " and t.Close <> '1' "
  
'  cField = "t.PDL,a.Keterangan as NamaPDL,t.Rekening,r.Nama,r.Alamat,t.Awal,t.awaltahun,t.Tgl,t.TglPenutupan,t.Close"
'  Set dbData = objData.Browse(GetDSN, "Tabungan t", cField, , , , cWhere, "t.Rekening", _
'                              Array("Left Join RegisterNasabah r on t.Kode = r.Kode", _
'                                    "Left Join pdl a on a.Kode = t.PDL"))

  cField = "t.PDL,a.Keterangan as NamaPDL,t.Rekening,r.Nama,r.Alamat,t.Awal,t.awaltahun,t.Tgl,t.TglPenutupan,t.Close"
  Set dbData = objData.Browse(GetDSN, "Tabungan t", cField, , , , cWhere, "t.PDL,t.GolonganTabungan,t.Rekening", _
                              Array("Left Join RegisterNasabah r on t.Kode = r.Kode", _
                                    "Left Join pdl a on a.Kode = t.PDL"))

  If Not dbData.eof Then
     FrmPB.InitPB dbData.RecordCount
     Do While Not dbData.eof
       FrmPB.RunPB
       dTanggalTutup = IIf(GetNull(dbData!TglPenutupan, 0) = 0, DateAdd("d", 1, dbData!Tgl), dbData!TglPenutupan)
       
'       If dTanggalTutup <= dDate.Value And dbData!Close = "1" Then
''        MsgBox ""
'       Else
'          vaArray.InsertRows vaArray.UpperBound(1) + 1
'          n = vaArray.UpperBound(1)
'
'          vaArray(n, 0) = UCase(dbData!PDL)
'          vaArray(n, 1) = (dbData!namapdl)
'          vaArray(n, 2) = (dbData!Rekening)
'          vaArray(n, 3) = (dbData!nama)
'          vaArray(n, 4) = (dbData!Alamat)
'          vaArray(n, 5) = (dbData!AwalTahun)
'       End If

      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)

      vaArray(n, 0) = UCase(dbData!PDL)
      vaArray(n, 1) = (dbData!namapdl)
      vaArray(n, 2) = (dbData!Rekening)
      vaArray(n, 3) = (dbData!nama)
      vaArray(n, 4) = (dbData!alamat)
      vaArray(n, 5) = (dbData!AwalTahun)
              
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    
    cWhere = " t.Tgl <= '" & Format(dDate.Value, "yyyy-MM-dd") & "' Group by t.Rekening"
    Set dbData = objData.Browse(GetDSN, "MutasiTabungan t", "t.Rekening,k.DK,Sum(t.Jumlah) as Jumlah,t.KodeTransaksi", , , , cWhere, , _
                 Array("Left Join KodeTransaksi k on k.Kode = t.KodeTransaksi"))
    If Not dbData.eof Then
      dbData.MoveFirst
      FrmPB.InitPB dbData.RecordCount
      Do While Not dbData.eof
        FrmPB.RunPB
        n = vaArray.Find(0, 2, GetNull(dbData!Rekening))
        If n >= 0 Then
          vaArray(n, 5) = GetNull(vaArray(n, 5)) + IIf((dbData!DK) = "D", -(dbData!Jumlah), (dbData!Jumlah)) 'Round(GetNull(vaArray(n, 5)) + IIf((dbData!DK) = "D", -(dbData!Jumlah), (dbData!Jumlah)), 0)
        End If
        dbData.MoveNext
      Loop
      FrmPB.EndPB
    End If
      
'    FrmPB.InitPB vaArray.UpperBound(1)
'
'    For n = 0 To vaArray.UpperBound(1)
'      FrmPB.RunPB
'      If vaArray.UpperBound(1) >= n Then
'        If Not (vaArray(n, 5) >= nSaldo(0).Value And vaArray(n, 5) <= nSaldo(1).Value) Then
'          vaArray.DeleteRows n
'          n = n - 1
'        End If
'        If optTampil(1).Value = True Then
'          If n >= 0 Then
'            If vaArray(n, 5) = 0 Then
'              vaArray.DeleteRows n
'              n = n - 1
'            End If
'          End If
'        End If
'      End If
'    Next
'
'    FrmPB.EndPB
    rpt
  Else
    MsgBox "Data tidak ada..", vbInformation
    Exit Sub
  End If
End Sub

Private Sub cPdl_ButtonClick(Index As Integer)
  Set dbData = objData.Pick(GetDSN, "PDL", "Kode", cPDL(Index), "Kode,Keterangan")
End Sub

Private Sub cPdl_Validate(Index As Integer, Cancel As Boolean)
  If cPDL(Index).LastKey = 13 Or cPDL(Index).LastKey = 40 Then
    cPdl_ButtonClick (Index)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  optTampil(1).Value = True
  nSaldo(0).Value = 0
  nSaldo(1).Value = 9999999999#
  dDate.Value = Date
  GetMinMax "PDl", cPDL, "Kode"
      
  optAnggota(2).Value = True
  optJenisKelamin(2).Value = True
  TabIndex dDate, n
  TabIndex cGolongan, n
  TabIndex cPDL(0), n
  TabIndex cPDL(1), n
  TabIndex nSaldo(0), n
  TabIndex nSaldo(1), n
  TabIndex optTampil(0), n
  TabIndex optTampil(1), n
  TabIndex optAnggota(0), n
  TabIndex optAnggota(1), n
  TabIndex optAnggota(2), n
  TabIndex optJenisKelamin(0), n
  TabIndex optJenisKelamin(1), n
  TabIndex optJenisKelamin(2), n
  TabIndex cmdPreview, n
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader "DAFTAR SALDO TABUNGAN", tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader cNamaGolongan.Text, tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "Sampai dengan Tanggal : " & Format(dDate.Value, "dd MMMM yyyy"), tdbHalignCenter, , , True, , 9, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader "Antara Saldo", , True, 17, True
    .AddPageHeader ": " & Format(nSaldo(0).Value, "###,###,###,###,##0.00") & " s/d " & Format(nSaldo(1).Value, "###,###,###,###,##0.00"), , , , False '
    
    .AddTableGroupHeader True, "[]", , , , 10
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Rekening", , , , 13
    .AddTableHeader "Nama", , , , 18
    .AddTableHeader "Alamat"
    .AddTableHeader "Saldo Akhir", , , , 15
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "Sub Total", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&SUM", Sis_Rpt_Number2
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub

Private Sub optAnggota_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optJenisKelamin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optTampil_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub
