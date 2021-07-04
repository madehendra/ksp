VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trOpenDepositoTeller 
   BorderStyle     =   0  'None
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2325
      Left            =   30
      Top             =   15
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4101
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
      Begin BiSADateProject.BiSADate dTempo 
         Height          =   330
         Left            =   2595
         TabIndex        =   5
         Top             =   1440
         Width           =   2925
         _ExtentX        =   5159
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Caption         =   "JATUH TEMPO"
         CaptionWidth    =   1500
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
      Begin BiSANumberBoxProject.BiSANumberBox cJangkaWaktu 
         Height          =   330
         Left            =   2595
         TabIndex        =   0
         Top             =   195
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   582
         Decimals        =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "JANGKA WAKTU"
         CaptionWidth    =   1500
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame9 
         Height          =   480
         Left            =   4185
         Top             =   555
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   847
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
         Begin VB.OptionButton optARO 
            Caption         =   "&Tidak"
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   2
            Top             =   75
            Width           =   1020
         End
         Begin VB.OptionButton optARO 
            Caption         =   "&Ya"
            Height          =   330
            Index           =   0
            Left            =   105
            TabIndex        =   1
            Top             =   75
            Width           =   1050
         End
      End
      Begin BiSANumberBoxProject.BiSANumberBox nBunga 
         Height          =   330
         Left            =   2595
         TabIndex        =   4
         Top             =   1080
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   582
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "BUNGA"
         CaptionWidth    =   1500
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
      Begin BiSANumberBoxProject.BiSANumberBox nNominal 
         Height          =   330
         Left            =   2595
         TabIndex        =   6
         Top             =   1800
         Width           =   3990
         _ExtentX        =   7038
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
         Caption         =   "NOMINAL"
         CaptionWidth    =   1500
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
      Begin VB.Label Label4 
         Caption         =   "SISTEM ARO"
         Height          =   360
         Left            =   2640
         TabIndex        =   3
         Top             =   690
         Width           =   1365
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame5 
      Height          =   510
      Left            =   30
      Top             =   2325
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   900
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
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9180
         TabIndex        =   7
         Top             =   45
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         Caption         =   "    &Save"
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
         Picture         =   "trOpenDepositoTeller1.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10260
         TabIndex        =   8
         Top             =   45
         Width           =   1080
         _ExtentX        =   1905
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
         Picture         =   "trOpenDepositoTeller1.frx":0416
      End
   End
End
Attribute VB_Name = "trOpenDepositoTeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New BisaMyDLL.data
Dim lclick As Boolean
Dim nPos As Single
Dim cSql As String
Dim lEdit As Boolean
Dim cRekening As String
Dim cNoFaktur As String
Dim dTanggal As Date

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim cRekeningTabungan As String
Dim cKodePembukaan As trDeposito
Dim cStatusCair As String

  
    If MsgBox("Apakah Data Benar-benar sudah Valid ?", vbYesNo) = vbYes Then


        If MsgBox("Akan mencetak Bilyet Deposito ?", vbYesNo) = vbYes Then
'          With trCetakBilyetDeposito
'            .cGolongan.Text = trTeller.cGolongan.Text
'            .cUrut.Text = trTeller.cUrut.Text
'            .cFrekuensi.Text = trTeller.cFrekuensi.Text
'            .Show '1
'          End With
        End If
    End If
    Initvalue
    cmdKeluar_Click
End Sub

Private Sub GetData()
Dim vaJoin
Dim cFields As String
  
  cFields = " d.JthTmp,d.Sukubunga,d.SistemARO,b.Lama"
  vaJoin = Array("Left Join GolonganDeposito b on b.Kode = d.GolonganDeposito")
  Set dbData = objData.Browse(GetDSN, "Deposito d", cFields, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    cJangkaWaktu.Value = GetNull(dbData!Lama)
    dTempo.Value = GetNull(dbData!JthTmp)
    nBunga.Value = GetNull(dbData!SukuBunga)
    SetOpt optARO, GetNull(dbData!SistemARO)
  End If
End Sub
  
Private Sub Initvalue()
  nNominal.Value = 0
  cJangkaWaktu.Value = 0
  dTempo.Value = Date
  optARO(0).Value = True
End Sub

Private Sub cmdKeluar_Click()
  Unload trOpenDepositoTeller
  Me.Hide
  With trTeller
    .Image1.Picture = LoadPicture(GetPicture(""))
    .Image2.Picture = LoadPicture(GetPicture(""))
    trTeller.Height = 2745
    .cShow.Text = "0"
    .cGolongan.Text = ""
    .cUrut.Default
    .cFrekuensi.Default
    .cAlamat.Default
    .cNama.Default
    .cFaktur.Default
    .dTgl.Value = Date
    .cGolongan.SetFocus
  End With
  Exit Sub
End Sub

Private Sub Form_Activate()
  Me.Top = 2300
  Me.left = 0
  Me.Width = 11623
  cRekening = SetNomorRekening(trTeller.cCabang.Text, trTeller.cGolongan.Text, trTeller.cUrut.Text, trTeller.cFrekuensi.Text)
  dTanggal = trTeller.dTgl.Value
  cNoFaktur = trTeller.cFaktur.Text
  
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single

  TabIndex nNominal, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub nNominal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    If nNominal.Value <= 0 Then
      MsgBox "Nilai nominal tidak valid. Silahkan mengulangi pengisian !", vbOKOnly
      nNominal.SetFocus
      Exit Sub
    End If
  End If
End Sub
