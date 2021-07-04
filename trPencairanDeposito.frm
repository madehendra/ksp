VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPencairanDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PENCAIRAN DEPOSITO"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9285
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4185
      Left            =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7382
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
         Left            =   6675
         TabIndex        =   22
         Top             =   75
         Width           =   2460
         _ExtentX        =   4339
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame3 
         Height          =   2445
         Left            =   3735
         Top             =   1680
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   4313
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
         Begin VB.OptionButton OptCair 
            Caption         =   "Pokok"
            Height          =   240
            Index           =   1
            Left            =   3375
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   105
            Width           =   1245
         End
         Begin VB.OptionButton OptCair 
            Caption         =   "Bunga"
            Height          =   240
            Index           =   0
            Left            =   2280
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   105
            Width           =   1245
         End
         Begin BiSANumberBoxProject.BiSANumberBox nFinalti 
            Height          =   330
            Left            =   690
            TabIndex        =   0
            Top             =   405
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "PINALTY"
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
         Begin BiSANumberBoxProject.BiSANumberBox nPokok 
            Height          =   330
            Left            =   690
            TabIndex        =   1
            Top             =   765
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "POKOK"
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
         Begin BiSANumberBoxProject.BiSANumberBox nBahas 
            Height          =   330
            Left            =   690
            TabIndex        =   2
            Top             =   1125
            Width           =   3795
            _ExtentX        =   6694
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
         Begin BiSANumberBoxProject.BiSANumberBox nTotal 
            Height          =   330
            Left            =   690
            TabIndex        =   23
            Top             =   1965
            Width           =   3795
            _ExtentX        =   6694
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
            BackColor       =   -2147483647
            ForeColor       =   -2147483634
            Caption         =   "TOTAL CAIR"
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
         Begin BiSANumberBoxProject.BiSANumberBox nPajak 
            Height          =   330
            Left            =   690
            TabIndex        =   24
            Top             =   1485
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "PAJAK BUNGA"
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
         Begin VB.Line Line1 
            X1              =   750
            X2              =   4680
            Y1              =   1875
            Y2              =   1875
         End
         Begin VB.Label Label4 
            Caption         =   "PENCAIRAN"
            Height          =   240
            Left            =   750
            TabIndex        =   21
            Top             =   120
            Width           =   1305
         End
      End
      Begin BiSAFramProject.BiSAFrame sisPesan 
         Height          =   525
         Left            =   3735
         Top             =   1155
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   926
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Anda Terkena Pinalty Karena Pencairan Pokok Belum Jatuh Tempo"
            Height          =   300
            Left            =   240
            TabIndex        =   3
            Top             =   135
            Width           =   5010
         End
      End
      Begin BiSADateProject.BiSADate dTempo 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   2115
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
      Begin BiSANumberBoxProject.BiSANumberBox nLama 
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   1695
         Width           =   2325
         _ExtentX        =   4101
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
      Begin BiSANumberBoxProject.BiSANumberBox nBunga 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   2550
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "BUNGA (%) p.a"
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
      Begin BiSANumberBoxProject.BiSANumberBox nNominalDeposito 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   2970
         Width           =   3345
         _ExtentX        =   5900
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
      Begin BiSADateProject.BiSADate dValuta 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   1245
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
         Caption         =   "TGL VALUTA"
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   450
         Width           =   5160
         _ExtentX        =   9102
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
         Caption         =   "NAMA DEPOSAN"
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   795
         Width           =   5970
         _ExtentX        =   10530
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
         Caption         =   "ALAMAT"
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
      Begin BiSATextBoxProject.BiSATextBox cCabang 
         Height          =   330
         Left            =   120
         TabIndex        =   14
         Top             =   90
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         Text            =   "1234"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         MaxLength       =   4
         Caption         =   "NO. REKENING"
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   2160
         TabIndex        =   15
         Top             =   90
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         Button          =   -1  'True
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
      Begin BiSATextBoxProject.BiSATextBox cUrut 
         Height          =   330
         Left            =   2985
         TabIndex        =   16
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Text            =   "123456"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         MaxLength       =   6
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   3885
         TabIndex        =   17
         Top             =   90
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
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
         Caption         =   "Bulan"
         Height          =   180
         Left            =   2535
         TabIndex        =   18
         Top             =   1725
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   240
         Left            =   5490
         TabIndex        =   9
         Top             =   870
         Width           =   210
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   630
      Left            =   0
      Top             =   4170
      Width           =   9255
      _ExtentX        =   16325
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
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   6960
         TabIndex        =   10
         Top             =   105
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
         Picture         =   "trPencairanDeposito.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   8040
         TabIndex        =   11
         Top             =   105
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
         Picture         =   "trPencairanDeposito.frx":0416
      End
   End
End
Attribute VB_Name = "trPencairanDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New BisaMyDLL.data
Dim cRekening As String
Dim cRekeningJT As String
Dim cRekeningKAS As String
Dim cRekeningFinalty As String
Dim cRekeningTitipanBunga As String
Dim cRekneningPajakBunga As String
Dim cFaktur As String

Private Sub Initvalue()
  dTgl.Value = Date
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
  dValuta.Value = Date
  nLama.Value = 0
  dTempo.Value = Date
  nBunga.Value = 0
  nNominalDeposito.Value = 0
  nFinalti.Value = 0
  nPokok.Value = 0
  nBahas.Value = 0
  nPajak.Value = 0
  nTotal.Value = 0
  OptCair(0).Value = True
  cRekeningKAS = GetKasTeller(cusername)
End Sub

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.nama,r.Alamat,d.Rekening", "r.Alamat", sisContent, cAlamat.Text, "And d.Status <> '1'", "r.Nama", _
                              Array("Left Join RegisterNasabah r on r.Kode=d.Kode"))
  cAlamat.Text = cAlamat.Browse(dbData)
  If Not dbData.eof Then
    cCabang.Text = left(GetNull(dbData!Rekening, ""), 2)
    cGolongan.Text = Mid(GetNull(dbData!Rekening, ""), 4, 2)
    cUrut.Text = Mid(GetNull(dbData!Rekening, ""), 7, 6)
    cFrekuensi.Text = Right(GetNull(dbData!Rekening, ""), 2)
    GetData
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganDeposito", "Kode", cGolongan, "Kode,Keterangan,RekeningAkuntansi,RekeningBunga,RekeningPajakBunga,rekeningjatuhTempo,CadanganBunga,Rekeningfinalty")
  If Not dbData.eof Then
    cRekeningJT = GetNull(dbData!RekeningJatuhtempo, "")
    cRekeningFinalty = GetNull(dbData!RekeningFinalty, "")
    cRekeningTitipanBunga = GetNull(dbData!Cadanganbunga, "")
    cRekneningPajakBunga = GetNull(dbData!RekeningPajakbunga, "")
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

Private Sub GetData()
Dim cFields As String
Dim vaJoin
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cFields = "d.Rekening,d.Nominaldeposito,d.Tgl,d.jthtmp,d.Status,d.StatusBlokir,d.Sukubunga,"
  cFields = cFields & " r.Nama,r.Alamat,b.Lama"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode = d.Kode", _
                 "Left Join GolonganDeposito b on b.Kode=d.GolonganDeposito")
  Set dbData = objData.Browse(GetDSN, "Deposito d", cFields, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    cNama.Text = GetNull(dbData!Nama, "")
    cAlamat.Text = GetNull(dbData!Alamat, "")
    dValuta.Value = GetNull(dbData!Tgl, "")
    dTempo.Value = GetNull(dbData!JthTmp, "")
    nLama.Value = GetNull(dbData!Lama)
    nBunga.Value = GetNull(dbData!SukuBunga)
    nNominalDeposito.Value = GetNull(dbData!NominalDeposito, "")
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
    
  If Not CheckData(cCabang.Text, "Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cCabang.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cGolongan.Text, "Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cUrut.Text, "Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cUrut.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cFrekuensi.SetFocus
    Exit Function
  End If
End Function

Private Sub cmdSimpan_Click()

  'JURNAL DEPOSITO
  '1. PEMBUKAAN
  '   KAS
  '       DEPOSITO
  
  '2. PENCAIRAN POKOK
  '   Pada saat Posting Akhir hari
  '   Deposito
  '       Deposito Jatuh Tempo
  
  '   Pada saat pencairan pokok
  '   Deposito Jatuh Tempo
  '       KAS (Pokok - Finalty)
  '       Finalty
  
  '.3. PENCAIRAN BUNGA
  '   Pada saat posting awal hari
  '   Biaya Bunga Deposito
  '       Titipan Bunga Deposito
  
  '   Pada saat Pencairan bunga
  '   Titipan Bunga Depsoito
  '       KAS (Bunga - Pajak)
  '       Pajak bunga Deposito
  
  If OptCair(0).Value = True Then
    If nBahas.Value <= 0 Then
      MsgBox "Bunga kosong...", vbInformation
      cCabang.SetFocus
      Exit Sub
    End If
  Else
    If nPokok.Value <= 0 Then
      MsgBox "Pokok tidak ada...", vbInformation
      cCabang.SetFocus
      Exit Sub
    End If
  End If
  
  If ValidSaving() Then
      If MsgBox("Apakah Rekening Benar-benar akan disimpan?", vbYesNo) = vbYes Then
        cFaktur = GetLastFaktur(fkt_Deposito, dTgl.Value, True)
        If OptCair(0).Value = True Then
          'Update di MutasiDeposito
          'Pajak + bunga
          GetSimpanMutasi cFaktur, dTgl.Value, cRekening, trPencairanPokok, nBahas.Value + nPajak.Value, cusername, Now
          
          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur, dTgl.Value, cRekeningTitipanBunga, "Pencairan Bunga Deposito a.n " & cNama.Text, nBahas.Value, 0, , Now
            UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur, dTgl.Value, cRekeningKAS, "Pencairan Bunga Deposito a.n " & cNama.Text, 0, nBahas.Value - nPajak.Value, , Now
            UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur, dTgl.Value, cRekneningPajakBunga, "Pajak Bunga Deposito a.n " & cNama.Text, 0, nPajak.Value, , Now
            
          'Hapus di BungaDeposito
          objData.Delete GetDSN, "BungaDeposito", "Rekening", sisAssign, cRekening
        Else
          'Update di MutasiDeposito
          'Pokok dan finalty
          GetSimpanMutasi cFaktur, dTgl.Value, cRekening, trPencairanPokok, nPokok.Value, cusername, Now
          GetSimpanMutasi cFaktur, dTgl.Value, cRekening, trPencairanPokok, nFinalti.Value, cusername, Now
          
          'Edit StatusCair dan tglCair
          objData.Edit GetDSN, "deposito", "rekening='" & cRekening & "'", Array("Statuscair", "Tglcair"), Array("1", dTgl.Value)
          
          'buku besar
          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur, dTgl.Value, cRekeningJT, "Pencairan Pokok Deposito a.n " & cNama.Text, nNominalDeposito.Value, 0, "N", Now
            UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur, dTgl.Value, cRekeningKAS, "Pencairan Pokok Deposito a.n " & cNama.Text, , nPokok.Value, "K", Now
            UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur, dTgl.Value, cRekeningFinalty, "Finalty Pencairan Pokok Deposito a.n " & cNama.Text, nFinalti.Value, 0, "N", Now
        End If
      End If
      Initvalue
      cGolongan.SetFocus
  End If
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cRekening As String
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  If cFrekuensi.LastKey = 13 Then
     Set dbData = objData.Browse(GetDSN, "Deposito", , "Rekening", sisAssign, cRekening)
     If Not dbData.eof Then
        If GetNull(dbData!status, "") = "1" Then
          MsgBox "Rekening tersebut sudah Tutup (Cair) !", vbOKOnly, "Blokir Tabungan"
          Initvalue
          cGolongan.SetFocus
          Exit Sub
        End If
        GetData
      Else
        MsgBox "Rekening dengan Nomor : " & cRekening & " Tidak ada. Silahkan Ulangi pengisian.", vbOKOnly, "Blokir Rekening Deposito"
        Cancel = True
        cFrekuensi.Default
        cFrekuensi.SetFocus
        Exit Sub
      End If
  End If
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.nama,r.Alamat,d.Rekening", "r.Nama", sisContent, cNama.Text, "And d.Status <>'1'", "r.Nama", _
                              Array("Left Join RegisterNasabah r on r.Kode=d.Kode"))
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.eof Then
    cCabang.Text = left(GetNull(dbData!Rekening, ""), 2)
    cGolongan.Text = Mid(GetNull(dbData!Rekening, ""), 4, 2)
    cUrut.Text = Mid(GetNull(dbData!Rekening, ""), 7, 6)
    cFrekuensi.Text = Right(GetNull(dbData!Rekening, ""), 2)
    GetData
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  Initvalue
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex OptCair(0), n
  TabIndex OptCair(1), n
  TabIndex nFinalti, n
  TabIndex nPokok, n
  TabIndex nBahas, n
  TabIndex nPajak, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub nBahas_Validate(Cancel As Boolean)
  nTotal.Value = nBunga.Value - nPajak.Value
End Sub

Private Sub nFinalti_Validate(Cancel As Boolean)
  nPokok.Value = IIf(dTgl.Value < dTempo.Value, 0, nNominalDeposito.Value - nFinalti.Value)
  nTotal.Value = nPokok.Value
End Sub

Private Sub nPajak_Validate(Cancel As Boolean)
  nTotal.Value = nBunga.Value - nPajak.Value
End Sub

Private Sub OptCair_Click(Index As Integer)
  If Index = 0 Then
    GetBunga cRekening
  Else
    nPokok.Value = nNominalDeposito.Value
  End If
End Sub

Private Sub OptCair_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub GetBunga(ByVal cRekening As String)
  Set dbData = objData.Browse(GetDSN, "BungaDeposito", "Sum(Bunga) as Bunga,Sum(Pajak) as Pajak", "Rekening", sisAssign, cRekening)
  If Not dbData.eof Then
    nBahas.Value = GetNull(dbData!Bunga)
    nPajak.Value = GetNull(dbData!pajak)
    nTotal.Value = nBahas.Value - nPajak.Value
  End If
End Sub

Private Sub GetSimpanMutasi(ByVal cNomorFaktur As String, ByVal dTgl As Date, ByVal cRekening As String, ByVal cKodePencairan As trDeposito, ByVal nJumlah As Double, ByVal cusername As String, ByVal dDateTime As Date)
Dim vaField
Dim vaValue
  
  If nJumlah <> 0 Then
    vaField = Array("Faktur", "Tgl", "KodeMutasi", "Rekening", "Jumlah", "UserName", "DateTime")
    vaValue = Array(cNomorFaktur, dTgl, cKodePencairan, cRekening, nJumlah, cusername, dDateTime)
    objData.Add GetDSN, "MutasiDeposito", vaField, vaValue
  End If
End Sub

Private Function GetKasTeller(ByVal cusername As String) As String
  GetKasTeller = ""
  Set dbData = objData.Browse(GetDSN, "Username", "kasTeller", "userName", sisAssign, cusername)
  If Not dbData.eof Then
    GetKasTeller = GetNull(dbData!KasTeller, "")
  End If
End Function

