VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BISA BUTTON.OCX"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BISA DATE.OCX"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BISA FRAME.OCX"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BISA NUMBERBOX.OCX"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BISA TEXTBOX.OCX"
Begin VB.Form trOpenDepositoTeller 
   BorderStyle     =   0  'None
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4170
      Left            =   30
      Top             =   15
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7355
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningJurnal 
         Height          =   330
         Left            =   105
         TabIndex        =   14
         Top             =   2985
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   582
         Text            =   "12345678901234567890"
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
         MaxLength       =   20
         Button          =   -1  'True
         Caption         =   "KAS/BANK"
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
      Begin BiSADateProject.BiSADate dTempo 
         Height          =   330
         Left            =   105
         TabIndex        =   13
         Top             =   2610
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
         Left            =   105
         TabIndex        =   4
         Top             =   810
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
      Begin BiSATextBoxProject.BiSATextBox cKetGolDeposan 
         Height          =   330
         Left            =   2235
         TabIndex        =   0
         Top             =   105
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSATextBox cGolonganDeposan 
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   105
         Width           =   2100
         _ExtentX        =   3704
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
         Caption         =   "GOL DEPOSAN"
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
      Begin BiSATextBoxProject.BiSATextBox cKetGolDeposito 
         Height          =   330
         Left            =   2235
         TabIndex        =   2
         Top             =   450
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSATextBox cGolonganDeposito 
         Height          =   330
         Left            =   105
         TabIndex        =   3
         Top             =   450
         Width           =   2100
         _ExtentX        =   3704
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
         Caption         =   "GOL DEPOSITO"
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   480
         Left            =   2010
         Top             =   1200
         Width           =   3645
         _ExtentX        =   6429
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
         Begin VB.OptionButton optSetoran 
            Caption         =   "&1. Tunai"
            Height          =   330
            Index           =   0
            Left            =   105
            TabIndex        =   7
            Top             =   75
            Width           =   1050
         End
         Begin VB.OptionButton optSetoran 
            Caption         =   "&2. Tabungan"
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   6
            Top             =   75
            Width           =   1320
         End
         Begin VB.OptionButton optSetoran 
            Caption         =   "&3. PB"
            Height          =   330
            Index           =   2
            Left            =   2550
            TabIndex        =   5
            Top             =   75
            Width           =   1050
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame9 
         Height          =   480
         Left            =   2010
         Top             =   1695
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
            TabIndex        =   10
            Top             =   75
            Width           =   1020
         End
         Begin VB.OptionButton optARO 
            Caption         =   "&Ya"
            Height          =   330
            Index           =   0
            Left            =   105
            TabIndex        =   9
            Top             =   75
            Width           =   1050
         End
      End
      Begin BiSANumberBoxProject.BiSANumberBox nBunga 
         Height          =   330
         Left            =   105
         TabIndex        =   12
         Top             =   2220
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekening 
         Height          =   330
         Left            =   4290
         TabIndex        =   15
         Top             =   2985
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSANumberBoxProject.BiSANumberBox nNominal 
         Height          =   330
         Left            =   105
         TabIndex        =   16
         Top             =   3360
         Width           =   3345
         _ExtentX        =   5900
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame6 
         Height          =   1965
         Left            =   5895
         Top             =   885
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   3466
         Caption         =   "TABUNGAN"
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
         Begin BiSATextBoxProject.BiSATextBox cCabang 
            Height          =   330
            Left            =   180
            TabIndex        =   17
            Top             =   300
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   582
            Text            =   "1234"
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
            Left            =   2175
            TabIndex        =   18
            Top             =   300
            Width           =   375
            _ExtentX        =   661
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
         Begin BiSATextBoxProject.BiSATextBox cUrut 
            Height          =   330
            Left            =   2565
            TabIndex        =   19
            Top             =   300
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
            Text            =   "123456"
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
            Left            =   3360
            TabIndex        =   20
            Top             =   300
            Width           =   390
            _ExtentX        =   688
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
         Begin BiSATextBoxProject.BiSABrowse cNamaNasabah 
            Height          =   330
            Left            =   180
            TabIndex        =   21
            Top             =   690
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
            BackColor       =   -2147483633
            Enabled         =   0   'False
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
         Begin BiSATextBoxProject.BiSABrowse cAlamatNasabah 
            Height          =   330
            Left            =   180
            TabIndex        =   22
            Top             =   1065
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
            BackColor       =   -2147483633
            Enabled         =   0   'False
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
         Begin BiSANumberBoxProject.BiSANumberBox nSaldoAkhirTab 
            Height          =   330
            Left            =   180
            TabIndex        =   23
            Top             =   1440
            Width           =   3585
            _ExtentX        =   6324
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
            Caption         =   "SALDO AKHIR"
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
      End
      Begin VB.Label Label4 
         Caption         =   "SISTEM ARO"
         Height          =   360
         Left            =   210
         TabIndex        =   11
         Top             =   1845
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "SISTEM PENYETORAN"
         Height          =   360
         Left            =   165
         TabIndex        =   8
         Top             =   1350
         Width           =   1875
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame5 
      Height          =   600
      Left            =   30
      Top             =   4200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1058
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
         Height          =   405
         Left            =   10185
         TabIndex        =   24
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
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
         Picture         =   "trOpenDepositoTeller.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   405
         Left            =   9000
         TabIndex        =   25
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "     &Save"
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
         Picture         =   "trOpenDepositoTeller.frx":00A6
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
Dim objdata As New bisaMyDLL.Data
Dim lClick As Boolean
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
    
      If optARO(0).Value = True Then
        cStatusCair = "1"
      Else
        cStatusCair = "0"
      End If
    
      cRekeningTabungan = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
      UpdPembukaanDeposito objdata, cCabang.Text, cRekening, cNoFaktur, nNominal.Value, dTanggal, cRekeningJurnal.Text, GetOpt(optSetoran), cGolonganDeposito.Text, cRekeningTabungan, cStatusCair, cRekeningJurnal.Text
      UpdUrutFaktur objdata, cNoFaktur
            
      'simpn di MutasiDeposito
      cKodePembukaan = trPembukaan
      vaField = Array("Faktur", "Tgl", "KodeMutasi", "Rekening", "Jumlah", "UserName", "DateTime")
      vaValue = Array(trTeller.cFaktur.Text, trTeller.dTgl.Value, cKodePembukaan, cRekening, nNominal.Value, cUserName, Now)
      objdata.Add GetDSN, "MutasiDeposito", vaField, vaValue
      
      'Jika Setoran Tunai Bisa langsung di Print
      If optSetoran(0).Value = True Then
        If MsgBox("Akan mencetak Bilyet Deposito ?", vbYesNo) = vbYes Then
          With trCetakBilyetDeposito
            .cGolongan.Text = trTeller.cGolongan.Text
            .cUrut.Text = trTeller.cUrut.Text
            .cFrekuensi.Text = trTeller.cFrekuensi.Text
            
            .Show '1
          End With
        End If
      End If
    End If
    Initvalue
    cmdKeluar_Click
End Sub

Private Sub GetData()
Dim vaJoin
Dim cFields As String
  
  cFields = " d.*,r. Nama, r.Alamat, r.telepon,"
  cFields = cFields & " a.Keterangan as KetGolonganDeposan,"
  cFields = cFields & " b.Keterangan as KetGolonganDeposito,b.Lama,"
  cFields = cFields & " t.Rekening,t.Awal,t.Akhir,e.Nama as NamaNasabah,e.Alamat as AlamatNasabah"
  vaJoin = Array(" Left Join RegisterNasabah r on r.Kode=d.Kode", _
                " Left Join GolonganDeposan a on a.Kode=d.GolonganDeposan", _
                " Left Join GolonganDeposito b on b.Kode=d.GolonganDeposito", _
                " Left Join Tabungan t on t.Rekening=d.RekeningTabungan", _
                " Left Join RegisterNasabah e on e.Kode=t.Kode")
  Set dbData = objdata.Browse(GetDSN, "Deposito d", cFields, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.EOF Then
    With dbData
      cGolonganDeposan.Text = !GolonganDeposan
      cKetGolDeposan.Text = !KetGolonganDeposan
      cGolonganDeposito.Text = !GolonganDeposito
      cKetGolDeposito.Text = !KetGolonganDeposito
      cJangkaWaktu.Text = !Lama
      dTempo.Value = !JthTmp
      cCabang.Text = Left(!RekeningTabungan, 2)
      cGolongan.Text = Mid(!RekeningTabungan, 4, 2)
      cUrut.Text = Mid(!RekeningTabungan, 7, 6)
      cFrekuensi.Text = Right(!RekeningTabungan, 2)
      cNamaNasabah.Text = !NamaNasabah
      cAlamatNasabah.Text = !AlamatNasabah
      nSaldoAkhirTab.Value = !Awal + !Akhir
      nBunga.Value = !SukuBunga
      
      SetOpt optSetoran, !AsalSetoran
      SetOpt optARO, !SistemARO
      
      'edited by adjie 13-10-03
'      If optSetoran(2).Value = True Then
'        cRekeningJurnal.Enabled = True
'      End If
'
      cRekeningJurnal.Enabled = False
     
      If optSetoran(2).Value = True Then
'        cRekeningJurnal.Enabled = True
        cRekeningJurnal.Text = aCfg(msKodePemindahBukuan)
      End If
        
      If optSetoran(0).Value = True Then
'        cRekeningJurnal.Enabled = True
        cRekeningJurnal.Text = aCfg(msKodeKas)
      End If
    End With
  End If
End Sub
  
Private Sub Initvalue()
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNamaNasabah.Default
  cAlamatNasabah.Default
  cRekeningJurnal.Default
  cNamaRekening.Default
  nNominal.Value = 0
  
  cGolonganDeposan.Default
  cKetGolDeposan.Default
  cGolonganDeposito.Default
  cKetGolDeposito.Default
  cJangkaWaktu.Default
  dTempo.Value = Date
  
End Sub

Private Sub cRekeningJurnal_ButtonClick()
  Set dbData = objdata.Pick(GetDSN, "Rekening", "Kode", cRekeningJurnal, "Kode,Keterangan", " and Jenis = 'D'")
  If Not dbData.EOF Then
    cNamaRekening.Text = dbData!Keterangan
  End If
End Sub

Private Sub cRekeningJurnal_Validate(Cancel As Boolean)
  If cRekeningJurnal.LastKey = 13 Or Trim(cRekeningJurnal.Text) <> "" Then
    cRekeningJurnal_ButtonClick
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload trOpenDepositoTeller
  Me.Hide
  With trTeller
    .Image1.Picture = LoadPicture(GetPicture(""))
    .Image2.Picture = LoadPicture(GetPicture(""))
    trTeller.Height = 2745
    .cShow.Text = "0"
    .cGolongan.Default
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
  Me.Left = 0
  Me.Width = 11623
  cRekening = SetNomorRekening(trTeller.cCabang.Text, trTeller.cGolongan.Text, trTeller.cUrut.Text, trTeller.cFrekuensi.Text)
  dTanggal = trTeller.dTgl.Value
  cNoFaktur = trTeller.cFaktur.Text
  
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single

  Initvalue
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex cRekeningJurnal, n
  TabIndex nNominal, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub nNominal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    If nNominal.Value = 0 Or nNominal.Value < 0 Then
      MsgBox "Inputan tidak boleh 0 atau lebih kecil 0. Silahkan mengulangi pengisian !", vbOKOnly
      nNominal.SetFocus
      Exit Sub
    End If
  End If
End Sub


