VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form TrOpenTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMBUKAAN REKENING SIMPANAN"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11640
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5370
      Left            =   -15
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   9472
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame6 
         Height          =   2445
         Left            =   2175
         Top             =   2370
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   4313
         Caption         =   "SPECIMEN"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BackColor       =   -2147483633
         Begin VB.Image Image2 
            Height          =   2130
            Left            =   75
            Stretch         =   -1  'True
            Top             =   225
            Width           =   3840
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   2445
         Left            =   150
         Top             =   2370
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   4313
         Caption         =   "FOTO"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BackColor       =   -2147483633
         Begin VB.Image Image1 
            Height          =   2160
            Left            =   90
            Stretch         =   -1  'True
            Top             =   210
            Width           =   1860
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   915
         Left            =   6195
         Top             =   60
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   1614
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
            Height          =   390
            Left            =   4545
            TabIndex        =   0
            Top             =   255
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   688
            Text            =   "12"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Verdana"
            FontSize        =   11.25
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
            Height          =   390
            Left            =   2610
            TabIndex        =   1
            Top             =   255
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
            Text            =   "12"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Verdana"
            FontSize        =   11.25
            MaxLength       =   2
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
         Begin BiSATextBoxProject.BiSATextBox cCabang 
            Height          =   390
            Left            =   255
            TabIndex        =   2
            Top             =   255
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   688
            Text            =   "12"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Verdana"
            FontSize        =   11.25
            MaxLength       =   2
            Caption         =   "NO. REKENING"
            CaptionWidth    =   1700
            CaptionBackColor=   12632256
            CaptionForeColor=   -2147483635
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin BiSATextBoxProject.BiSATextBox cUrut 
            Height          =   390
            Left            =   3375
            TabIndex        =   3
            Top             =   255
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   688
            Text            =   "123456"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Verdana"
            FontSize        =   11.25
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
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   4245
         Left            =   6225
         Top             =   1050
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   7488
         Caption         =   "DATA PINJAMAN"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Begin BiSANumberBoxProject.BiSANumberBox nAdministrasi 
            Height          =   330
            Left            =   75
            TabIndex        =   23
            Top             =   2670
            Width           =   3930
            _ExtentX        =   6932
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
            Caption         =   "Biaya Administrasi"
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
            Left            =   75
            TabIndex        =   21
            Top             =   1425
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
            Caption         =   "PDL"
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
         Begin BiSANumberBoxProject.BiSANumberBox nSaldoMinimum 
            Height          =   330
            Left            =   75
            TabIndex        =   4
            Top             =   690
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   582
            Appearance      =   0
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
            BackColor       =   12632256
            Caption         =   "Saldo Minimum"
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
         Begin BiSATextBoxProject.BiSABrowse cGolonganTabungan 
            Height          =   330
            Left            =   75
            TabIndex        =   5
            Top             =   315
            Width           =   2370
            _ExtentX        =   4180
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
            BackColor       =   12632256
            Enabled         =   0   'False
            Appearance      =   0
            Caption         =   "Gol Pinjaman"
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
         Begin BiSATextBoxProject.BiSATextBox cKetGolTabungan 
            Height          =   330
            Left            =   2460
            TabIndex        =   6
            Top             =   315
            Width           =   2760
            _ExtentX        =   4868
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
            BackColor       =   12632256
            Enabled         =   0   'False
            Appearance      =   0
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
         Begin BiSANumberBoxProject.BiSANumberBox nSetoranMinimum 
            Height          =   330
            Left            =   75
            TabIndex        =   7
            Top             =   1065
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   582
            Appearance      =   0
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
            BackColor       =   12632256
            Caption         =   "Setoran Minimum"
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
         Begin BiSATextBoxProject.BiSATextBox cNamaPDL 
            Height          =   330
            Left            =   75
            TabIndex        =   22
            Top             =   1785
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
            BackColor       =   12632256
            Enabled         =   0   'False
            Appearance      =   0
            Caption         =   "Nama PDL"
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
         Begin BiSATextBoxProject.BiSABrowse cPOSKas 
            Height          =   330
            Left            =   75
            TabIndex        =   24
            Top             =   3450
            Width           =   3240
            _ExtentX        =   5715
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
            Caption         =   "Rek Kas"
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
         Begin BiSATextBoxProject.BiSABrowse cRekPendapatanAdministrasi 
            Height          =   330
            Left            =   915
            TabIndex        =   25
            Top             =   3840
            Width           =   3600
            _ExtentX        =   6350
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
            Caption         =   "Rek Pendapatan"
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
         Begin VB.Label Label1 
            Caption         =   "Jurnal Umum"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   26
            Top             =   3105
            Width           =   1560
         End
      End
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   150
         TabIndex        =   8
         Top             =   855
         Width           =   5430
         _ExtentX        =   9578
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
         Caption         =   "Nama Nasabah"
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
      Begin BiSATextBoxProject.BiSATextBox cKode 
         Height          =   330
         Left            =   2355
         TabIndex        =   9
         Top             =   495
         Width           =   990
         _ExtentX        =   1746
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   150
         TabIndex        =   10
         Top             =   135
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
         Caption         =   "Tgl Pembukaan"
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
      Begin BiSATextBoxProject.BiSATextBox cCabang1 
         Height          =   330
         Left            =   150
         TabIndex        =   11
         Top             =   495
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   582
         Text            =   "AA"
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
         MaxLength       =   4
         Caption         =   "No Register"
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
      Begin BiSATextBoxProject.BiSATextBox cTelepon 
         Height          =   330
         Left            =   150
         TabIndex        =   12
         Top             =   1575
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   582
         Text            =   "1234567890123456789012345678901234567890"
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
         BackColor       =   12632256
         MaxLength       =   40
         Appearance      =   0
         Caption         =   "Telepon"
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   150
         TabIndex        =   13
         Top             =   1215
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
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Alamat Nasabah"
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
      Begin BiSATextBoxProject.BiSATextBox cPekerjaan 
         Height          =   330
         Left            =   150
         TabIndex        =   14
         Top             =   1935
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   582
         Text            =   "1234567890123456789012345678901234567890"
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
         BackColor       =   12632256
         MaxLength       =   40
         Appearance      =   0
         Caption         =   "Pekerjaan"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   -15
      Top             =   5340
      Width           =   11625
      _ExtentX        =   20505
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
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2220
         TabIndex        =   15
         Top             =   105
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "    &Delete"
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
         Picture         =   "TrOpenTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
         TabIndex        =   16
         Top             =   105
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   767
         Caption         =   ""
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
         Picture         =   "TrOpenTabungan.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9330
         TabIndex        =   17
         Top             =   120
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
         Picture         =   "TrOpenTabungan.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   18
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Caption         =   "  &Edit"
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
         Picture         =   "TrOpenTabungan.frx":083F
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   19
         Top             =   105
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         Caption         =   "  &Add"
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
         Picture         =   "TrOpenTabungan.frx":096B
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10410
         TabIndex        =   20
         Top             =   120
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
         Picture         =   "TrOpenTabungan.frx":0B16
      End
   End
End
Attribute VB_Name = "TrOpenTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim dbData1 As New ADODB.Recordset
Dim lClick As Boolean
Dim nPos As SisPos
Dim cSQL As String
Dim lEdit As Boolean
Dim StatusTutup As String
Dim cKTP As String
Dim cNomorRekening As String

Private Sub cFrekuensi_Validate(Cancel As Boolean)
  cNomorRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  Set dbData = objData.Browse(GetDSN, "Tabungan", "Rekening,close", "Rekening", sisAssign, cNomorRekening)
  If Not dbData.eof Then
    If GetNull(dbData!Close) = "1" Then
      MsgBox "Rekening tersebut sudah di TUTUP", vbInformation, Me.Caption
      Cancel = True
      initvalue
      GetEdit False
      Exit Sub
    End If
    GetMemory
    If nPos = Delete Then DeleteData
  Else
    MsgBox "Nomor Rekening tidak ada. Silahkan Ulangi pengisian", vbInformation, Me.Caption
    Cancel = True
    cFrekuensi.Default
    cFrekuensi.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cGolongan_ButtonClick()
Dim cField As String
  
  cField = "Kode,Keterangan,saldominimumdapatbunga,SaldoMinimum,SetoranMinimum"
  Set dbData = objData.Browse(GetDSN, "GolonganTabungan", cField, "Kode", sisContent, cGolongan.Text, , "Kode")
  cGolongan.Text = cGolongan.Browse(dbData)
  If Not dbData.eof Then
    cGolonganTabungan.Text = GetNull(dbData!Kode)
    cKetGolTabungan.Text = GetNull(dbData!Keterangan)
    nSaldoMinimum.Value = GetNull(dbData!SaldoMinimum)
    nSetoranMinimum.Value = GetNull(dbData!SetoranMinimum)
    
    If nPos = Add Then
      cFrekuensi.Text = GetFrekuensi("Tabungan", cCabang.Text, 1, cGolongan.Text, cUrut.Text)
    End If
  End If
  cNomorRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
  cKode.Text = Padl(Trim(cKode.Text), cKode.MaxLength, "0")
  Set dbData = objData.Browse(GetDSN, "RegisterNasabah", , "Kode", sisAssign, cCabang1.Text & "." & cKode.Text)
  If dbData.eof Then
    MsgBox "Maaf, Nomor Register Nasabah : " & cCabang1.Text & "." & cKode.Text & " Tidak Ada. Silahkan Mengulangi Pengisian !", vbInformation, Me.Caption
    Cancel = True
    cKode.Default
    cNama.Default
    cAlamat.Default
    cTelepon.Default
    cPekerjaan.Default
    cKode.SetFocus
    Exit Sub
  End If
  GetDataRegister
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  dTgl.SetFocus
  Koreksi False
  CekEdit True
End Sub

Private Sub Koreksi(ByVal lKoreksi As Boolean)
  cCabang.Enabled = lKoreksi
  cUrut.Enabled = lKoreksi
  cFrekuensi.Enabled = lKoreksi
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  initvalue
  Koreksi True
  CekEdit False
  cCabang.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  initvalue
  GetEdit True
  Koreksi True
  CekEdit False
  cCabang.SetFocus
End Sub

Private Sub DeleteData()
  cNomorRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  If MsgBox("Data Benar-benar akan Dihapus ?", vbYesNo + vbExclamation, "PEMBUKAAN RKENING TABUNGAN") = vbYes Then
    objData.Delete GetDSN, "Tabungan", "Rekening", sisAssign, cNomorRekening
    initvalue
    GetEdit False
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue

  If ValidSaving Then
    If MsgBox("Apakah Data Benar-benar sudah Valid ?", vbYesNo + vbInformation, "PEMBUKAAN REKENING") = vbYes Then
      vaField = Array("Rekening", "Tgl", "GolonganTabungan", "Kode", "PDL", "biayaadministrasi", "kasbank", "rekeningpendapatan", "close")
      vaValue = Array(cNomorRekening, dTgl.Value, cGolonganTabungan.Text, cCabang1.Text & "." & cKode.Text, cPDL.Text, nAdministrasi.Value, cPOSKas.Text, cRekPendapatanAdministrasi.Text, "0")
      If nPos = Add Then
        objData.Update GetDSN, "Tabungan", "Rekening='" & cNomorRekening & "'", vaField, vaValue
      Else
        objData.Edit GetDSN, "Tabungan", "Rekening='" & cNomorRekening & "'", vaField, vaValue
      End If
      
      If nAdministrasi.Value > 0 Then
        objData.Delete GetDSN, "bukubesar", "faktur", sisAssign, "ADM-" & cNomorRekening
        UpdKodeTr objData, msAdministrasiTabungan, cCabang.Text, "ADM-" & cNomorRekening, dTgl.Value, cPOSKas.Text, "Administrasi tabunga an " & cNama.Text, nAdministrasi.Value, 0, "D", SNow
          UpdKodeTr objData, msAdministrasiTabungan, cCabang.Text, "ADM-" & cNomorRekening, dTgl.Value, cRekPendapatanAdministrasi.Text, "Administrasi tabunga an " & cNama.Text, 0, nAdministrasi.Value, "K", SNow
      End If
      If nPos = Add Then
        'Tampkan Informasi Nomor Rekening Nasabah
        MsgBox "Nomor Rekening anda Adalah : " & cNomorRekening & " " & Chr(13) & _
                                             " Atas Nama : " & cNama.Text & " " & Chr(13) & _
                                             " Alamat        : " & cAlamat.Text & " ", vbInformation, "Informasi Nomor Rekening Tabungan"
      End If
    Else
      cGolongan.SetFocus
      Exit Sub
    End If
    initvalue
    GetEdit False
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
    
  'Cek Register Nasabah
  If Not CheckData(cCabang1.Text, "No. Register Nasabah tidak valid, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cCabang1.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKode.Text, "No. Register Nasabah tidak valid, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cCabang.Text, "No. Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cCabang.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cGolongan.Text, "No. Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cPOSKas.Text, "Pos Kas tidak boleh kosong") Then
    ValidSaving = False
    cPOSKas.SetFocus
    Exit Function
  End If
  
  If nAdministrasi.Value > 0 Then
    If Not CheckData(cRekPendapatanAdministrasi.Text, "Rekening pendapatan tidak boleh kosong") Then
      ValidSaving = False
      cRekPendapatanAdministrasi.SetFocus
      Exit Function
    End If
  End If
  
  'Jika PDL kosong maka jangan dijinkan untuk menyimpan
  If Not CheckData(cPDL.Text, "PDL tidak boleh kosong") Then
    ValidSaving = False
    cPDL.SetFocus
    Exit Function
  End If
  
End Function

Private Sub GetMemory()
Dim cField As String
Dim vaJoin
  
  cField = "d.Rekening,d.Tgl,d.GolonganTabungan,d.Kode,d.Pdl,d.Close,d.kasbank,d.biayaadministrasi,d.rekeningpendapatan,"
  cField = cField & " r. Nama, r.Alamat, r.telepon,r.Path,r.Path1,r.Pekerjaan,p.keterangan as NamaPekerjaan,"
  cField = cField & " b.Keterangan as KetGolonganTabungan, b.SetoranMinimum,b.SaldoMinimum,b.SaldoMinimumDapatBunga,"
  cField = cField & " p1.Keterangan as NamaPdl"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode=d.Kode", _
                 "Left Join GolonganTabungan b on b.Kode = d.GolonganTabungan", _
                 "Left Join Pekerjaan p on p.Kode = r.Pekerjaan", _
                 "Left Join PDL p1 on p1.Kode = d.Pdl")
  Set dbData = objData.Browse(GetDSN, "Tabungan d", cField, "d.Rekening", sisAssign, cNomorRekening, , "d.Rekening", vaJoin)
  If Not dbData.eof Then
    With dbData
      dTgl.Value = GetNull(!Tgl)
      cCabang1.Text = left(GetNull(!Kode), 2)
      cKode.Text = Right(GetNull(!Kode), 6)
      cNama.Text = GetNull(!nama)
      cAlamat.Text = GetNull(!alamat)
      cTelepon.Text = GetNull(!Telepon)
      cPekerjaan.Text = GetNull(!NamaPekerjaan)
      cGolonganTabungan.Text = GetNull(!GolonganTabungan)
      cKetGolTabungan.Text = GetNull(!KetGolonganTabungan)
      nSaldoMinimum.Value = GetNull(!SaldoMinimum)
      nSetoranMinimum.Value = GetNull(!SetoranMinimum)
      StatusTutup = GetNull(!Close)
      cPDL.Text = GetNull(!PDL, "")
      cNamaPDL.Text = GetNull(!namapdl, "")
      nAdministrasi.Value = GetNull(!biayaadministrasi)
      cPOSKas.Text = GetNull(!kasbank)
      cRekPendapatanAdministrasi.Text = GetNull(!rekeningpendapatan)
      Image1.Picture = LoadPicture(GetPicture(GetNull(dbData!Path)))
      Image2.Picture = LoadPicture(GetPicture(GetNull(dbData!Path1)))
    End With
  End If
End Sub
  
Private Sub cNama_ButtonClick()
  If nPos = Add Then
    Set dbData = objData.Browse(GetDSN, "RegisterNasabah r", "r.Nama,r.Alamat,r.Kode,r.Telepon,r.Pekerjaan,r.path,r.Path1,p.keterangan as namaPekerjaan", "Nama", sisContent, cNama.Text, , "r.Kode,r.Nama", Array("left join pekerjaan p on p.kode = r.pekerjaan"))
    cNama.Text = cNama.Browse(dbData)
    If Not dbData.eof Then
      cKode.Text = Right(GetNull(dbData!Kode), 6)
      cNama.Text = GetNull(dbData!nama)
      cAlamat.Text = GetNull(dbData!alamat)
      cTelepon.Text = GetNull(dbData!Telepon)
      cPekerjaan.Text = GetNull(dbData!NamaPekerjaan)
      cUrut.Text = cKode.Text
      GetGambar Image1, Image2, GetNull(GetNull(dbData!Path), ""), GetNull(GetNull(dbData!Path1), "")
    End If
  End If
End Sub

Private Sub initvalue()
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  dTgl.Value = Date
  cKode.Default
  cNama.Default
  cAlamat.Default
  cTelepon.Default
  cPekerjaan.Default
  cGolonganTabungan.Default
  cKetGolTabungan.Default
  nSaldoMinimum.Value = 0
  nSetoranMinimum.Value = 0
  cPDL.Default
  cNamaPDL.Default
  nAdministrasi.Default
  cPOSKas.Default
  cRekPendapatanAdministrasi.Default
  
  Image1.Picture = LoadPicture(GetPicture(""))
  Image2.Picture = LoadPicture(GetPicture(""))
  CekEdit True
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    GetEdit False
    initvalue
  End If
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  BiSAFrame4.Enabled = lPar
  lEdit = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar
End Sub

Private Sub cPdl_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "PDL", "Kode,Keterangan", "Kode", sisContent, cPDL.Text, , "Kode")
  cPDL.Text = cPDL.Browse(dbData)
  If Not dbData.eof Then
    cNamaPDL.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cPOSKas_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "Rekening", "Kode,Keterangan,Jenis", "Kode", sisContent, cPOSKas.Text, " and Jenis = 'D' and left(kode,1)='1'", "Kode")
  cPOSKas.Text = cPOSKas.Browse(db)
  If Not db.eof Then
    cPOSKas.Text = GetNull(db!Kode, "")
  End If
End Sub

Private Sub cRekPendapatanAdministrasi_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "Rekening", "Kode,Keterangan,Jenis", "Kode", sisContent, cRekPendapatanAdministrasi.Text, " and jenis = 'D' and left(kode,1) = '4'", "Kode")
  cRekPendapatanAdministrasi.Text = cRekPendapatanAdministrasi.Browse(db)
  If Not db.eof Then
    cRekPendapatanAdministrasi.Text = GetNull(db!Kode, "")
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(Trim(cUrut.Text), cUrut.MaxLength, "0")
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Then
    Cancel = True
    dTgl.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Dim n As Single
  
  CenterForm Me
  GetEdit False
  Me.Top = 0
  initvalue
  cCabang1.Text = aCfg(msKodeCabang, "")
  cCabang.Text = cCabang1.Text
  
  TabIndex dTgl, n
  TabIndex cCabang1, n
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cPDL, n
  TabIndex nAdministrasi, n
  TabIndex cPOSKas, n
  TabIndex cRekPendapatanAdministrasi, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub GetDataRegister()
  Set dbData = objData.Browse(GetDSN, "RegisterNasabah r", "r.Nama,r.Alamat,r.telepon,r.Kode,r.Path,r.Path1,p.Keterangan", "r.Kode", sisAssign, cCabang1.Text & "." & cKode.Text, , , _
                              Array("Left Join Pekerjaan p on p.Kode = r.Pekerjaan"))
  If Not dbData.eof Then
    cKode.Text = Right(GetNull(dbData!Kode), 6)
    cNama.Text = GetNull(dbData!nama)
    cAlamat.Text = GetNull(dbData!alamat)
    cTelepon.Text = GetNull(dbData!Telepon)
    cPekerjaan.Text = GetNull(dbData!Keterangan)
    cUrut.Text = cKode.Text
    GetGambar Image1, Image2, GetNull(GetNull(dbData!Path), ""), GetNull(GetNull(dbData!Path1), "")
  End If
End Sub

Private Function GetRegPDL(ByVal cPDL As String) As Double
  GetRegPDL = 1
  Set dbData = objData.Browse(GetDSN, "Tabungan", "Max(NoPDL) as NoPDL", "PDL", sisAssign, cPDL)
  If Not dbData.eof Then
    GetRegPDL = GetNull(dbData!NoPDL) + 1
  End If
End Function

Private Sub CekEdit(ByVal lKoreksi As Boolean)
  dTgl.Enabled = lKoreksi
  cCabang1.Enabled = lKoreksi
  cKode.Enabled = lKoreksi
  cNama.Enabled = lKoreksi
  
  If lKoreksi = False Then
    dTgl.BackColor = &H8000000F
    cCabang1.BackColor = &H8000000F
    cKode.BackColor = &H8000000F
    cNama.BackColor = &H8000000F
  Else
    dTgl.BackColor = &H80000005
    cCabang1.BackColor = &H80000005
    cKode.BackColor = &H80000005
    cNama.BackColor = &H80000005
  End If
End Sub
