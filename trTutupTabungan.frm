VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trTutupTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENUTUPAN REKENING SIMPANAN"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10020
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5205
      Left            =   0
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   9181
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   3495
         Left            =   225
         Top             =   1650
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   6165
         Caption         =   "MUTASI"
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
         Begin BiSANumberBoxProject.BiSANumberBox nSaldoAkhir 
            Height          =   330
            Left            =   105
            TabIndex        =   16
            Top             =   330
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   582
            Appearance      =   0
            MinValue        =   0
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
            Caption         =   "Saldo Akhir"
            CaptionWidth    =   2000
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
         Begin BiSANumberBoxProject.BiSANumberBox nBiayaAdm 
            Height          =   330
            Left            =   105
            TabIndex        =   17
            Top             =   1050
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   582
            Appearance      =   0
            MinValue        =   0
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
            Caption         =   "Biaya Administrasi"
            CaptionWidth    =   2000
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
         Begin BiSANumberBoxProject.BiSANumberBox nSisaPenarikan 
            Height          =   330
            Left            =   105
            TabIndex        =   18
            Top             =   1410
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   582
            Appearance      =   0
            MinValue        =   0
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
            Caption         =   "Sisa Penarikan"
            CaptionWidth    =   2000
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
         Begin BiSANumberBoxProject.BiSANumberBox nSelisih 
            Height          =   330
            Left            =   105
            TabIndex        =   19
            Top             =   1785
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   582
            MinValue        =   0
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Selisih Pembulatan"
            CaptionWidth    =   2000
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
         Begin BiSANumberBoxProject.BiSANumberBox nPenarikan 
            Height          =   375
            Left            =   105
            TabIndex        =   20
            Top             =   2385
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   661
            Appearance      =   0
            MinValue        =   0
            Enabled         =   0   'False
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16579821
            Caption         =   "Penarikan Tunai"
            CaptionWidth    =   2000
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
         Begin BiSANumberBoxProject.BiSANumberBox nBiayalain 
            Height          =   330
            Left            =   105
            TabIndex        =   21
            Top             =   690
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   582
            Appearance      =   0
            MinValue        =   0
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
            Caption         =   "Biaya lain-Lain"
            CaptionWidth    =   2000
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
            X1              =   135
            X2              =   4335
            Y1              =   2235
            Y2              =   2235
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   3495
         Left            =   4725
         Top             =   1650
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   6165
         Caption         =   "INFORMASI TELLER"
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
         Begin BiSATextBoxProject.BiSATextBox cDK 
            Height          =   330
            Left            =   3180
            TabIndex        =   14
            Top             =   600
            Width           =   1155
            _ExtentX        =   2037
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
            Appearance      =   0
            Caption         =   "D/K"
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
         Begin BiSATextBoxProject.BiSATextBox cKode 
            Height          =   330
            Left            =   105
            TabIndex        =   9
            Top             =   600
            Width           =   2220
            _ExtentX        =   3916
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
            Caption         =   "Kode Trans"
            CaptionWidth    =   1300
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
         Begin BiSATextBoxProject.BiSATextBox cNamaKode 
            Height          =   330
            Left            =   105
            TabIndex        =   10
            Top             =   945
            Width           =   4935
            _ExtentX        =   8705
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
            Caption         =   "Keterangan"
            CaptionWidth    =   1300
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
         Begin BiSATextBoxProject.BiSATextBox cTeller 
            Height          =   330
            Left            =   105
            TabIndex        =   11
            Top             =   255
            Width           =   4230
            _ExtentX        =   7461
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
            Caption         =   "Teller"
            CaptionWidth    =   1300
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
         Begin BiSATextBoxProject.BiSATextBox cRekening 
            Height          =   330
            Left            =   105
            TabIndex        =   12
            Top             =   1290
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
            FontName        =   "Verdana"
            BackColor       =   12632256
            Enabled         =   0   'False
            Appearance      =   0
            Caption         =   "Rek Jurnal"
            CaptionWidth    =   1300
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
            Left            =   105
            TabIndex        =   13
            Top             =   1635
            Width           =   4935
            _ExtentX        =   8705
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
            Caption         =   "Nama Rek"
            CaptionWidth    =   1300
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
            Left            =   105
            TabIndex        =   15
            Top             =   1980
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   582
            Appearance      =   0
            MinValue        =   0
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
            BackColor       =   16579821
            Caption         =   "Saldo Teller"
            CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   4110
         TabIndex        =   0
         Top             =   180
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         Text            =   "12"
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   2445
         TabIndex        =   1
         Top             =   180
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   582
         Text            =   "12"
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
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Text            =   "12"
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
         MaxLength       =   2
         Caption         =   "No Rekening"
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
      Begin BiSATextBoxProject.BiSATextBox cUrut 
         Height          =   330
         Left            =   3165
         TabIndex        =   3
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   180
         TabIndex        =   4
         Top             =   540
         Width           =   5055
         _ExtentX        =   8916
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   6015
         _ExtentX        =   10610
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   180
         TabIndex        =   6
         Top             =   1275
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Caption         =   "Tanggal"
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
      Left            =   0
      Top             =   5205
      Width           =   9990
      _ExtentX        =   17621
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
         Left            =   7545
         TabIndex        =   7
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
         Picture         =   "trTutupTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   8625
         TabIndex        =   8
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
         Picture         =   "trTutupTabungan.frx":0416
      End
   End
End
Attribute VB_Name = "trTutupTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub initvalue()
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
  dTgl.Value = Date
  nSaldoAkhir.Value = 0
  nBiayalain.Value = 0
  nBiayaAdm.Value = 0
  nSisaPenarikan.Value = 0
  nSelisih.Value = 0
  nPenarikan.Value = 0
  
  cTeller.Default
  cKode.Default
  cDK.Default
  cNamaKode.Default
  cRekening.Default
  cNamaRekening.Default
  nSaldo.Value = 0
  KodeTransaksi
  GetSaldoTeller
End Sub

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Tabungan t", "r.Alamat,r.nama,t.Rekening", "r.Alamat", sisContent, cAlamat.Text, "And t.Close <>'1' And r.Nama Like '" & cNama.Text & "%'", "r.Nama", Array("Left Join RegisterNasabah r on r.Kode=t.Kode"))
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
  Set dbData = objData.Browse(GetDSN, "GolonganTabungan", "Kode,Keterangan,AdministrasiTutup,SaldoMinimum", "Kode", sisContent, cGolongan.Text)
  cGolongan.Text = cGolongan.Browse(dbData)
  If Not dbData.eof Then
    nBiayaAdm.Value = GetNull(dbData!administrasitutup)
    nBiayalain.Value = GetNull(dbData!SaldoMinimum)
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  cGolongan_ButtonClick
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub GetData()
Dim cField As String
Dim vaJoin
Dim cRekening As String

  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cField = "r.Nama,r.Alamat"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode=t.Kode")
  Set dbData = objData.Browse(GetDSN, "Tabungan t", cField, "t.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    cNama.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    nSaldoAkhir.Value = GetSaldoTab(objData, cRekening, Date)
    nSisaPenarikan.Value = nSaldoAkhir.Value - nBiayaAdm.Value - nBiayalain.Value
   End If
End Sub

Private Sub cmdSimpan_Click()
Dim cRekening As String
Dim cKodeAdmin As String
Dim cKodeBunga As String
Dim cKodePajak As String
Dim cKodePembulatan As String
Dim cKodePenarikan As String
Dim cRekAdmin As String
Dim cRekPembulatan As String
Dim cRekPenarikan As String
Dim cFakturTabungan As String

  If ValidSaving Then
    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    If MsgBox("Data benar-benar sudah valid dan akan disimpan?", vbYesNo + vbInformation) = vbYes Then
      cFakturTabungan = GetLastFaktur(fkt_MutasiTabungan, dTgl.Value, True)
      UpdUrutFaktur objData, cFakturTabungan
      
      GetKode aCfg(msKodeAdministrasi), cKodeAdmin, cRekAdmin
      GetKode aCfg(msKodePembulatankas), cKodePembulatan, cRekPembulatan
      GetKode aCfg(msKodePenarikanTunai), cKodePenarikan, cRekPenarikan
      
'      UpdMutasiTabungan objData, cKodeAdmin, cFakturTabungan, dTgl.Value, cRekening, nBiayaAdm.Value + nBiayaLain.Value, , "Admin. Tutup Rek Tabungan an. " & cNama.Text, True, "D", cRekAdmin, SNow
'      UpdMutasiTabungan objData, cKodePembulatan, cFakturTabungan, dTgl.Value, cRekening, nSelisih.Value, , "Pembulatan Kas an. " & cNama.Text, True, "D", cRekPembulatan, SNow
'      UpdMutasiTabungan objData, cKodePenarikan, cFakturTabungan, dTgl.Value, cRekening, nPenarikan.Value, , "Penarikan Tutup Tabungan an. " & cNama.Text, True, "K", cRekPenarikan, SNow
      
      UpdMutasiTabungan objData, cKodeAdmin, cFakturTabungan, dTgl.Value, cRekening, nBiayaAdm.Value + nBiayalain.Value, , "Admin. Tutup Rek Tabungan an. " & cNama.Text, True, "K", cRekAdmin, SNow
      UpdMutasiTabungan objData, cKodePembulatan, cFakturTabungan, dTgl.Value, cRekening, nSelisih.Value, , "Pembulatan Kas an. " & cNama.Text, True, "D", cRekPembulatan, SNow
      UpdMutasiTabungan objData, cKodePenarikan, cFakturTabungan, dTgl.Value, cRekening, nPenarikan.Value, , "Penarikan Tutup Tabungan an. " & cNama.Text, True, "K", cRekPenarikan, SNow
      
      objData.Edit GetDSN, "tabungan", "Rekening='" & cRekening & "'", Array("close", "TglPenutupan"), Array("1", dTgl.Value)
      
    End If
    initvalue
    cCabang.SetFocus
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
 
  If Not CheckData(cGolongan.Text, "Invalid kode rekening..!") Then
    ValidSaving = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cUrut.Text, "Invalid kode rekening..!") Then
    ValidSaving = False
    cUrut.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Invalid kode rekening..!") Then
    ValidSaving = False
    cFrekuensi.SetFocus
    Exit Function
  End If
End Function

Private Sub GetKode(ByVal cDefault, cKD As String, cKT As String)
  cKD = cDefault
  Set dbData = objData.Browse(GetDSN, "KodeTransaksi", "Kode,Rekening", "Kode", sisAssign, cDefault)
  If Not dbData.eof Then
    cKT = GetNull(dbData!Rekening, "")
  End If
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cRekening As String
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  Set dbData = objData.Browse(GetDSN, "tabungan", "Rekening,Close", "Rekening", sisAssign, cRekening, "And Close <>'1'")
  If dbData.eof Then
    MsgBox "Rekening dengan nomor: " & cRekening & " Tidak ada. Silahkan mengulangi pengisian !", vbOKOnly + vbInformation, "Blokir Tabungan"
    Cancel = True
    cFrekuensi.SetFocus
    Exit Sub
  End If
  GetData
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
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
  Me.Top = 0
  initvalue
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex dTgl, n
  TabIndex nSelisih, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub KodeTransaksi()
  cTeller.Text = cusername
  cKode.Text = aCfg(msKodePenarikanTunai, "")
  Set dbData = objData.Browse(GetDSN, "KodeTransaksi k", "k.Kode,k.Keterangan,k.DK,k.Kas,k.Rekening,r.Keterangan as NamaRekening", "k.Kode", sisAssign, cKode.Text, " and (t.Level > 0 or " & nUserLevel & " = 0)", , _
               Array("Left Join KodetransaksiTeller t on k.Kode = t.Kode and Level = " & nUserLevel, _
                     "Left Join Rekening r on r.Kode = k.Rekening"))
                     
  If Not dbData.eof Then
    cNamaKode.Text = GetNull(dbData!Keterangan, "")
    cDK.Text = GetNull(dbData!DK, "")
    cRekening.Text = GetNull(dbData!Rekening, "")
    cNamaRekening.Text = GetNull(dbData!NamaRekening, "")
  End If
End Sub

Private Sub GetSaldoTeller()
  Set dbData = objData.Browse(GetDSN, "BukuBesar", "Faktur,Keterangan,sum(Debet) as Debet,sum(Kredit) as Kredit", "Tgl", sisLTEqual, Format(Date, "yyyy-mm-dd"), " and Rekening = '" & cKasTeller & "' Group By Rekening", "Tgl,Rekening,ID")
  If Not dbData.eof Then
    nSaldo.Value = GetNull(dbData!Debet) - GetNull(dbData!Kredit)
  End If
End Sub

Private Sub nSelisih_Validate(Cancel As Boolean)
  nPenarikan.Value = nSisaPenarikan.Value - nSelisih.Value
End Sub
