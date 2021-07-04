VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trOpenDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMBUKAAN DEPOSITO"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11775
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   6120
      Left            =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10795
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   750
         Left            =   180
         Top             =   1230
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   1323
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
            Height          =   330
            Left            =   4320
            TabIndex        =   23
            Top             =   210
            Width           =   420
            _ExtentX        =   741
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
         Begin BiSATextBoxProject.BiSATextBox cUrut 
            Height          =   330
            Left            =   3390
            TabIndex        =   24
            Top             =   210
            Width           =   900
            _ExtentX        =   1588
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
         Begin BiSATextBoxProject.BiSABrowse cGolongan 
            Height          =   330
            Left            =   2550
            TabIndex        =   25
            Top             =   210
            Width           =   825
            _ExtentX        =   1455
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
         Begin BiSATextBoxProject.BiSATextBox cCabang 
            Height          =   330
            Left            =   330
            TabIndex        =   26
            Top             =   210
            Width           =   2190
            _ExtentX        =   3863
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
      End
      Begin BiSATextBoxProject.BiSATextBox cKode 
         Height          =   330
         Left            =   2175
         TabIndex        =   0
         Top             =   105
         Width           =   795
         _ExtentX        =   1402
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame8 
         Height          =   2430
         Left            =   8100
         Top             =   840
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   4286
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
            Height          =   2115
            Left            =   60
            Stretch         =   -1  'True
            Top             =   255
            Width           =   3465
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame7 
         Height          =   2430
         Left            =   5925
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   4286
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
            Height          =   2115
            Left            =   60
            Stretch         =   -1  'True
            Top             =   255
            Width           =   2055
         End
      End
      Begin BiSANumberBoxProject.BiSANumberBox cJangkaWaktu 
         Height          =   330
         Left            =   150
         TabIndex        =   1
         Top             =   2775
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483634
         Caption         =   "Jangka Waktu"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   150
         TabIndex        =   2
         Top             =   2055
         Width           =   2955
         _ExtentX        =   5212
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
         Caption         =   "Tgl Valuta"
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
         Left            =   150
         TabIndex        =   3
         Top             =   480
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
         Caption         =   "Nama Deposan"
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
      Begin BiSATextBoxProject.BiSATextBox cCabang1 
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   105
         Width           =   1995
         _ExtentX        =   3519
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
         MaxLength       =   2
         Caption         =   "No Register"
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
         Left            =   150
         TabIndex        =   5
         Top             =   840
         Width           =   5595
         _ExtentX        =   9869
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
         Caption         =   "Alamat"
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
      Begin BiSATextBoxProject.BiSABrowse cGolonganDeposito 
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   2415
         Width           =   2265
         _ExtentX        =   3995
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
         Caption         =   "Gol Deposito"
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
         Left            =   2430
         TabIndex        =   7
         Top             =   2415
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
         Left            =   135
         TabIndex        =   8
         Top             =   3555
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Caption         =   "Jatuh Tempo"
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
         Left            =   1740
         Top             =   3900
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
         BackColor       =   -2147483633
         Begin VB.OptionButton optARO 
            Caption         =   "&Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   10
            Top             =   75
            Width           =   1020
         End
         Begin VB.OptionButton optARO 
            Caption         =   "&Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
         Left            =   150
         TabIndex        =   11
         Top             =   3150
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Suku Bunga"
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
      Begin BiSANumberBoxProject.BiSANumberBox nFinalti 
         Height          =   330
         Left            =   150
         TabIndex        =   21
         Top             =   4425
         Width           =   2400
         _ExtentX        =   4233
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
         Caption         =   "Finalty"
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
      Begin BiSATextBoxProject.BiSABrowse cPdl 
         Height          =   345
         Left            =   165
         TabIndex        =   27
         Top             =   4785
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   609
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
      Begin BiSATextBoxProject.BiSATextBox cNamaPDL 
         Height          =   330
         Left            =   1770
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   5160
         Width           =   3615
         _ExtentX        =   6376
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
      Begin BiSATextBoxProject.BiSABrowse cRekSimpanan 
         Height          =   330
         Left            =   150
         TabIndex        =   29
         Top             =   5550
         Width           =   4350
         _ExtentX        =   7673
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
         Caption         =   "Rek. Simpanan"
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
      Begin VB.Label Label1 
         Caption         =   "% * Bunga"
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
         Left            =   2625
         TabIndex        =   22
         Top             =   4500
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "Sistem ARO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   14
         Top             =   4020
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "Bulan"
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
         Left            =   2460
         TabIndex        =   13
         Top             =   2820
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "% per tahun"
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
         Left            =   2460
         TabIndex        =   12
         Top             =   3210
         Width           =   1095
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   6105
      Width           =   11775
      _ExtentX        =   20770
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
         Left            =   2235
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
         Picture         =   "trOpenDeposito.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3405
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
         Picture         =   "trOpenDeposito.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9435
         TabIndex        =   17
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
         Picture         =   "trOpenDeposito.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1185
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
         Picture         =   "trOpenDeposito.frx":083F
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   120
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
         Picture         =   "trOpenDeposito.frx":096B
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10515
         TabIndex        =   20
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
         Picture         =   "trOpenDeposito.frx":0B16
      End
   End
End
Attribute VB_Name = "trOpenDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim nPos As SisPos
Dim lEdit As Boolean
Dim cRekening As String

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "RegisterNasabah", "Alamat,Nama,Kode,Telepon,path,Path1", "Alamat", sisContent, cAlamat.Text, , "Alamat")
  cAlamat.Text = cAlamat.Browse(dbData, Array("Alamat", "Nama", "Kode"))
  If Not dbData.eof Then
    GetRegister
  End If
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  Set dbData = objData.Browse(GetDSN, "Deposito", "Rekening,NominalDeposito", "Rekening", sisAssign, cRekening, "And Status <> '1'")
  If Not dbData.eof Then
    If nPos = Add Then
      MsgBox "Nomor Rekening sudah ada. Silahkan Ulangi pengisian", vbExclamation + vbOKOnly
      Cancel = True
      cFrekuensi.Default
      cFrekuensi.SetFocus
      Exit Sub
    End If
    If nPos = Edit And GetNull(dbData!nominaldeposito) > 0 Then
      MsgBox "Nominal Deposito Sudah Diisi, Hapus Transaksi Teller Terlebih Dahulu", vbExclamation + vbOKOnly
      Cancel = True
      initvalue
      GetEdit False
      Exit Sub
    End If
    GetMemory
    If nPos = Delete Then
      DeleteData
      Exit Sub
    End If
  Else
    MsgBox "Data tidak ada..", vbInformation, Me.Caption
  End If
End Sub

Private Sub DeleteData()
 Set dbData = objData.Browse(GetDSN, "MutasiDeposito", , "rekening", sisAssign, cRekening)
 If Not dbData.eof Then
  MsgBox "Rekening ini sudah pernah melakukan Mutasi Deposito. Hapus dulu semua mutasi deposito..", vbInformation
  Exit Sub
 End If
 If MsgBox("Data Benar-benar akan Dihapus", vbQuestion + vbYesNo) = vbYes Then
    objData.Delete GetDSN, "Deposito", "Rekening", sisAssign, cRekening
    MsgBox "Data Sudah Dihapus", vbExclamation + vbOKOnly
    GetEdit False
    initvalue
    Exit Sub
 End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganDeposito", "Kode", cGolongan, "Kode,Keterangan,Lama,Bunga")
  If Not dbData.eof Then
    cGolonganDeposito.Text = GetNull(dbData!Kode, "")
    cKetGolDeposito.Text = GetNull(dbData!Keterangan, "")
    cJangkaWaktu.Value = GetNull(dbData!Lama)
    nBunga.Value = GetNull(dbData!bunga)
    
    If nPos = Add Then
      cFrekuensi.Text = GetFrekuensi("Deposito", cCabang.Text, 2, cGolongan.Text, cUrut.Text)
    End If
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Then
    cGolongan_ButtonClick
  End If
End Sub


Private Sub cJangkaWaktu_Validate(Cancel As Boolean)
  dTempo.Value = DateAdd("m", Val(cJangkaWaktu.Value), dTgl.Value)
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
Dim cNoRegister As String

  cKode.Text = Padl(Trim(cKode.Text), cKode.MaxLength, "0")
  cNoRegister = cCabang1.Text & "." & cKode.Text
  Set dbData = objData.Browse(GetDSN, "RegisterNasabah", "Kode,Nama,Alamat,Telepon,path,path1", "Kode", sisAssign, cNoRegister)
  If dbData.eof Then
    MsgBox "Maaf, Nomor Register Deposan : " & cCabang1.Text & "." & cKode.Text & " Tidak Ada. Silahkan Mengulangi Pengisian !", vbInformation
    cKode.Default
    cNama.Default
    cAlamat.Default
    cKode.SetFocus
    Exit Sub
  End If
  GetRegister
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  cCabang1.SetFocus
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  initvalue
  cCabang.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  initvalue
  cCabang.SetFocus
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim cNoRegister As String
Dim cRekening As String

  If ValidSaving Then
    If MsgBox("Apakah Data Benar-benar sudah Valid ?", vbYesNo + vbInformation) = vbYes Then
      cNoRegister = cCabang1.Text & "." & cKode.Text
      cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
      objData.Delete GetDSN, "Deposito", "Rekening", sisAssign, cRekening
      vaField = Array("Rekening", "Tgl", "jthTmp", "PDL", "Kode", "GolonganDeposito", "SistemARO", "DateTime", "SukuBunga", "PersentaseFinalti", "Lama", "rekeningsimpanan")
      vaValue = Array(cRekening, dTgl.Value, dTempo.Value, cPDL.Text, cNoRegister, cGolonganDeposito.Text, GetOpt(optARO), Format(Now, "yyyy-mm-dd hh:mm:ss"), nBunga.Value, nFinalti.Value, cJangkaWaktu.Value, cRekSimpanan.Text)
      objData.Add GetDSN, "Deposito", vaField, vaValue
    Else
      cKode.SetFocus
      Exit Sub
    End If
    initvalue
    GetEdit False
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cGolongan.Text, "Golongan Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Frekuensi Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cFrekuensi.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKode.Text, "Kode Deposan Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
   
  If Not CheckData(cGolonganDeposito.Text, "Golongan Deposito Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cGolonganDeposito.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cRekSimpanan.Text, "Rekeing Simpanan Deposito Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cRekSimpanan.SetFocus
    Exit Function
  End If
End Function

Private Sub GetMemory()
Dim vaJoin
Dim cFields As String

  cFields = " d.lama as lamaDeposito,d.Tgl,d.Kode,d.GolonganDeposito,d.SukuBunga,d.SistemARO,d.JthTmp,r. Nama, r.Alamat, r.telepon,r.Path,r.Path1,d.PDL,p.Keterangan as NamaPDL,"
  cFields = cFields & " b.Keterangan as KetGolonganDeposito,b.Lama,d.PersentaseFinalti"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode = d.Kode", _
                 "Left Join GolonganDeposito b on b.Kode = d.GolonganDeposito", _
                 "Left Join PDL p on p.Kode = d.PDL")
                 
  Set dbData = objData.Browse(GetDSN, "Deposito d", cFields, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    With dbData
      dTgl.Value = GetNull(!Tgl, "")
      cCabang1.Text = Mid(GetNull(dbData!Kode, ""), 1, 2)
      cKode.Text = Right(GetNull(!Kode), 6)
      cNama.Text = GetNull(!nama, "")
      cAlamat.Text = GetNull(!alamat, "")
      cGolonganDeposito.Text = GetNull(!GolonganDeposito, "")
      cKetGolDeposito.Text = GetNull(!KetGolonganDeposito, "")
      cJangkaWaktu.Value = GetNull(!LamaDeposito, "")
      dTempo.Value = GetNull(!jthtmp, "")
      nBunga.Value = GetNull(!SukuBunga)
      nFinalti.Value = GetNull(dbData!PersentaseFinalti)
      SetOpt optARO, GetNull(!SistemARO, "")
      cPDL.Text = GetNull(dbData!PDL, "")
      cNamaPDL.Text = GetNull(dbData!namapdl, "")
      GetImage GetNull(dbData!Path, ""), GetNull(dbData!Path1, "")
    End With
  End If
End Sub
  
Private Sub cNama_ButtonClick()
  If nPos = Add Then
    Set dbData = objData.Browse(GetDSN, "RegisterNasabah", "Nama,Alamat,Kode,Telepon,path,Path1", "Nama", sisContent, cNama.Text, , "Nama")
    cNama.Text = cNama.Browse(dbData, Array("Nama", "Alamat", "Kode"))
    If Not dbData.eof Then
      GetRegister
    End If
  End If
End Sub

Private Sub GetRegister()
  cKode.Text = Right(GetNull(dbData!Kode), 6)
  cNama.Text = GetNull(dbData!nama)
  cAlamat.Text = GetNull(dbData!alamat)
  GetImage GetNull(dbData!Path, ""), GetNull(dbData!Path1, "")
  cUrut.Text = cKode.Text
End Sub

Private Sub initvalue()
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  dTgl.Value = Date
  cKode.Default
  cNama.Default
  cAlamat.Default
  cGolonganDeposito.Default
  cKetGolDeposito.Default
  cJangkaWaktu.Default
  dTempo.Value = Date
  Image1.Picture = LoadPicture(GetPicture(""))
  Image2.Picture = LoadPicture(GetPicture(""))
  optARO(0).Value = True
  nBunga.Value = 0
  nFinalti.Value = 0
  cPDL.Default
  cNamaPDL.Default
  cCabang.Text = aCfg(msKodeCabang, "")
  cCabang1.Text = cCabang.Text
  cFrekuensi.Enabled = False
  cRekSimpanan.Default
  If nPos <> Add Then
    cFrekuensi.Enabled = True
  End If
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
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub cPdl_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "PDL", "Kode,Keterangan", "Kode", sisContent, cPDL.Text, , "Kode")
  If Not dbData.eof Then
    cPDL.Text = cPDL.Browse(dbData, Array("Kode", "Keterangan"), , Array(4, 20))
    cPDL.Text = GetNull(dbData!Kode, "")
    cNamaPDL.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cRekSimpanan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "tabungan", "rekening", "kode", sisAssign, cCabang1.Text & "." & cKode.Text)
  If Not dbData.eof Then
    cRekSimpanan.Text = cRekSimpanan.Browse(dbData)
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Then
    Cancel = True
    dTgl.SetFocus
  End If
  dTempo.Value = DateAdd("m", Val(cJangkaWaktu.Value), dTgl.Value)
End Sub

Private Sub Form_Load()
  Dim n As Single
  
  CenterForm Me
  Me.Top = 0
  initvalue
  GetEdit False
    
  TabIndex cCabang1, n
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex dTgl, n
  TabIndex cJangkaWaktu, n
  TabIndex nBunga, n
  TabIndex optARO(0), n
  TabIndex optARO(1), n
  TabIndex nFinalti, n
  TabIndex cPDL, n
  TabIndex cRekSimpanan, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub GetImage(cPhoto As String, cSpecimen As String)
On Error Resume Next

  Image1.Picture = LoadPicture(GetPicture(cPhoto))
  Image2.Picture = LoadPicture(GetPicture(cSpecimen))
End Sub

Private Sub Label3_Click()

End Sub

Private Sub optARO_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 40 Then
    SendKeysA vbKeyTab, True
  End If
End Sub
