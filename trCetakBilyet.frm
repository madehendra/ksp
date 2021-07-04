VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trCetakBilyet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK BILYET DEPOSITO"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6930
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5775
      Left            =   0
      Top             =   -15
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   10186
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
      Begin BiSADateProject.BiSADate dTanggal 
         Height          =   330
         Left            =   4425
         TabIndex        =   0
         Top             =   105
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   582
         Value           =   "15-11-2009"
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
         Caption         =   "TGL"
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
      Begin BiSANumberBoxProject.BiSANumberBox cJangkaWaktu 
         Height          =   330
         Left            =   150
         TabIndex        =   1
         Top             =   3255
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Appearance      =   0
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
         BackColor       =   12632256
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
         Top             =   2535
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
         Value           =   "15-11-2009"
         Appearance      =   0
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
         ForeColor       =   -2147483640
         Enabled         =   0   'False
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
         Top             =   465
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
         Enabled         =   0   'False
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   810
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
         TabIndex        =   5
         Top             =   2895
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
         TabIndex        =   6
         Top             =   2895
         Width           =   3270
         _ExtentX        =   5768
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
         Left            =   150
         TabIndex        =   7
         Top             =   4125
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   582
         Value           =   "15-11-2009"
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
         Left            =   1770
         Top             =   3615
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   847
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
         Enabled         =   0   'False
         Begin VB.OptionButton optARO 
            Caption         =   "&Ya"
            BeginProperty Font 
               Name            =   "Verdana"
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
         Begin VB.OptionButton optARO 
            Caption         =   "&Tidak"
            BeginProperty Font 
               Name            =   "Verdana"
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
            TabIndex        =   8
            Top             =   75
            Width           =   1020
         End
      End
      Begin BiSATextBoxProject.BiSATextBox cCabang 
         Height          =   330
         Left            =   150
         TabIndex        =   10
         Top             =   105
         Width           =   2010
         _ExtentX        =   3545
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
      Begin BiSATextBoxProject.BiSATextBox cUrut 
         Height          =   330
         Left            =   2985
         TabIndex        =   11
         Top             =   105
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   3900
         TabIndex        =   12
         Top             =   105
         Width           =   435
         _ExtentX        =   767
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
      Begin BiSANumberBoxProject.BiSANumberBox nBunga 
         Height          =   330
         Left            =   150
         TabIndex        =   13
         Top             =   4500
         Width           =   2445
         _ExtentX        =   4313
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
         Caption         =   "Bunga"
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
         Left            =   2655
         TabIndex        =   14
         Top             =   4500
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "NOMINAL"
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
      Begin BiSATextBoxProject.BiSATextBox cNoBilyet 
         Height          =   330
         Left            =   150
         TabIndex        =   15
         Top             =   2190
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   582
         Text            =   "1234567890"
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
         Caption         =   "No Bilyet"
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
      Begin BiSATextBoxProject.BiSATextBox cSeri 
         Height          =   330
         Left            =   4035
         TabIndex        =   16
         Top             =   2190
         Width           =   1965
         _ExtentX        =   3466
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
         Caption         =   "NO. SERI"
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
         TabIndex        =   20
         Top             =   105
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
      Begin BiSATextBoxProject.BiSATextBox cNamakasir 
         Height          =   330
         Left            =   150
         TabIndex        =   21
         Top             =   4890
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   582
         Text            =   "12"
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
         Caption         =   "Nama Kasir"
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
      Begin BiSATextBoxProject.BiSATextBox cPimpinan 
         Height          =   330
         Left            =   135
         TabIndex        =   22
         Top             =   5250
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   582
         Text            =   "12"
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
         Caption         =   "Pimpinan"
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
      Begin VB.PictureBox CR 
         Height          =   480
         Left            =   6135
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   27
         Top             =   750
         Width           =   1200
      End
      Begin BiSATextBoxProject.BiSABrowse cKota 
         Height          =   330
         Left            =   150
         TabIndex        =   25
         Top             =   1170
         Width           =   4530
         _ExtentX        =   7990
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
         Caption         =   "Kota"
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
      Begin BiSATextBoxProject.BiSABrowse cRekTabungan 
         Height          =   330
         Left            =   150
         TabIndex        =   26
         Top             =   1515
         Width           =   4530
         _ExtentX        =   7990
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
         Caption         =   "Rek Tabungan"
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
      Begin VB.PictureBox CR2 
         Height          =   480
         Left            =   6120
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   28
         Top             =   1260
         Width           =   1200
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
         TabIndex        =   18
         Top             =   3315
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "Sistem ARO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   17
         Top             =   3765
         Width           =   1470
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5745
      Width           =   6915
      _ExtentX        =   12197
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   465
         Left            =   75
         TabIndex        =   23
         Top             =   105
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   820
         Caption         =   "Cetak Bilyet Hal 1 (Depan)"
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
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   465
         Left            =   5655
         TabIndex        =   19
         Top             =   75
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
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
         Picture         =   "trCetakBilyet.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton2 
         Height          =   465
         Left            =   2865
         TabIndex        =   24
         Top             =   75
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   820
         Caption         =   "Cetak Bilyet Hal 2 (Blkng)"
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
      End
   End
End
Attribute VB_Name = "trCetakBilyet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim nPos As SisPos
Dim cBulan As String
Dim cRekening As String

Private Sub BiSAButton1_Click()
With CR
'      .ReportFileName = App.Path & "\rptdepositoa.rpt"
'
'      .ParameterFields(0) = "dTglValuta;" & Format(dTgl.Value, "dd MMM yyyy") & ";True"
'      .ParameterFields(1) = "dJangkaWaktu;" & cJangkaWaktu.Value & ";True"
'      .ParameterFields(2) = "dJatuhTempo;" & Format(dTempo.Value, "dd MMM yyyy") & ";True"
'      .ParameterFields(3) = "nBunga;" & nBunga.Value & ";True"
'      .ParameterFields(4) = "nJumlah;" & Format(nNominal.Value, "###,###,###,###") & ";True"
'      .ParameterFields(5) = "cTerbilang;" & "Terbilang : # " & Dec2Text(nNominal.Value) & " Rupiah #" & ";True"
'      .ParameterFields(6) = "cNama;" & cNama.Text & ";True"
'      .ParameterFields(7) = "cAlamat;" & cAlamat.Text & ";True"
'      .ParameterFields(8) = "cKota;" & cKota.Text & ";True"
'      .ParameterFields(9) = "cKSP1;" & cPimpinan.Text & ";True"
'      .ParameterFields(10) = "cKSP2;" & cNamakasir.Text & ";True"
'      .ParameterFields(11) = "cNo;" & cNoBilyet.Text & ";True"
'      .ParameterFields(12) = "dTgl;" & Format(Date, "dd MMMM yyyy") & ";True"
'      .ParameterFields(13) = "cRekeningSimpanan;" & cRekTabungan.Text & ";True"
'      .ParameterFields(14) = "cJenisSimpanan;" & IIf(optARO(0).Value = True, "Automatic Roll Over", "NON Automatic Roll Over") & ";True"
'      .Action = 1
  End With
End Sub

Private Sub BiSAButton2_Click()

 With CR2
'      .ReportFileName = App.Path & "\rptdepositob.rpt"
'      .ParameterFields(0) = "dTgl;" & Format(Date, "dd MMMM yyyy") & ";True"
'      .ParameterFields(1) = "cNamaDeposan;" & cNama.Text & ";True"
'      .Action = 1
  End With
End Sub

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.Nama,d.Rekening,r.Alamat,d.*", "r.Nama", sisContent, cNama.Text, , "r.Alamat", _
                              Array("Left Join RegisterNasabah r on r.Kode = d.Kode"))
  cAlamat.Text = cAlamat.Browse(dbData)
  If Not dbData.eof Then
    GetRegister
  End If
End Sub

Private Sub GetRegister()
  cGolongan.Text = Mid(GetNull(dbData!Rekening, ""), 4, 2)
  cUrut.Text = Mid(GetNull(dbData!Rekening, ""), 7, 6)
  cFrekuensi.Text = Right(GetNull(dbData!Rekening, ""), 2)
  cNama.Text = GetNull(dbData!nama, "")
  cAlamat.Text = GetNull(dbData!alamat, "")
  If GetNull(dbData!Nobilyet, "") <> "0" Then
    MsgBox "Sudah Mempunyai Nomor Bilyet.", vbInformation
'    cmdSimpan.Enabled = False
    cNoBilyet.Text = GetNull(dbData!Nobilyet, "")
'    cmdPrint.SetFocus
  Else
    cNoBilyet.Text = GetBilyetDeposito(objData, cCabang.Text)
'    cmdSimpan.Enabled = True
  End If
  GetData
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganDeposito", "Kode", cGolongan, "Kode,Keterangan,Lama,Bunga")
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cmdPrint_Click()
Dim cRekening As String
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  If MsgBox("Akan mencetak Bilyet Deposito ?", vbYesNo + vbInformation) = vbYes Then
      CetakBilyet cRekening, cNama.Text, cAlamat.Text, Dec2Text(cJangkaWaktu.Value) & " Bulan", dTempo.Value, _
                       nBunga.Value, dTgl.Value, cNoBilyet.Text, nNominal.Value, cNamakasir.Text, cPimpinan.Text, dTanggal.Value, objData
  Else
    initvalue
    cGolongan.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim cRekening As String
Dim cTerbilangNominal As String
Dim cTerbilangBunga As String
Dim cKota As String
       
  If ValidSaving() Then
    If MsgBox("Apakah Data Benar-benar sudah Valid ?", vbYesNo + vbInformation) = vbYes Then
       
       cTerbilangNominal = Dec2Text(nNominal.Value)
       cTerbilangBunga = Dec2Text(nBunga.Value)
       cKota = aCfg(msKota)
       cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
       
        'Update di Deposito
        objData.Edit GetDSN, "DEPOSITO", "Rekening = '" & cRekening & "'", Array("NoBilyet", "NoSeri"), Array(cNoBilyet.Text, cSeri.Text)
        
        'Add di NomorBilyet
        vaField = Array("RekeningDeposito", "Tanggal", "NomorBilyet", "Keterangan", "Status", "NoSeri")
        vaValue = Array(cRekening, dTgl.Value, cNoBilyet.Text, "", "0", cSeri.Text)
        objData.Add GetDSN, "NomorBilyet", vaField, vaValue
        
        If MsgBox("Akan mencetak Bilyet Deposito ?", vbYesNo + vbInformation) = vbYes Then
           CetakBilyet cRekening, cNama.Text, cAlamat.Text, Dec2Text(cJangkaWaktu.Value) & " Bulan", dTempo.Value, _
                       nBunga.Value, dTgl.Value, cNoBilyet.Text, nNominal.Value, cNamakasir.Text, cPimpinan.Text, objData
        End If
       
       
    End If
  End If
    initvalue
    cGolongan.SetFocus
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cGolongan.Text, "Golongan Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cUrut.Text, "Nomor Urut Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cUrut.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Frekuensi Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cFrekuensi.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cNoBilyet.Text, "Nomor Bilyet Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cNoBilyet.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cSeri.Text, "Nomor Seri Bilyet Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cSeri.SetFocus
    Exit Function
  End If
End Function

Private Sub cFrekuensi_Validate(Cancel As Boolean)
  If cFrekuensi.LastKey = 13 Or cFrekuensi.LastKey = 40 Then
    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
      Set dbData = objData.Browse(GetDSN, "Deposito", "Rekening,NominalDeposito,noBilyet", "Rekening", sisAssign, cRekening, "And Status <>'1'")
      If Not dbData.eof Then
        If GetNull(dbData!Nobilyet, "") <> "" Then
          If MsgBox("Rekening ini sudah mempunyai nomor bilyet. Apakah hanya mencetak Nomor Bilyet ?", vbYesNo + vbInformation) = vbYes Then
'            cmdSimpan.Enabled = False
            GetData
            cNoBilyet.Enabled = False
            cSeri.Enabled = False
'            cmdPrint.Enabled = True
'            cmdPrint.SetFocus
            Exit Sub
          Else
            Cancel = True
            initvalue
            cCabang.SetFocus
            Exit Sub
          End If
        End If
        cNoBilyet.Enabled = True
        cSeri.Enabled = True
'        cmdSimpan.Enabled = True
        GetData
        cNoBilyet.Text = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text) 'GetBilyetDeposito(objData, cCabang.Text)
      Else
         MsgBox "Data Tidak Ada.", vbOKOnly + vbInformation, "Cetak Bilyet Deposito"
         Cancel = True
         initvalue
         cGolongan.SetFocus
         Exit Sub
      End If
  End If
End Sub
 
Private Sub GetData()
Dim cFields As String
Dim vaJoin
Dim cRekening As String
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cFields = "d.Lama as lamaDeposito,d.Rekening,d.Nominaldeposito,d.GolonganDeposito,d.Tgl,d.jthtmp,d.Status,d.SistemARO,d.SukuBunga,d.NoBilyet,d.NoSeri,"
  cFields = cFields & " r.Nama,r.Alamat,r.Telepon,r.Path,w.keterangan as keteranganwilayah,d.rekeningsimpanan,"
  cFields = cFields & " b.Keterangan as KeteranganGolDeposito,b.Lama"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode = d.Kode", _
               "Left Join GolonganDeposito b on b.Kode=d.GolonganDeposito", _
               "left join wilayah w on w.kode = r.wilayah")
  Set dbData = objData.Browse(GetDSN, "Deposito d", cFields, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
      cNama.Text = GetNull(dbData!nama, "")
      cAlamat.Text = GetNull(dbData!alamat, "")
      dTgl.Value = GetNull(dbData!Tgl, "")
      dTempo.Value = GetNull(dbData!jthtmp, "")
      nNominal.Value = GetNull(dbData!nominaldeposito)
      cGolonganDeposito.Text = GetNull(dbData!GolonganDeposito, "")
      cKetGolDeposito.Text = GetNull(dbData!KeteranganGolDeposito, "")
      cJangkaWaktu.Value = GetNull(dbData!LamaDeposito)
      nNominal.Value = GetNull(dbData!nominaldeposito)
      nBunga.Value = GetNull(dbData!SukuBunga)
      cNoBilyet.Text = GetNull(dbData!Nobilyet, "")
      cSeri.Text = GetNull(dbData!NoSeri, "")
      SetOpt optARO, GetNull(dbData!SistemARO, "")
      cKota.Text = GetNull(dbData!keteranganwilayah)
      cRekTabungan.Text = GetNull(dbData!rekeningsimpanan)
  End If
End Sub
  
Private Sub initvalue()
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNoBilyet.Default
  dTgl.Value = Date
  cNama.Default
  cAlamat.Default
  cGolonganDeposito.Default
  cKetGolDeposito.Default
  cJangkaWaktu.Default
  dTempo.Value = Date
  nNominal.Value = 0
  nBunga.Value = 0
  cSeri.Default
  cNamakasir.Default
  cPimpinan.Text = aCfg(msNamaDirut)
  cNamakasir.Text = GetRegistry(reg_FullName)
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.Nama,d.Rekening,r.Alamat,d.*", "r.Nama", sisContent, cNama.Text, "And Status <>'1'", "r.Nama", _
                              Array("Left Join RegisterNasabah r on r.Kode = d.Kode"))
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.eof Then
    GetRegister
  End If
End Sub

Private Sub cSeri_Validate(Cancel As Boolean)
  If cSeri.LastKey = 13 Then
    Set dbData = objData.Browse(GetDSN, "NomorBilyet", "NoSeri", "NoSeri", sisAssign, cSeri.Text)
    If Not dbData.eof Then
      MsgBox "Nomor Bilyet sudah ada pernah di entrikan.", vbInformation
      Cancel = True
      cSeri.SetFocus
    End If
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  initvalue
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cNoBilyet, n
  TabIndex cSeri, n
  TabIndex cNamakasir, n
  TabIndex cPimpinan, n
  TabIndex BiSAButton1, n
  TabIndex BiSAButton2, n
'  TabIndex cmdSimpan, n
'  TabIndex cmdPrint, n
  TabIndex cmdKeluar, n
End Sub

Private Sub Tanggal()
Dim Keterangan, i
Dim n As Single

     Keterangan = Array("Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember")
     i = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
     For n = 0 To 11
        If i(n) = Month(Date) Then
          cBulan = Day(Date) & "  " & Keterangan(n)
          Exit For
        End If
     Next
End Sub

Private Sub CetakBilyet(ByVal cRekening As String, ByVal cNama As String, ByVal cAlamat As String, _
                ByVal cJangkaWaktu As String, ByVal dJatuhTempo As Date, _
                nSukuBunga As Double, ByVal dTglValuta As Date, ByVal cNoBilyet As String, ByVal nNominal As Double, _
                ByVal cKasir As String, ByVal cDirut As String, ByVal dTglCetak As Date, Optional ByVal obj As CodeSuiteLibrary.data)
Dim cNamaDir As String
Dim n As Integer
Dim vaField, vaValue
  
  Set dbData = objData.SQL(GetDSN, "Select * From SetupBilyet")
  With dbData
    frmTDBR1.SetMargin GetNull(!Tinggi), GetNull(!Lebar)
    frmTDBR1.AddPoint cRekening, GetNull(!xRekening), GetNull(!yRekening), GetNull(!wRekening), n
    frmTDBR1.AddPoint cNama, GetNull(!xNama), GetNull(!yNama), GetNull(!wNama), n
    frmTDBR1.AddPoint cAlamat, GetNull(!xAlamat), GetNull(!yAlamat), GetNull(!wAlamat), n
    frmTDBR1.AddPoint SisFormat(nNominal, Sis_BilRpPict2), GetNull(!XJumlah), GetNull(!YJumlah), GetNull(!WJumlah), n
    frmTDBR1.AddPoint Dec2Text(nNominal) & "Rupiah", GetNull(!xterbilang), GetNull(!yterbilang), GetNull(!wterbilang), n
    frmTDBR1.AddPoint cJangkaWaktu, GetNull(!xLama), GetNull(!yLama), GetNull(!wlama), n
    frmTDBR1.AddPoint dTglValuta, GetNull(!xvaluta), GetNull(!yValuta), GetNull(!wvaluta), n
    frmTDBR1.AddPoint dJatuhTempo, GetNull(!xTempo), GetNull(!yTempo), GetNull(!wTempo), n
    frmTDBR1.AddPoint nSukuBunga, GetNull(!xBunga), GetNull(!yBunga), GetNull(!wBunga), n
    frmTDBR1.AddPoint Dec2Text(nSukuBunga), GetNull(!xTerbilangSB), GetNull(!yTerbilangSB), GetNull(!wTerbilangSB), n
    frmTDBR1.AddPoint dTglCetak, GetNull(!xTglCetak), GetNull(!yTglCetak), GetNull(!wTglCetak), n
    frmTDBR1.AddPoint cKasir, GetNull(!xkasir), GetNull(!ykasir), GetNull(!wKasir), n
    frmTDBR1.AddPoint cDirut, GetNull(!XDirut), GetNull(!YDirut), GetNull(!WDirut), n
    frmTDBR1.PrintPreview
  End With
End Sub
