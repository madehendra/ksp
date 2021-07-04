VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form TrPengajuanKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PENGAJUAN PINJAMAN"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8055
   Begin TabDlg.SSTab SSTab1 
      Height          =   4650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   8202
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Nama"
      TabPicture(0)   =   "TrPengajuanKredit.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "BiSAFrame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Alamat"
      TabPicture(1)   =   "TrPengajuanKredit.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BiSAFrame5"
      Tab(1).Control(1)=   "BiSAFrame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Analisa Keuangan"
      TabPicture(2)   =   "TrPengajuanKredit.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "BiSAFrame3"
      Tab(2).ControlCount=   1
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   1725
         Left            =   -74895
         Top             =   2670
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   3043
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
         Begin BiSATextBoxProject.BiSATextBox cJaminan 
            Height          =   330
            Left            =   180
            TabIndex        =   1
            Top             =   1230
            Width           =   7170
            _ExtentX        =   12647
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
            Caption         =   "Jenis Jaminan"
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
         Begin BiSANumberBoxProject.BiSANumberBox nPlafond 
            Height          =   330
            Left            =   165
            TabIndex        =   2
            Top             =   510
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
            Caption         =   "Jumlah Pengajuan"
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
         Begin BiSATextBoxProject.BiSATextBox cNamaAO 
            Height          =   330
            Left            =   3345
            TabIndex        =   3
            Top             =   150
            Width           =   3150
            _ExtentX        =   5556
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
         Begin BiSATextBoxProject.BiSABrowse cAO 
            Height          =   330
            Left            =   165
            TabIndex        =   4
            Top             =   150
            Width           =   3180
            _ExtentX        =   5609
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
            Button          =   -1  'True
            Caption         =   "Account Officer"
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
         Begin BiSANumberBoxProject.BiSANumberBox nLama 
            Height          =   330
            Left            =   165
            TabIndex        =   5
            Top             =   870
            Width           =   2940
            _ExtentX        =   5186
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
            Caption         =   "Lama"
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
         Begin VB.Label Label4 
            Caption         =   "Bulan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3210
            TabIndex        =   6
            Top             =   930
            Width           =   825
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   2145
         Left            =   -74895
         Top             =   525
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   3784
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
         Begin BiSATextBoxProject.BiSATextBox cAlamatRumah 
            Height          =   330
            Left            =   630
            TabIndex        =   7
            Top             =   105
            Width           =   6360
            _ExtentX        =   11218
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
            Caption         =   "Alamat Rumah"
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
         Begin BiSATextBoxProject.BiSATextBox cTeleponRumah 
            Height          =   330
            Left            =   630
            TabIndex        =   8
            Top             =   465
            Width           =   4740
            _ExtentX        =   8361
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
            Caption         =   "Telepon"
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
         Begin BiSATextBoxProject.BiSATextBox cAlamatKantor 
            Height          =   330
            Left            =   630
            TabIndex        =   9
            Top             =   945
            Width           =   6360
            _ExtentX        =   11218
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
            Caption         =   "Alamat kantor"
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
         Begin BiSATextBoxProject.BiSATextBox cTeleponKantor 
            Height          =   330
            Left            =   630
            TabIndex        =   10
            Top             =   1305
            Width           =   4740
            _ExtentX        =   8361
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
            Caption         =   "Telepon kantor"
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
         Begin BiSATextBoxProject.BiSATextBox cFaxKantor 
            Height          =   330
            Left            =   630
            TabIndex        =   11
            Top             =   1665
            Width           =   4740
            _ExtentX        =   8361
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
            Caption         =   "fax kantor"
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   4035
         Left            =   120
         Top             =   450
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   7117
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
         Begin VB.OptionButton optStatusKawin 
            Caption         =   "&KAWIN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2250
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2460
            Width           =   900
         End
         Begin VB.OptionButton optStatusKawin 
            Caption         =   "&BELUM"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3615
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2460
            Width           =   1065
         End
         Begin BiSATextBoxProject.BiSATextBox cKode 
            Height          =   330
            Left            =   2715
            TabIndex        =   12
            Top             =   495
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
            Text            =   "1234567"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   7
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
            Left            =   135
            TabIndex        =   13
            Top             =   495
            Width           =   2565
            _ExtentX        =   4524
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
            Caption         =   "No Pengajuan"
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame10 
            Height          =   375
            Left            =   2250
            Top             =   1245
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   661
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
            Begin VB.OptionButton optSex 
               Caption         =   "&PEREMPUAN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   1530
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   75
               Width           =   1455
            End
            Begin VB.OptionButton optSex 
               Caption         =   "&LAKI-LAKI"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   90
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   60
               Width           =   1305
            End
         End
         Begin BiSATextBoxProject.BiSABrowse cNama 
            Height          =   330
            Left            =   135
            TabIndex        =   18
            Top             =   870
            Width           =   6210
            _ExtentX        =   10954
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
            Button          =   -1  'True
            Caption         =   "Nama lengkap"
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
         Begin BiSADateProject.BiSADate dTglRegister 
            Height          =   330
            Left            =   135
            TabIndex        =   19
            Top             =   120
            Width           =   3480
            _ExtentX        =   6138
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
            Caption         =   "Tgl Pengajuan"
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
         Begin BiSATextBoxProject.BiSATextBox cTempatLahir 
            Height          =   330
            Left            =   135
            TabIndex        =   20
            Top             =   1665
            Width           =   5340
            _ExtentX        =   9419
            _ExtentY        =   582
            Text            =   "12"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Tempat Lahir"
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
         Begin BiSADateProject.BiSADate dTglLahir 
            Height          =   330
            Left            =   135
            TabIndex        =   21
            Top             =   2055
            Width           =   3480
            _ExtentX        =   6138
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
            Caption         =   "Tanggal lahir"
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
         Begin BiSATextBoxProject.BiSATextBox cKTP 
            Height          =   330
            Left            =   135
            TabIndex        =   22
            Top             =   2790
            Width           =   5340
            _ExtentX        =   9419
            _ExtentY        =   582
            Text            =   "12"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   20
            Caption         =   "No KTP/SIM"
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
         Begin BiSATextBoxProject.BiSATextBox cNamaPekerjaan 
            Height          =   330
            Left            =   3315
            TabIndex        =   23
            Top             =   3180
            Width           =   2760
            _ExtentX        =   4868
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
         Begin BiSATextBoxProject.BiSABrowse cPekerjaan 
            Height          =   330
            Left            =   135
            TabIndex        =   24
            Top             =   3180
            Width           =   3180
            _ExtentX        =   5609
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
            Button          =   -1  'True
            Caption         =   "Pekerjaan"
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
         Begin BiSATextBoxProject.BiSATextBox cNamaWilayah 
            Height          =   330
            Left            =   3315
            TabIndex        =   33
            Top             =   3555
            Width           =   2760
            _ExtentX        =   4868
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
         Begin BiSATextBoxProject.BiSABrowse cWilayah 
            Height          =   330
            Left            =   135
            TabIndex        =   34
            Top             =   3555
            Width           =   3180
            _ExtentX        =   5609
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
            Button          =   -1  'True
            Caption         =   "Wilayah"
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
         Begin VB.Label Label3 
            Caption         =   "Status Perkawinan"
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
            Left            =   150
            TabIndex        =   26
            Top             =   2460
            Width           =   2010
         End
         Begin VB.Label Label1 
            Caption         =   "Jenis kelamin"
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
            Left            =   150
            TabIndex        =   25
            Top             =   1335
            Width           =   1650
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame3 
         Height          =   4245
         Left            =   -74970
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7488
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
         Begin BiSANumberBoxProject.BiSANumberBox nBiayaRT 
            Height          =   330
            Left            =   4035
            TabIndex        =   35
            Top             =   855
            Width           =   3810
            _ExtentX        =   6720
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
            Caption         =   "Biaya Rumah Tangga"
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
         Begin BiSANumberBoxProject.BiSANumberBox nBiayaTK 
            Height          =   330
            Left            =   4035
            TabIndex        =   36
            Top             =   1215
            Width           =   3810
            _ExtentX        =   6720
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
            Caption         =   "Biaya Telepon"
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
         Begin BiSANumberBoxProject.BiSANumberBox nBiayaListrik 
            Height          =   330
            Left            =   4035
            TabIndex        =   37
            Top             =   1575
            Width           =   3810
            _ExtentX        =   6720
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
            Caption         =   "Biaya Listrik/Air"
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
         Begin BiSANumberBoxProject.BiSANumberBox nBiayaPemeliharaan 
            Height          =   330
            Left            =   4035
            TabIndex        =   38
            Top             =   1920
            Width           =   3810
            _ExtentX        =   6720
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
            Caption         =   "Biaya Pemeliharaan"
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
         Begin BiSANumberBoxProject.BiSANumberBox nBiayaLain 
            Height          =   330
            Left            =   4035
            TabIndex        =   39
            Top             =   2280
            Width           =   3810
            _ExtentX        =   6720
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
            Caption         =   "Biaya Lain - Lain"
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
         Begin BiSANumberBoxProject.BiSANumberBox nPendapatanUtama 
            Height          =   345
            Left            =   75
            TabIndex        =   40
            Top             =   120
            Width           =   3915
            _ExtentX        =   6906
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
            Caption         =   "Pendapatan Utama"
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
         Begin BiSANumberBoxProject.BiSANumberBox nPendapatanLain 
            Height          =   345
            Left            =   75
            TabIndex        =   41
            Top             =   495
            Width           =   3915
            _ExtentX        =   6906
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
            Caption         =   "Pendapatan Lain Lain"
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
         Begin BiSANumberBoxProject.BiSANumberBox nJumlahPendapatan 
            Height          =   345
            Left            =   45
            TabIndex        =   42
            Top             =   2775
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   609
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
            Caption         =   "Jumlah"
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
         Begin BiSANumberBoxProject.BiSANumberBox nJumlahBiaya 
            Height          =   345
            Left            =   4035
            TabIndex        =   43
            Top             =   2775
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   609
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
            Caption         =   "Jumlah"
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
         Begin VB.Line Line2 
            X1              =   5895
            X2              =   7785
            Y1              =   2685
            Y2              =   2685
         End
         Begin VB.Line Line1 
            X1              =   2145
            X2              =   4020
            Y1              =   2670
            Y2              =   2670
         End
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   630
      Left            =   0
      Top             =   4650
      Width           =   8010
      _ExtentX        =   14129
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
         Left            =   2205
         TabIndex        =   27
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
         Picture         =   "TrPengajuanKredit.frx":0054
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3375
         TabIndex        =   28
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
         Picture         =   "TrPengajuanKredit.frx":02DE
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   5790
         TabIndex        =   29
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
         Picture         =   "TrPengajuanKredit.frx":047D
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1155
         TabIndex        =   30
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
         Picture         =   "TrPengajuanKredit.frx":0893
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   90
         TabIndex        =   31
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
         Picture         =   "TrPengajuanKredit.frx":09BF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6870
         TabIndex        =   32
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
         Picture         =   "TrPengajuanKredit.frx":0B6A
      End
   End
End
Attribute VB_Name = "TrPengajuanKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim lEdit As Boolean
Dim nPos As SisPos
Dim cSQL As String

Private Sub cAlamatRumah_LostFocus()
  If cAlamatRumah.LastKey = 16 Then 'Panah atas
    SSTab1.Tab = 0
  End If
End Sub

Private Sub cAO_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "AO", "Kode,Nama", "Kode", sisContent, cAO.Text)
  cAO.Text = cAO.Browse(dbData)
  If Not dbData.eof Then
    cNamaAO.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub GetKode()
  If Trim(cKode.Text) = "" Then
      Set dbData = objData.Browse(GetDSN, "PengajuanKredit", "Max(Kode) Kode", "Kode", sisPrefix, cCabang.Text)
      cKode.Text = "1"
      If Not dbData.eof Then
        cKode.Text = Val(Mid(GetNull(dbData!Kode), Len(cCabang.Text) + 2)) + 1
      End If
    End If
  cKode.Text = Padl(Trim(cKode.Text), 7, "0")
End Sub

Private Sub cJaminan_Validate(Cancel As Boolean)
  SSTab1.Tab = 2
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
  GetKode
  Set dbData = objData.Browse(GetDSN, "PengajuanKredit", "Kode", "Kode", sisAssign, cCabang.Text & "." & cKode.Text)
  If Not dbData.eof Then
    If nPos = Add Then
      MsgBox "Nomor Pengajuan Sudah Ada. Silahkan ulangi pengisian", vbInformation, "Pengajuan Pembiayaan"
      Cancel = True
      cKode.Default
      cKode.SetFocus
      Exit Sub
    End If
    GetMemory
    If nPos = Delete Then DeleteData
  ElseIf dbData.eof And nPos <> Add Then
    MsgBox "Data Tidak Ada. Silahkan ulangi pengisian !", vbInformation
    Cancel = True
    cKode.Default
    cKode.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  cKode.Enabled = True
  cNama.Button = False
  dTglRegister.SetFocus
End Sub

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  SSTab1.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  initvalue
  cNama.Button = True
  cCabang.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  initvalue
  cNama.Button = True
  cCabang.SetFocus
End Sub

Private Sub DeleteData()
If MsgBox("Data Benar-benar Dihapus ?", vbYesNo + vbExclamation) = vbYes Then
    objData.Delete GetDSN, "PengajuanKredit", "Kode", sisAssign, cCabang.Text & "." & cKode.Text
  End If
  GetEdit False
  initvalue
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    initvalue
    GetEdit False
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim cNoPengaJuan As String

  cNoPengaJuan = cCabang.Text & "." & cKode.Text
  If ValidSaving() Then
    If MsgBox("Data benar-benar sudah valid ?", vbYesNo + vbInformation) = vbYes Then
      vaField = Array("Kode", "Nama", "Kelamin", _
                      "TempatLahir", "TglLahir", "StatusPerkawinan", _
                      "KTP", "Pekerjaan", "Wilayah", "Alamat", "Telepon", _
                      "AlamatKantor", "TeleponKantor", "FaxKantor", "AO", "Plafond", "Lama", "Jaminan", "TglRegister", _
                      "nPendapatanUtama", "nPendapatanLain", _
                      "nBiayaRT", "nBiayaTK", "nBiayaListrik", "nBiayaDanaSosial", "nBiayaAdministrasi", "nBiayaUmum", "nBiayaPenyusutan", "nBiayaPemeliharaan", "nBiayaLain")
      vaValue = Array(cNoPengaJuan, cNama.Text, GetOpt(optSex), _
                      cTempatLahir.Text, dTglLahir.Value, GetOpt(optStatusKawin), _
                      cKTP.Text, cPekerjaan.Text, cWilayah.Text, cAlamatRumah.Text, cTeleponRumah.Text, _
                      cAlamatKantor.Text, cTeleponKantor.Text, cFaxKantor.Text, cAO.Text, nPlafond.Value, nLama.Value, cJaminan.Text, dTglRegister.Value, _
                      nPendapatanUtama.Value, nPendapatanLain.Value, _
                      nBiayaRT.Value, nBiayaTK.Value, nBiayaListrik.Value, 0, 0, 0, 0, nBiayaPemeliharaan.Value, nBiayaLain.Value)
      objData.Update GetDSN, "PengajuanKredit", "Kode = '" & cNoPengaJuan & "'", vaField, vaValue
    Else
      cNama.SetFocus
      Exit Sub
    End If
    initvalue
    GetEdit False
  End If
End Sub

Static Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cKode.Text, "Kode Register Nasabah Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 0
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cNama.Text, "Nama Register Nasabah Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 0
    cNama.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cTempatLahir.Text, "Tempat Lahir Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 0
    cTempatLahir.SetFocus
    Exit Function
  End If
   
  If Not CheckData(cKTP.Text, "No. Identitas Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 0
    cKTP.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cPekerjaan.Text, "Pekerjaan Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 0
    cPekerjaan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cWilayah.Text, "Wilayah Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 0
    cWilayah.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cAlamatRumah.Text, "Alamat Rumah Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 1
    cAlamatRumah.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cAO.Text, "Account Officer Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 1
    cAO.SetFocus
    Exit Function
  End If
  
  If Not CheckData(nPlafond.Value, "Jumlah Pengajuan Plafond tidak boleh 0, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 1
    nPlafond.SetFocus
    Exit Function
  End If
  
  If Not CheckData(nLama.Value, "Lama tidak boleh 0, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 1
    nLama.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cJaminan.Text, "Jaminan harus diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    SSTab1.Tab = 1
    cJaminan.SetFocus
    Exit Function
  End If
End Function

Private Sub cNama_ButtonClick()
  If nPos = Add Then
    optSex(0).SetFocus
    Exit Sub
  Else
    Set dbData = objData.Browse(GetDSN, "PengajuanKredit", "Nama,Alamat,Kode", "Nama", sisContent, cNama.Text, , "Nama")
    cNama.Text = cNama.Browse(dbData)
    If Not dbData.eof Then
      cKode.Text = Mid(GetNull(dbData!Kode, ""), 4)
      GetMemory
    End If
  End If
End Sub

Private Sub cPekerjaan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Pekerjaan", "Kode,Keterangan", "Kode", sisContent, cPekerjaan.Text)
  cPekerjaan.Text = cPekerjaan.Browse(dbData)
  If Not dbData.eof Then
    cNamaPekerjaan.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cWilayah_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Wilayah", "Kode,Keterangan", "Kode", sisContent, cWilayah.Text)
  cWilayah.Text = cWilayah.Browse(dbData)
  If Not dbData.eof Then
    cNamaWilayah.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub GetMemory()
Dim cFields As String
Dim vaJoin
Dim cKodePengajuan As String

  cKodePengajuan = cCabang.Text & "." & cKode.Text
  cFields = "r.*,p.Keterangan as NamaPekerjaan,a.Nama as NamaAO,w.Keterangan as NamaWilayah"
  vaJoin = Array("Left Join AO a on a.Kode = r.Ao", _
                 "Left Join Pekerjaan p on r.Pekerjaan=p.Kode", _
                 "left Join Wilayah w on w.Kode = r.Wilayah")
  Set dbData = objData.Browse(GetDSN, "PengajuanKredit r", cFields, "r.Kode", sisAssign, cKodePengajuan, , , vaJoin)
  If Not dbData.eof Then
    cKode.Text = Mid(GetNull(dbData!Kode, ""), 4)
    dTglRegister.Value = GetNull(dbData!TglRegister, "")
    cNama.Text = GetNull(dbData!nama, "")
    SetOpt optSex, GetNull(dbData!Kelamin, "")
    SetOpt optStatusKawin, GetNull(dbData!StatusPerkawinan, "")
    cTempatLahir.Text = GetNull(dbData!TempatLahir, "")
    dTglLahir.Value = GetNull(dbData!TglLahir, "")
    cKTP.Text = GetNull(dbData!KTP, "")
    cPekerjaan.Text = GetNull(dbData!Pekerjaan, "")
    cNamaPekerjaan.Text = GetNull(dbData!NamaPekerjaan, "")
    cWilayah.Text = GetNull(dbData!Wilayah, "")
    cNamaWilayah.Text = GetNull(dbData!Namawilayah, "")
    cAlamatRumah.Text = GetNull(dbData!alamat, "")
    cTeleponRumah.Text = GetNull(dbData!Telepon, "")
    
    cAlamatKantor.Text = GetNull(dbData!AlamatKantor, "")
    cTeleponKantor.Text = GetNull(dbData!TeleponKantor, "")
    cFaxKantor.Text = GetNull(dbData!FaxKantor, "")
    cAO.Text = GetNull(dbData!AO, "")
    cNamaAO.Text = GetNull(dbData!namaao, "")
    nPlafond.Value = GetNull(dbData!plafond)
    nLama.Value = GetNull(dbData!Lama)
    cJaminan.Text = GetNull(dbData!Jaminan, "")
    'data analisa keuangan
    nBiayaRT.Value = GetNull(dbData!nBiayaRT)
    nBiayaTK.Value = GetNull(dbData!nBiayaTK)
    nBiayaListrik.Value = GetNull(dbData!nBiayaListrik)
    nBiayaPemeliharaan.Value = GetNull(dbData!nBiayaPemeliharaan)
    nBiayaLain.Value = GetNull(dbData!nBiayaLain)
    nPendapatanLain.Value = GetNull(dbData!nPendapatanLain)
    nPendapatanUtama.Value = GetNull(dbData!nPendapatanUtama)
    SUMJUMLAH
  End If
End Sub

Private Sub initvalue()
  dTglRegister.Value = Date
  cNama.Default
  cKode.Default
  cTempatLahir.Default
  dTglLahir.Value = Date
  cKTP.Default
  cPekerjaan.Default
  cNamaPekerjaan.Default
  cWilayah.Default
  cNamaWilayah.Default
  cAlamatRumah.Default
  cTeleponRumah.Default
  cAlamatKantor.Default
  cTeleponKantor.Default
  cFaxKantor.Default
  cAO.Default
  cNamaAO.Default
  nPlafond.Value = 0
  nLama.Value = 0
  cJaminan.Default
  optSex(0).Value = True
  optStatusKawin(0).Value = True
  SSTab1.Tab = 0
  'ANALISA KEUANGAN
  nPendapatanUtama.Default
  nPendapatanLain.Default
  nJumlahPendapatan.Default
  nJumlahBiaya.Default
  nBiayaRT.Default
  nBiayaTK.Default
  nBiayaListrik.Default
  nBiayaPemeliharaan.Default
  nBiayaLain.Default
End Sub

Private Sub cWilayah_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SSTab1.Tab = 1
    cAlamatRumah.SetFocus
  End If
End Sub

Private Sub dTglRegister_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTglRegister.Value) Then
    Cancel = True
    dTglRegister.SetFocus
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  initvalue
  GetEdit False
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex dTglRegister, n
  TabIndex cCabang, n
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex optSex(0), n
  TabIndex optSex(1), n
  TabIndex cTempatLahir, n
  TabIndex dTglLahir, n
  TabIndex optStatusKawin(0), n
  TabIndex optStatusKawin(1), n
  TabIndex cKTP, n
  TabIndex cPekerjaan, n
  TabIndex cWilayah, n
  TabIndex cAlamatRumah, n
  TabIndex cTeleponRumah, n
  TabIndex cAlamatKantor, n
  TabIndex cTeleponKantor, n
  TabIndex cFaxKantor, n
  TabIndex cAO, n
  TabIndex nPlafond, n
  TabIndex nLama, n
  TabIndex cJaminan, n
  
  TabIndex nPendapatanUtama, n
  TabIndex nPendapatanLain, n
  TabIndex nBiayaRT, n
  TabIndex nBiayaTK, n
  TabIndex nBiayaListrik, n
  TabIndex nBiayaPemeliharaan, n
  TabIndex nBiayaLain, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub nBiayaLain_Validate(Cancel As Boolean)
  SUMJUMLAH
End Sub

Private Sub nBiayaListrik_Validate(Cancel As Boolean)
  SUMJUMLAH
End Sub

Private Sub nBiayaPemeliharaan_Validate(Cancel As Boolean)
  SUMJUMLAH
End Sub

Private Sub nBiayaRT_Validate(Cancel As Boolean)
  SUMJUMLAH
End Sub

Private Sub nBiayaTK_Validate(Cancel As Boolean)
  SUMJUMLAH
End Sub

Private Sub nPendapatanLain_Validate(Cancel As Boolean)
  SUMJUMLAH
End Sub

Private Sub nPendapatanUtama_Validate(Cancel As Boolean)
  SUMJUMLAH
End Sub

Private Sub SUMJUMLAH()
  nJumlahPendapatan.Value = nPendapatanLain.Value + nPendapatanUtama.Value
  nJumlahBiaya.Value = nBiayaRT.Value + _
                       nBiayaTK.Value + _
                       nBiayaListrik.Value + _
                       nBiayaPemeliharaan.Value + _
                       nBiayaLain.Value
End Sub

Private Sub optSex_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optStatuskawin_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub
