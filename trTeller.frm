VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form trTeller 
   BorderStyle     =   0  'None
   Caption         =   "Teller"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4680
      Left            =   15
      TabIndex        =   6
      Top             =   1860
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   8255
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "MUTASI TABUNGAN"
      TabPicture(0)   =   "trTeller.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameMutasiTabungan"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "DEPOSITO"
      TabPicture(1)   =   "trTeller.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameDeposito"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PENCAIRAN PINJAMAN"
      TabPicture(2)   =   "trTeller.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FramePencairan"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "ANGSURAN"
      TabPicture(3)   =   "trTeller.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "SALDO TELLER"
      TabPicture(4)   =   "trTeller.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "BiSAFrame6"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   4200
         Left            =   -74910
         TabIndex        =   64
         Top             =   405
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   7408
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Data Angsuran"
         TabPicture(0)   =   "trTeller.frx":008C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FrameAngsuran"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Kartu Angsuran"
         TabPicture(1)   =   "trTeller.frx":00A8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TDBGrid2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Periode Pembayaran"
         TabPicture(2)   =   "trTeller.frx":00C4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TDBGrid3"
         Tab(2).ControlCount=   1
         Begin BiSAFramProject.BiSAFrame FrameAngsuran 
            Height          =   3780
            Left            =   75
            Top             =   390
            Width           =   10950
            _ExtentX        =   19315
            _ExtentY        =   6668
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
            Begin BiSATextBoxProject.BiSATextBox cCaraAngsuran 
               Height          =   315
               Left            =   4035
               TabIndex        =   65
               Top             =   480
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   556
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
            Begin BiSADateProject.BiSADate dTglRealisasiAngsuran 
               Height          =   330
               Left            =   120
               TabIndex        =   66
               Top             =   480
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   582
               Value           =   "13-10-2005"
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
               Caption         =   "Tgl Realisasi"
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
            Begin BiSATextBoxProject.BiSATextBox cSpkAngsuran 
               Height          =   330
               Left            =   120
               TabIndex        =   67
               Top             =   135
               Width           =   4320
               _ExtentX        =   7620
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
               Caption         =   "Nomor SPK"
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
            Begin BiSANumberBoxProject.BiSANumberBox nPlafondAngsuran 
               Height          =   330
               Left            =   120
               TabIndex        =   68
               Top             =   825
               Width           =   3855
               _ExtentX        =   6800
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
               Caption         =   "Plafond"
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
            Begin BiSANumberBoxProject.BiSANumberBox nBungaAngsuran 
               Height          =   330
               Left            =   120
               TabIndex        =   69
               Top             =   1170
               Width           =   2430
               _ExtentX        =   4286
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
               Caption         =   "Bunga (%) p.a"
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
            Begin BiSANumberBoxProject.BiSANumberBox nLamaAngsuran 
               Height          =   330
               Left            =   120
               TabIndex        =   70
               Top             =   1515
               Width           =   2430
               _ExtentX        =   4286
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
               Caption         =   "Lama"
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
            Begin BiSANumberBoxProject.BiSANumberBox nAngsuranPokok 
               Height          =   330
               Left            =   5760
               TabIndex        =   71
               Top             =   1020
               Width           =   4275
               _ExtentX        =   7541
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
               Caption         =   "Angsuran Pokok"
               CaptionWidth    =   2100
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
            Begin BiSANumberBoxProject.BiSANumberBox nAngsuranBunga 
               Height          =   330
               Left            =   5760
               TabIndex        =   72
               Top             =   1395
               Width           =   4275
               _ExtentX        =   7541
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
               Caption         =   "Angsuran Bunga"
               CaptionWidth    =   2100
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
            Begin BiSANumberBoxProject.BiSANumberBox nDenda 
               Height          =   330
               Left            =   5760
               TabIndex        =   73
               Top             =   1770
               Width           =   4275
               _ExtentX        =   7541
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
               Caption         =   "Denda"
               CaptionWidth    =   2100
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
            Begin BiSANumberBoxProject.BiSANumberBox nKewajiban 
               Height          =   330
               Left            =   5760
               TabIndex        =   74
               Top             =   2280
               Width           =   4275
               _ExtentX        =   7541
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
               Caption         =   "Total Kewajiban"
               CaptionWidth    =   2100
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
            Begin BiSANumberBoxProject.BiSANumberBox nPeriodeAngsuran 
               Height          =   330
               Left            =   105
               TabIndex        =   75
               Top             =   2925
               Width           =   4020
               _ExtentX        =   7091
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
               Caption         =   "Max Periode Bayar ( Dlm 1 Bulan)."
               CaptionWidth    =   3200
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
            Begin BiSANumberBoxProject.BiSANumberBox nSisaPokok 
               Height          =   330
               Left            =   5760
               TabIndex        =   76
               Top             =   2640
               Width           =   4275
               _ExtentX        =   7541
               _ExtentY        =   582
               Appearance      =   0
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
               BackColor       =   12632256
               ForeColor       =   -2147483635
               Caption         =   "Baki Debet"
               CaptionWidth    =   2100
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
            Begin BiSANumberBoxProject.BiSANumberBox nMinimumPeriode 
               Height          =   330
               Left            =   105
               TabIndex        =   77
               Top             =   2565
               Width           =   4020
               _ExtentX        =   7091
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
               Caption         =   "Min Periode Bayar ( Dlm 1 Bulan)."
               CaptionWidth    =   3200
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
            Begin BiSANumberBoxProject.BiSANumberBox nKonpensasi 
               Height          =   330
               Left            =   105
               TabIndex        =   78
               Top             =   3285
               Width           =   4020
               _ExtentX        =   7091
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
               Caption         =   "Konpensasi Keterlambatan"
               CaptionWidth    =   3200
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
            Begin BiSANumberBoxProject.BiSANumberBox nBungaLalu 
               Height          =   330
               Left            =   5760
               TabIndex        =   79
               Top             =   645
               Width           =   4275
               _ExtentX        =   7541
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
               Caption         =   "Sisa Bunga Bulan Lalu"
               CaptionWidth    =   2100
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
            Begin BiSANumberBoxProject.BiSANumberBox nPokokLalu 
               Height          =   330
               Left            =   5760
               TabIndex        =   80
               Top             =   285
               Width           =   4275
               _ExtentX        =   7541
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
               Caption         =   "Sisa Pokok Bulan Lalu"
               CaptionWidth    =   2100
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
            Begin BiSANumberBoxProject.BiSANumberBox nSimpananWajibPeminjam 
               Height          =   330
               Left            =   120
               TabIndex        =   109
               Top             =   2100
               Width           =   4020
               _ExtentX        =   7091
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
               Caption         =   "Simpanan Wajib"
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
            Begin VB.Label lbCaraPerhitungan 
               Caption         =   "CaraPerhitungan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2610
               TabIndex        =   107
               Top             =   1245
               Width           =   1830
            End
            Begin VB.Label Label12 
               Caption         =   "Late *"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   10095
               TabIndex        =   104
               Top             =   690
               Width           =   600
            End
            Begin VB.Label Label11 
               Caption         =   "Late *"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   10080
               TabIndex        =   103
               Top             =   345
               Width           =   600
            End
            Begin VB.Line Line3 
               X1              =   5820
               X2              =   10095
               Y1              =   2190
               Y2              =   2190
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Left            =   2610
               TabIndex        =   82
               Top             =   1530
               Width           =   540
            End
            Begin VB.Label Label7 
               Caption         =   "Hari"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   4245
               TabIndex        =   81
               Top             =   3345
               Width           =   465
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   2895
            Left            =   -74910
            TabIndex        =   83
            Top             =   570
            Width           =   9570
            _ExtentX        =   16880
            _ExtentY        =   5106
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "No"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tgl Angsuran"
            Columns(1).DataField=   ""
            Columns(1).NumberFormat=   "dd-MM-yyyy"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Angsuran Pokok"
            Columns(2).DataField=   ""
            Columns(2).NumberFormat=   "###,###,###,###,##0.00"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Angsuran Bunga"
            Columns(3).DataField=   ""
            Columns(3).NumberFormat=   "###,###,###,###,##0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Denda"
            Columns(4).DataField=   ""
            Columns(4).NumberFormat=   "###,###,###,###,##0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Total"
            Columns(5).DataField=   ""
            Columns(5).NumberFormat=   "###,###,###,###,##0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1217"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1138"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3572"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3493"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=3969"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3889"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=3572"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3493"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=212"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=132"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(25)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(27)=   "Column(5).Width=4048"
            Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=3969"
            Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            ColumnFooters   =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=172,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFCFCED&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HEBDACB&,.bold=0"
            _StyleDefs(15)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(28)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(29)  =   ":id=14,.fontname=Tahoma"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(63)  =   "Named:id=33:Normal"
            _StyleDefs(64)  =   ":id=33,.parent=0"
            _StyleDefs(65)  =   "Named:id=34:Heading"
            _StyleDefs(66)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   ":id=34,.wraptext=-1"
            _StyleDefs(68)  =   "Named:id=35:Footing"
            _StyleDefs(69)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   "Named:id=36:Selected"
            _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(72)  =   "Named:id=37:Caption"
            _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(74)  =   "Named:id=38:HighlightRow"
            _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(76)  =   "Named:id=39:EvenRow"
            _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(78)  =   "Named:id=40:OddRow"
            _StyleDefs(79)  =   ":id=40,.parent=33"
            _StyleDefs(80)  =   "Named:id=41:RecordSelector"
            _StyleDefs(81)  =   ":id=41,.parent=34"
            _StyleDefs(82)  =   "Named:id=42:FilterBar"
            _StyleDefs(83)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
            Height          =   2895
            Left            =   -74865
            TabIndex        =   84
            Top             =   525
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   5106
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Periode"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tanggal"
            Columns(1).DataField=   ""
            Columns(1).NumberFormat=   "dd-MM-yyyy"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Sampai Dengan"
            Columns(2).DataField=   ""
            Columns(2).NumberFormat=   "dd-MM-yyyy"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2566"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3201"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3122"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=3043"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2963"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   2
            FootLines       =   2
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=172,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFCFCED&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HEBDACB&,.bold=0"
            _StyleDefs(15)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(28)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(29)  =   ":id=14,.fontname=Tahoma"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(51)  =   "Named:id=33:Normal"
            _StyleDefs(52)  =   ":id=33,.parent=0"
            _StyleDefs(53)  =   "Named:id=34:Heading"
            _StyleDefs(54)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   ":id=34,.wraptext=-1"
            _StyleDefs(56)  =   "Named:id=35:Footing"
            _StyleDefs(57)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   "Named:id=36:Selected"
            _StyleDefs(59)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(60)  =   "Named:id=37:Caption"
            _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(62)  =   "Named:id=38:HighlightRow"
            _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(64)  =   "Named:id=39:EvenRow"
            _StyleDefs(65)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(66)  =   "Named:id=40:OddRow"
            _StyleDefs(67)  =   ":id=40,.parent=33"
            _StyleDefs(68)  =   "Named:id=41:RecordSelector"
            _StyleDefs(69)  =   ":id=41,.parent=34"
            _StyleDefs(70)  =   "Named:id=42:FilterBar"
            _StyleDefs(71)  =   ":id=42,.parent=33"
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame6 
         Height          =   3630
         Left            =   75
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6403
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   3045
            Left            =   90
            TabIndex        =   46
            Top             =   90
            Width           =   11145
            _ExtentX        =   19659
            _ExtentY        =   5371
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "No"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "No Transaksi"
            Columns(1).DataField=   "Faktur"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Keterangan"
            Columns(2).DataField=   "Rekening"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "DK"
            Columns(3).DataField=   "Debet"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Jumlah"
            Columns(4).DataField=   "Kredit"
            Columns(4).NumberFormat=   "FormatText Event"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=820"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=4154"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4075"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=9075"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=8996"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=873"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=794"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=4154"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=4075"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   2
            FootLines       =   2
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=1200,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=Times New Roman"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Arial"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=Arial"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=Arial"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(28)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(29)  =   ":id=14,.fontname=Tahoma"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(59)  =   "Named:id=33:Normal"
            _StyleDefs(60)  =   ":id=33,.parent=0"
            _StyleDefs(61)  =   "Named:id=34:Heading"
            _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   ":id=34,.wraptext=-1"
            _StyleDefs(64)  =   "Named:id=35:Footing"
            _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   "Named:id=36:Selected"
            _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(68)  =   "Named:id=37:Caption"
            _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(70)  =   "Named:id=38:HighlightRow"
            _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(72)  =   "Named:id=39:EvenRow"
            _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(74)  =   "Named:id=40:OddRow"
            _StyleDefs(75)  =   ":id=40,.parent=33"
            _StyleDefs(76)  =   "Named:id=41:RecordSelector"
            _StyleDefs(77)  =   ":id=41,.parent=34"
            _StyleDefs(78)  =   "Named:id=42:FilterBar"
            _StyleDefs(79)  =   ":id=42,.parent=33"
         End
         Begin BiSANumberBoxProject.BiSANumberBox nTotDebet 
            Height          =   330
            Left            =   1785
            TabIndex        =   47
            Top             =   3195
            Width           =   2430
            _ExtentX        =   4286
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
            Caption         =   "DB"
            CaptionWidth    =   300
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
         Begin BiSANumberBoxProject.BiSANumberBox nTotKredit 
            Height          =   330
            Left            =   4260
            TabIndex        =   48
            Top             =   3195
            Width           =   2430
            _ExtentX        =   4286
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
            Caption         =   "CR"
            CaptionWidth    =   300
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
         Begin BiSANumberBoxProject.BiSANumberBox nSaldoTeller 
            Height          =   330
            Left            =   6930
            TabIndex        =   49
            Top             =   3195
            Width           =   4245
            _ExtentX        =   7488
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
            ForeColor       =   255
            Caption         =   "TOTAL SALDO TELLER"
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
      End
      Begin BiSAFramProject.BiSAFrame FrameMutasiTabungan 
         Height          =   3495
         Left            =   -74910
         Top             =   420
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   6165
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
         Begin BiSAFramProject.BiSAFrame Frameblokir 
            Height          =   750
            Left            =   6495
            Top             =   2325
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   1323
            Caption         =   "REKENING DIBLOKIR SENILAI"
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
            Begin VB.Label lbNilai 
               Alignment       =   2  'Center
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   105
               TabIndex        =   7
               Top             =   330
               Width           =   4365
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame2 
            Height          =   2040
            Left            =   6495
            Top             =   120
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   3598
            Caption         =   "MUTASI"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            Begin BiSANumberBoxProject.BiSANumberBox nAwal 
               Height          =   420
               Left            =   135
               TabIndex        =   8
               Top             =   390
               Width           =   4305
               _ExtentX        =   7594
               _ExtentY        =   741
               Appearance      =   0
               Enabled         =   0   'False
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   12632256
               Caption         =   "SALDO AWAL"
               CaptionWidth    =   1600
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
            Begin BiSANumberBoxProject.BiSANumberBox nMutasi 
               Height          =   420
               Left            =   135
               TabIndex        =   9
               Top             =   855
               Width           =   4305
               _ExtentX        =   7594
               _ExtentY        =   741
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "JUMLAH"
               CaptionWidth    =   1600
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
            Begin BiSANumberBoxProject.BiSANumberBox nAkhir 
               Height          =   420
               Left            =   135
               TabIndex        =   10
               Top             =   1335
               Width           =   4305
               _ExtentX        =   7594
               _ExtentY        =   741
               Appearance      =   0
               Enabled         =   0   'False
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   12632256
               Caption         =   "SALDO AKHIR"
               CaptionWidth    =   1600
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
         Begin BiSATextBoxProject.BiSATextBox cKeteranganTabungan 
            Height          =   330
            Left            =   135
            TabIndex        =   11
            Top             =   2850
            Width           =   6300
            _ExtentX        =   11113
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
            Caption         =   "Keterangan"
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
         Begin BiSATextBoxProject.BiSATextBox cNamaRekeningJurnal 
            Height          =   330
            Left            =   3390
            TabIndex        =   12
            Top             =   2490
            Width           =   3015
            _ExtentX        =   5318
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
         Begin BiSATextBoxProject.BiSATextBox cDK 
            Height          =   330
            Left            =   135
            TabIndex        =   13
            Top             =   2130
            Width           =   2130
            _ExtentX        =   3757
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
            Caption         =   "D/K"
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
         Begin BiSATextBoxProject.BiSATextBox cNamaKodeTransaksi 
            Height          =   330
            Left            =   2565
            TabIndex        =   14
            Top             =   1770
            Width           =   3420
            _ExtentX        =   6033
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
         Begin BiSATextBoxProject.BiSABrowse cKodeTransaksi 
            Height          =   330
            Left            =   135
            TabIndex        =   15
            Top             =   1770
            Width           =   2415
            _ExtentX        =   4260
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
            Caption         =   "Kode Transaksi"
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
         Begin BiSANumberBoxProject.BiSANumberBox nSaldoMinimum 
            Height          =   330
            Left            =   135
            TabIndex        =   16
            Top             =   1050
            Width           =   3270
            _ExtentX        =   5768
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
         Begin BiSATextBoxProject.BiSATextBox cNamaGolTabungan 
            Height          =   330
            Left            =   2415
            TabIndex        =   17
            Top             =   705
            Width           =   3150
            _ExtentX        =   5556
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
         Begin BiSATextBoxProject.BiSATextBox cGolTabungan 
            Height          =   330
            Left            =   135
            TabIndex        =   18
            Top             =   705
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
            Caption         =   "Gol Tabungan"
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
         Begin BiSANumberBoxProject.BiSANumberBox nSetoranMinimum 
            Height          =   330
            Left            =   135
            TabIndex        =   19
            Top             =   1410
            Width           =   3270
            _ExtentX        =   5768
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
            Caption         =   "Setoran Min."
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
         Begin BiSATextBoxProject.BiSABrowse cRekeningJurnal 
            Height          =   330
            Left            =   135
            TabIndex        =   20
            Top             =   2490
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
            BackColor       =   12632256
            Enabled         =   0   'False
            Appearance      =   0
            Caption         =   "Rekening"
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
         Begin BiSATextBoxProject.BiSATextBox cPDL 
            Height          =   330
            Left            =   135
            TabIndex        =   58
            Top             =   345
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
            Left            =   2415
            TabIndex        =   59
            Top             =   345
            Width           =   3150
            _ExtentX        =   5556
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
         Begin VB.Label Label2 
            Caption         =   "[K] = Setoran    [D] = Penarikan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2325
            TabIndex        =   21
            Top             =   2190
            Width           =   2775
         End
      End
      Begin BiSAFramProject.BiSAFrame FrameDeposito 
         Height          =   3705
         Left            =   -74925
         Top             =   330
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6535
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
         Begin VB.OptionButton optAsalPencairan 
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   3240
            TabIndex        =   99
            Top             =   675
            Width           =   765
         End
         Begin VB.OptionButton optAsalPencairan 
            Caption         =   "Tabungan"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   98
            Top             =   675
            Width           =   1200
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame10 
            Height          =   1290
            Left            =   3015
            Top             =   1020
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   2275
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
            Begin BiSATextBoxProject.BiSATextBox ccTab1 
               Height          =   330
               Left            =   270
               TabIndex        =   92
               Top             =   750
               Width           =   405
               _ExtentX        =   714
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
               MaxLength       =   2
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
            Begin BiSATextBoxProject.BiSATextBox ccTab2 
               Height          =   330
               Left            =   690
               TabIndex        =   93
               Top             =   750
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   582
               Text            =   "T1"
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
               GetPicture      =   1
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
            Begin BiSATextBoxProject.BiSATextBox ccTab3 
               Height          =   330
               Left            =   1110
               TabIndex        =   94
               Top             =   750
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
            Begin BiSATextBoxProject.BiSATextBox ccTab4 
               Height          =   330
               Left            =   1905
               TabIndex        =   95
               Top             =   750
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   582
               Text            =   "T1"
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
            Begin BiSANumberBoxProject.BiSANumberBox nSaldoTabungan 
               Height          =   330
               Left            =   255
               TabIndex        =   96
               TabStop         =   0   'False
               Top             =   390
               Width           =   2055
               _ExtentX        =   3625
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
               CaptionWidth    =   0
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
            Begin VB.Label Label9 
               Caption         =   "Saldo"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   97
               Top             =   180
               Width           =   510
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame9 
            Height          =   990
            Left            =   45
            Top             =   2685
            Width           =   5550
            _ExtentX        =   9790
            _ExtentY        =   1746
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
            Begin BiSATextBoxProject.BiSATextBox cTab1 
               Height          =   300
               Left            =   315
               TabIndex        =   88
               Top             =   555
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   529
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
               MaxLength       =   2
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
            Begin VB.OptionButton optTujuanPencairan 
               Caption         =   "Tabungan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   2745
               TabIndex        =   87
               Top             =   240
               Width           =   1200
            End
            Begin VB.OptionButton optTujuanPencairan 
               Caption         =   "Cash"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   1935
               TabIndex        =   86
               Top             =   240
               Width           =   975
            End
            Begin BiSATextBoxProject.BiSATextBox cTab2 
               Height          =   300
               Left            =   750
               TabIndex        =   89
               Top             =   555
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   529
               Text            =   "T1"
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
               GetPicture      =   1
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
            Begin BiSATextBoxProject.BiSATextBox cTab3 
               Height          =   300
               Left            =   1170
               TabIndex        =   90
               Top             =   555
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cTab4 
               Height          =   300
               Left            =   1980
               TabIndex        =   91
               Top             =   555
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   529
               Text            =   "T1"
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
            Begin BiSATextBoxProject.BiSABrowse cKodeTransaksiDepositoTujuanPencairan 
               Height          =   300
               Left            =   2415
               TabIndex        =   102
               Top             =   555
               Visible         =   0   'False
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   529
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
            Begin VB.Label Label8 
               Caption         =   "Tujuan Pencairan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   285
               TabIndex        =   85
               Top             =   225
               Width           =   1575
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame7 
            Height          =   2760
            Left            =   5610
            Top             =   915
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   4868
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
            Begin VB.OptionButton OptCair 
               Caption         =   "Bunga"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   2280
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   105
               Width           =   960
            End
            Begin VB.OptionButton OptCair 
               Caption         =   "Pokok"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   3375
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   105
               Width           =   1245
            End
            Begin BiSANumberBoxProject.BiSANumberBox nFinalti 
               Height          =   330
               Left            =   690
               TabIndex        =   25
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
            Begin BiSANumberBoxProject.BiSANumberBox nPokok 
               Height          =   330
               Left            =   690
               TabIndex        =   26
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
               Caption         =   "Pokok"
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
               TabIndex        =   27
               Top             =   1470
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
               BackColor       =   -2147483634
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
            Begin BiSANumberBoxProject.BiSANumberBox nTotal 
               Height          =   330
               Left            =   690
               TabIndex        =   28
               Top             =   2310
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
               TabIndex        =   29
               Top             =   1830
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
               Caption         =   "Pajak Bunga"
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
            Begin BiSANumberBoxProject.BiSANumberBox nDPMaterai 
               Height          =   330
               Left            =   690
               TabIndex        =   57
               Top             =   1110
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
               Caption         =   "Materai"
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
               Caption         =   "Jenis Pencairan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   750
               TabIndex        =   30
               Top             =   120
               Width           =   1305
            End
            Begin VB.Line Line1 
               X1              =   750
               X2              =   4680
               Y1              =   2220
               Y2              =   2220
            End
         End
         Begin BiSAFramProject.BiSAFrame FrmPesan 
            Height          =   435
            Left            =   5610
            Top             =   495
            Width           =   5445
            _ExtentX        =   9604
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
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Anda Terkena Pinalty Karena Pencairan Pokok Belum Jatuh Tempo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   105
               Width           =   5010
            End
         End
         Begin BiSADateProject.BiSADate dTempo 
            Height          =   330
            Left            =   45
            TabIndex        =   33
            Top             =   1260
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   582
            Value           =   "13-10-2005"
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
            Caption         =   "Jatuh tempo"
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
            Left            =   45
            TabIndex        =   34
            Top             =   540
            Width           =   2325
            _ExtentX        =   4101
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
         Begin BiSANumberBoxProject.BiSANumberBox nBunga 
            Height          =   330
            Left            =   45
            TabIndex        =   35
            Top             =   1620
            Width           =   2430
            _ExtentX        =   4286
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
            Caption         =   "Bunga (%) p.a"
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
            Left            =   45
            TabIndex        =   22
            Top             =   2340
            Width           =   3930
            _ExtentX        =   6932
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
            Caption         =   "Nominal"
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
            Left            =   45
            TabIndex        =   32
            Top             =   180
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   582
            Value           =   "13-10-2005"
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
         Begin BiSANumberBoxProject.BiSANumberBox nPersFinalti 
            Height          =   330
            Left            =   45
            TabIndex        =   56
            Top             =   1980
            Width           =   2430
            _ExtentX        =   4286
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
            Caption         =   "Finalty (%)"
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
         Begin BiSATextBoxProject.BiSABrowse cKodeTransaksiDeposito 
            Height          =   345
            Left            =   3990
            TabIndex        =   101
            Top             =   2325
            Width           =   855
            _ExtentX        =   1508
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
         Begin BiSANumberBoxProject.BiSANumberBox nARO 
            Height          =   330
            Left            =   45
            TabIndex        =   105
            Top             =   900
            Width           =   2325
            _ExtentX        =   4101
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
            Caption         =   "ARO"
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
         Begin VB.Label Label13 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2445
            TabIndex        =   106
            Top             =   960
            Width           =   270
         End
         Begin VB.Label Label10 
            Caption         =   "Pencairan Dari"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3240
            TabIndex        =   100
            Top             =   390
            Width           =   1335
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
            Height          =   255
            Left            =   2430
            TabIndex        =   52
            Top             =   600
            Width           =   720
         End
      End
      Begin BiSAFramProject.BiSAFrame FramePencairan 
         Height          =   3420
         Left            =   -74850
         Top             =   465
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6033
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
         Begin BiSANumberBoxProject.BiSANumberBox nPlafondCair 
            Height          =   330
            Left            =   105
            TabIndex        =   36
            Top             =   165
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
            Caption         =   "Plafond"
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
         Begin BiSANumberBoxProject.BiSANumberBox nAdministrasi 
            Height          =   330
            Left            =   6150
            TabIndex        =   37
            Top             =   225
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "Administrasi"
            CaptionWidth    =   1600
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
         Begin BiSANumberBoxProject.BiSANumberBox nMaterai 
            Height          =   330
            Left            =   6150
            TabIndex        =   38
            Top             =   945
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "Materai"
            CaptionWidth    =   1600
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
         Begin BiSANumberBoxProject.BiSANumberBox nTotalPencairan 
            Height          =   330
            Left            =   6150
            TabIndex        =   39
            Top             =   2925
            Width           =   3795
            _ExtentX        =   6694
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
            BackColor       =   -2147483647
            ForeColor       =   -2147483634
            Caption         =   "TOTAL CAIR"
            CaptionWidth    =   1600
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
         Begin BiSADateProject.BiSADate dTglRealisasiCair 
            Height          =   330
            Left            =   120
            TabIndex        =   42
            Top             =   885
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   582
            Value           =   "13-10-2005"
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
            Caption         =   "Tgl Realisasi"
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
         Begin BiSATextBoxProject.BiSATextBox cSpkCair 
            Height          =   330
            Left            =   120
            TabIndex        =   43
            Top             =   525
            Width           =   4320
            _ExtentX        =   7620
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
            Caption         =   "Nomor SPK"
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
         Begin BiSANumberBoxProject.BiSANumberBox nBungaCair 
            Height          =   330
            Left            =   120
            TabIndex        =   44
            Top             =   1245
            Width           =   2430
            _ExtentX        =   4286
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
            Caption         =   "Bunga (%) p.a"
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
         Begin BiSANumberBoxProject.BiSANumberBox nLamaCair 
            Height          =   330
            Left            =   120
            TabIndex        =   45
            Top             =   1605
            Width           =   2430
            _ExtentX        =   4286
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
            Caption         =   "Lama"
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
         Begin BiSANumberBoxProject.BiSANumberBox nProvisi 
            Height          =   330
            Left            =   6150
            TabIndex        =   53
            Top             =   585
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "Provisi"
            CaptionWidth    =   1600
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
         Begin BiSANumberBoxProject.BiSANumberBox nNotaris 
            Height          =   330
            Left            =   6150
            TabIndex        =   54
            Top             =   1305
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "Notaris"
            CaptionWidth    =   1600
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
         Begin BiSANumberBoxProject.BiSANumberBox nLainLain 
            Height          =   330
            Left            =   6150
            TabIndex        =   55
            Top             =   1665
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "Biaya Lain-Lain"
            CaptionWidth    =   1600
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
         Begin BiSANumberBoxProject.BiSANumberBox nSimpananWajib 
            Height          =   330
            Left            =   6150
            TabIndex        =   108
            Top             =   2025
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "SimpananWajib"
            CaptionWidth    =   1600
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
         Begin VB.Label Label5 
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
            Height          =   210
            Left            =   2640
            TabIndex        =   51
            Top             =   1650
            Width           =   720
         End
         Begin VB.Line Line2 
            X1              =   6210
            X2              =   10140
            Y1              =   2835
            Y2              =   2835
         End
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1845
      Left            =   0
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   3254
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame8 
         Height          =   645
         Left            =   135
         Top             =   405
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   1138
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
         Begin BiSATextBoxProject.BiSABrowse cGolongan 
            Height          =   390
            Left            =   2535
            TabIndex        =   60
            Top             =   135
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   688
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Verdana"
            FontSize        =   12
            BackColor       =   16777215
            ForeColor       =   0
            GetPicture      =   1
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
            Left            =   225
            TabIndex        =   61
            Top             =   135
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   688
            Text            =   "12"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Verdana"
            FontSize        =   12
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   2
            Caption         =   "NO. REKENING"
            CaptionWidth    =   1700
            CaptionBackColor=   12632256
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
            Left            =   3345
            TabIndex        =   62
            Top             =   135
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   688
            Text            =   "123456"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Verdana"
            FontSize        =   12
            BackColor       =   16777215
            ForeColor       =   0
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
            Height          =   390
            Left            =   4500
            TabIndex        =   63
            Top             =   135
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   688
            Text            =   "12"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Verdana"
            FontSize        =   12
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   2
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
      Begin BiSATextBoxProject.BiSATextBox cShow 
         Height          =   360
         Left            =   4635
         TabIndex        =   0
         Top             =   1350
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   635
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
      Begin BiSATextBoxProject.BiSATextBox cJenisProduk 
         Height          =   360
         Left            =   4260
         TabIndex        =   1
         Top             =   1350
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   1740
         Left            =   7800
         Top             =   45
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   3069
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
         Begin VB.Image Image2 
            Height          =   1635
            Left            =   60
            Stretch         =   -1  'True
            Top             =   60
            Width           =   3495
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame3 
         Height          =   1740
         Left            =   6135
         Top             =   45
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   3069
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
         Begin VB.Image Image1 
            Height          =   1635
            Left            =   60
            Stretch         =   -1  'True
            Top             =   60
            Width           =   1545
         End
      End
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   1065
         Width           =   5955
         _ExtentX        =   10504
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
         Caption         =   "Nama"
         CaptionWidth    =   800
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
      Begin BiSATextBoxProject.BiSATextBox cFaktur 
         Height          =   300
         Left            =   2475
         TabIndex        =   3
         Top             =   75
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
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
         BackColor       =   12632256
         Enabled         =   0   'False
         MaxLength       =   20
         Appearance      =   0
         Caption         =   "No Transaksi"
         CaptionWidth    =   1200
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
         Height          =   300
         Left            =   150
         TabIndex        =   4
         Top             =   75
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   529
         Value           =   "13-10-2005"
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
         CaptionWidth    =   800
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
      Begin BiSATextBoxProject.BiSATextBox cAlamat 
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   1410
         Width           =   5955
         _ExtentX        =   10504
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
         CaptionWidth    =   800
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame5 
      Height          =   585
      Left            =   0
      Top             =   6540
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   1032
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
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9300
         TabIndex        =   40
         Top             =   75
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
         Picture         =   "trTeller.frx":00E0
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10380
         TabIndex        =   41
         Top             =   75
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
         Picture         =   "trTeller.frx":067A
      End
      Begin BiSAButtonProject.BiSAButton cmdBatal 
         Height          =   435
         Left            =   7800
         TabIndex        =   50
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
         Caption         =   "    &Clear/Batal"
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
         Picture         =   "trTeller.frx":0720
      End
   End
End
Attribute VB_Name = "trTeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As Single
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim dbData1 As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim objMenu As New CodeSuiteLibrary.Menu
Dim lStatusBlokir As Boolean
Dim nJumlahBlokir As Double
Dim vaArray As New XArrayDB
Dim xArray As New XArrayDB
Dim nBakiDebet As Double
Dim cPilihanTransaksi As String
Dim lStatusNominal As Boolean
Dim cStatusPostingPokok As String
Dim nSisaAngsBunga As Double
Dim nSisaAngsPokok As Double
Dim nDendaKeterlamabatan As Double
Dim vaGrid As New XArrayDB

Private Sub ccTab3_Validate(Cancel As Boolean)
  If Trim(ccTab3.Text) <> "" Then
    ccTab3.Text = Padl(ccTab3.Text, 6, "0")
  End If
End Sub

Private Sub ccTab4_Validate(Cancel As Boolean)
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "Tabungan t", "t.rekening,r.Nama,r.Alamat", "t.Rekening", sisAssign, GetRekTabungan2, " and t.Close<>'1'", "t.Rekening", Array("Left Join RegisterNasabah r on r.Kode = t.Kode"))
  If Not db.eof Then
    MsgBox "Informasi Rekening: " & vbCrLf & _
           "Rekening No. " & GetNull(db!Rekening) & vbCrLf & _
           "Nama. " & GetNull(db!nama) & vbCrLf & _
           "Alamat. " & GetNull(db!alamat)
    nSaldoTabungan.Value = GetSaldoTab(objData, GetRekTabungan2, Date)
  Else
    optAsalPencairan(0).Value = True
    optAsalPencairan(0).SetFocus
    MsgBox "No Rekening yang dimasukkan tidak valid", vbInformation, Me.Caption
    GetAsalPencairan False
    Exit Sub
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Produk", "Kode,Keterangan", "Kode", sisContent, cGolongan.Text, " Group By Kode")
  cGolongan.Text = cGolongan.Browse(dbData)
  cJenisProduk.Text = left(cGolongan.Text, 1)
  Select Case cJenisProduk.Text
    Case Is = "T"
    Case Is = "D"
    Case Is = "K"
  End Select
End Sub
Private Sub GetDataTabungan()
Dim cField As String
Dim vaJoin
Dim cNoRekening As String
Dim nTAkhir As Double
  
    cNoRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    cField = "t.rekening,t.StatusBlokir,t.JumlahBlokir,r.Nama,r.Alamat,r.Path,r.Path1,g.Keterangan,g.SaldoMinimum,g.SetoranMinimum,t.PDL,p.Keterangan as NamaPDL"
    vaJoin = Array("Left Join RegisterNasabah r on r.Kode = t.Kode", _
                   "Left Join Golongantabungan g on g.Kode = t.golongantabungan", _
                   "Left Join PDL p on p.Kode = t.PDL")
    Set dbData = objData.Browse(GetDSN, "Tabungan t", cField, "t.Rekening", sisAssign, cNoRekening, , , vaJoin)
    If Not dbData.eof Then
      cPDL.Text = GetNull(dbData!PDL)
      cNamaPDL.Text = GetNull(dbData!namapdl)
      cGolTabungan.Text = cGolongan.Text
      cNamaGolTabungan.Text = GetNull(dbData!Keterangan, "")
      nSaldoMinimum.Value = GetNull(dbData!SaldoMinimum)
      nSetoranMinimum.Value = GetNull(dbData!SetoranMinimum)
      cNama.Text = GetNull(dbData!nama, "")
      cAlamat.Text = GetNull(dbData!alamat, "")
      Frameblokir.Visible = False
      If GetNull(dbData!StatusBlokir) = "1" Or GetNull(dbData!JumlahBlokir) > 0 Then
        lStatusBlokir = True
        nJumlahBlokir = GetNull(dbData!JumlahBlokir)
        Frameblokir.Visible = True
        lbNilai.Caption = "RP : " & Format(nJumlahBlokir, "###,###,###,###,##0.00")
      Else
         lStatusBlokir = False
      End If
      GetGambar Image1, Image2, GetNull(dbData!Path, ""), GetNull(dbData!Path1, "")
      nAwal.Value = GetSaldoTab(objData, cNoRekening, Date)
      Exit Sub
    End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  Set dbData = objData.Browse(GetDSN, "Produk", "Kode,Keterangan", "Kode", sisContent, cGolongan.Text, " Group By Kode")
  cGolongan.Text = GetNull(dbData!Kode) 'cGolongan.Browse(dbData)
  cJenisProduk.Text = left(cGolongan.Text, 1)
  Select Case cJenisProduk.Text
    Case Is = "T"
    Case Is = "D"
    Case Is = "K"
  End Select
End Sub

Private Sub cKodeTransaksi_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "KodeTransaksi k", "k.Kode", cKodeTransaksi, "k.Kode,k.Keterangan,k.DK,k.Kas,k.Rekening", " and (t.Level > 0 or " & nUserLevel & " = 0)", _
               Array("Left Join KodetransaksiTeller t on k.Kode = t.Kode and Level = " & nUserLevel))
  If Not dbData.eof Then
    cNamaKodeTransaksi.Text = GetNull(dbData!Keterangan)
    cDK.Text = GetNull(dbData!DK)
    cRekeningJurnal.Default
    cNamaRekeningJurnal.Default
    
    cRekeningJurnal.Text = IIf(GetNull(dbData!Kas) = "K", cKasTeller, GetNull(dbData!Rekening))
    Set dbData = objData.Browse(GetDSN, "Rekening", , "Kode", sisAssign, cRekeningJurnal.Text)
    If Not dbData.eof Then
      cNamaRekeningJurnal.Text = GetNull(dbData!Keterangan, "")
    End If
    cKeteranganTabungan.Text = cNamaKodeTransaksi.Text & " a.n " & cNama.Text
  End If
End Sub

'Private Function GetValidMutasiTabungan(ByVal fObj As CodeSuiteLibrary.data, ByVal fRekening As String, ByVal fTgl As Date) As Boolean
'Dim db As New ADODB.Recordset
'
'  GetValidMutasiTabungan = True
'
'  set db = fObj.Browse(getdsn,"MutasiTabungan",,"Rekening",sisAssign,fRekening," and "
'End Function

Private Sub cKodeTransaksiDeposito_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "KodeTransaksi", "Kode,Keterangan", "Kode", sisContent, cKodeTransaksiDeposito.Text)
  If Not dbData.eof Then
    cKodeTransaksiDeposito.Text = cKodeTransaksiDeposito.Browse(dbData)
    cKodeTransaksiDeposito.Text = GetNull(dbData!Kode)
  End If
End Sub

Private Sub cKodeTransaksiDepositoTujuanPencairan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "KodeTransaksi", "Kode,Keterangan", "Kode", sisContent, cKodeTransaksiDepositoTujuanPencairan.Text)
  If Not dbData.eof Then
    cKodeTransaksiDepositoTujuanPencairan.Text = cKodeTransaksiDepositoTujuanPencairan.Browse(dbData)
    cKodeTransaksiDepositoTujuanPencairan.Text = GetNull(dbData!Kode, "")
  End If
End Sub

Private Sub cmdBatal_Click()
  initvalue
  GetLock
  If cGolongan.Enabled = True Then
    cGolongan.SetFocus
  Else
    cCabang.SetFocus
  End If
  Exit Sub
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  If ValidSimpan() Then
    If MsgBox("Data benar-benar sudah valid ?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
      Select Case cPilihanTransaksi
        Case Is = "1"
          SimpanTabungan
        Case Is = "2"
          SimpanDeposito SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
        Case Is = "3"
          SimpanPencairanKredit
        Case Is = "4"
'          If isValidSimpanAngsuran() Then
'            SimpanAngsuranKredit SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
'          End If
           SimpanAngsuranKredit SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
      End Select
      GetSaldoTeller
      initvalue
      GetLock
      Exit Sub
    End If
  End If
End Sub

Private Function isValidSimpanAngsuran() As Boolean
isValidSimpanAngsuran = True

  If UCase(lbCaraPerhitungan.Caption) = "FLAT" Then
    If nSisaPokok.Value <= 0 Then
      If nAngsuranBunga.Value < nBungaLalu.Value Then
        MsgBox "Sisa bunga harus dilunasi, data tidak bisa disimpan"
        isValidSimpanAngsuran = False
        Exit Function
      End If
    End If
  End If
  
End Function

Private Sub GetDataProduk()
Dim cSQL As String
  
  cSQL = "Select Kode,Keterangan From GolonganTabungan"
  cSQL = cSQL & " Union"
  cSQL = cSQL & " Select Kode,Keterangan From GolonganDeposito"
  cSQL = cSQL & " Union"
  cSQL = cSQL & " Select Kode,Keterangan From GolonganKredit"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      objData.Update GetDSN, "Produk", "Kode='" & GetNull(dbData!Kode, "") & "'", Array("Kode", "Keterangan"), Array(GetNull(dbData!Kode, ""), GetNull(dbData!Keterangan, ""))
      dbData.MoveNext
    Loop
  End If
End Sub

Private Sub cTab3_Validate(Cancel As Boolean)
  If Trim(cTab3.Text) <> "" Then
    cTab3.Text = Padl(cTab3.Text, 6, "0")
  End If
End Sub

Private Sub cTab4_Validate(Cancel As Boolean)
Dim db As New ADODB.Recordset
Cancel = False

  Set db = objData.Browse(GetDSN, "Tabungan t", "t.rekening,r.Nama,r.Alamat", "t.Rekening", sisAssign, GetRekTabungan, " and t.Close<>'1'", "t.Rekening", Array("Left Join RegisterNasabah r on r.Kode = t.Kode"))
  If Not db.eof Then
    MsgBox "Informasi Rekening: " & vbCrLf & _
           "Rekening No. " & GetNull(db!Rekening) & vbCrLf & _
           "Nama. " & GetNull(db!nama) & vbCrLf & _
           "Alamat. " & GetNull(db!alamat)
    'cKodeTransaksiDepositoTujuanPencairan.SetFocus
  Else
    optTujuanPencairan(0).Value = True
    'optTujuanPencairan(0).SetFocus
    'cKodeTransaksiDepositoTujuanPencairan.SetFocus
    MsgBox "No Rekening yang dimasukkan tidak valid", vbInformation, Me.Caption
'    GetTujuanPencairan False
    Cancel = True
    Exit Sub
  End If
End Sub

Private Function GetRekTabungan() As String
  GetRekTabungan = cTab1.Text & "." & cTab2.Text & "." & cTab3.Text & "." & cTab4.Text
End Function

Private Function GetRekTabungan2() As String
  GetRekTabungan2 = ccTab1.Text & "." & ccTab2.Text & "." & ccTab3.Text & "." & ccTab4.Text
End Function

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Then
    Cancel = True
    dTgl.SetFocus
  End If
  GetSaldoTeller
End Sub

Private Sub nAngsuranBunga_Change()
  nKewajiban.Value = nAngsuranPokok.Value + nAngsuranBunga.Value + nDenda.Value
End Sub

Private Sub nAngsuranPokok_Change()
  nKewajiban.Value = nAngsuranPokok.Value + nAngsuranBunga.Value + nDenda.Value
End Sub

Private Sub nAngsuranPokok_Validate(Cancel As Boolean)
  nSisaPokok.Value = Round(nPlafondAngsuran.Value - nBakiDebet - nAngsuranPokok.Value)
  If nSisaPokok.Value < 0 Then
    MsgBox "Angsuran Pokok melebihi Baki Debet..", vbInformation
    Cancel = True
    nAngsuranPokok.SetFocus
    Exit Sub
  End If
End Sub

Private Sub nDenda_Change()
  nKewajiban.Value = nAngsuranPokok.Value + nAngsuranBunga.Value + nDenda.Value
End Sub

Private Sub nDPMaterai_Change()
  nTotal.Value = nPokok.Value - nFinalti.Value - nDPMaterai.Value
End Sub

Private Sub nMutasi_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nNilaiAkhir As Double
Dim db As New ADODB.Recordset

  If KeyCode = 13 Or KeyCode = 40 Then
    If nMutasi.Value <= 0 Then
       MsgBox "Nilai Mutasi tidak VALID. Silahkan ulangi pengisian", vbOKOnly + vbInformation, Me.Caption
       nMutasi.SetFocus
       Exit Sub
    End If
    
    If cDK.Text = "K" And nMutasi.Value < nSetoranMinimum.Value And nMutasi.Value <> 0 Then
       MsgBox "Maaf, Setoran Tabungan Minimal : Rp. " & Format((nSetoranMinimum.Value), "#,##,###.00"), vbInformation, Me.Caption
       nMutasi.SetFocus
       Exit Sub
    End If
    
     nNilaiAkhir = nAwal.Value + IIf(cDK.Text = "K", nMutasi.Value, -nMutasi.Value)
     If lStatusBlokir = True And cDK.Text = "D" Then
      If nNilaiAkhir < nJumlahBlokir + nSaldoMinimum.Value Then
         MsgBox "Maaf, Saldo Tabungan Anda Tidak Cukup. Silahkan Mengulangi Pengisian !", vbOKOnly + vbInformation, Me.Caption
         nAkhir.Value = 0
         nMutasi.SetFocus
         Exit Sub
      End If
      
    Else
      If (nNilaiAkhir < nSaldoMinimum.Value) Or nNilaiAkhir < 0 Then
          MsgBox "Maaf, Setoran tunai tidak boleh lebih kecil dari SALDO MINIMUM. Silahkan Mengulangi Pengisian !", vbOKOnly + vbInformation, Me.Caption
          nAkhir.Value = 0
          nMutasi.SetFocus
          Exit Sub
      End If
    End If
    nAkhir.Value = nAwal.Value + IIf(cDK.Text = "K", nMutasi.Value, -nMutasi.Value)
  End If
End Sub

Private Sub GetSaldoTeller()
Dim n As Long
Dim nSaldo As Double
Dim cSQL As String

  nTotDebet.Value = 0
  nTotKredit.Value = 0
  nSaldo = 0
  cSQL = cSQL & "Select Awal From SaldoRekening Where Rekening = '" & cKasTeller & "' "
  cSQL = cSQL & " Union "
  cSQL = cSQL & "Select Sum(b.Debet-b.Kredit) as Awal From BukuBesar b Where b.Tgl < '" & Format(dTgl.Value, "yyyy-MM-dd") & "' and b.Rekening = '" & cKasTeller & "'"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    nSaldo = nSaldo + GetNull(dbData!Awal)
  End If
  vaArray.ReDim 0, 0, 0, 4
  vaArray(n, 2) = "Saldo Awal Teller"
  vaArray(n, 3) = IIf(nSaldo >= 0, "D", "K")
  vaArray(n, 4) = nSaldo
  nTotDebet.Value = nTotDebet.Value + IIf(vaArray(n, 3) = "D", vaArray(n, 4), 0)
  nTotKredit.Value = nTotKredit.Value + IIf(vaArray(n, 3) = "K", vaArray(n, 4), 0)
  
  Set dbData = objData.Browse(GetDSN, "BukuBesar", "Faktur,Keterangan,Debet,Kredit", "Tgl", sisAssign, Format(dTgl.Value, "yyyy-mm-dd"), " and Rekening = '" & cKasTeller & "'", "Tgl,Rekening,ID")
  n = 1
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
        vaArray.InsertRows n
        vaArray(n, 0) = (n)
        vaArray(n, 1) = GetNull(dbData!Faktur)
        vaArray(n, 2) = GetNull(dbData!Keterangan)
        If GetNull(dbData!Debet) <> 0 Then
          vaArray(n, 3) = "D"
          vaArray(n, 4) = GetNull(dbData!Debet)
        Else
          vaArray(n, 3) = "K"
          vaArray(n, 4) = GetNull(dbData!Kredit)
        End If
        nTotDebet.Value = nTotDebet.Value + IIf(vaArray(n, 3) = "D", vaArray(n, 4), 0)
        nTotKredit.Value = nTotKredit.Value + IIf(vaArray(n, 3) = "K", vaArray(n, 4), 0)
        n = n + 1
      dbData.MoveNext
    Loop
  End If
  nSaldoTeller.Value = nTotDebet.Value - nTotKredit.Value
  nSaldoTeller.ForeColor = IIf(nSaldoTeller.Value < 0, &HFF&, &H80000008)
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub nMutasi_Validate(Cancel As Boolean)
Dim nNilaiAkhir As Double
Dim db As New ADODB.Recordset
  
    If nMutasi.Value <= 0 Then
       MsgBox "Nilai Mutasi tidak VALID. Silahkan ulangi pengisian", vbOKOnly + vbInformation, Me.Caption
       nMutasi.SetFocus
       Cancel = True
       Exit Sub
    End If
    
    If cDK.Text = "K" And nMutasi.Value < nSetoranMinimum.Value And nMutasi.Value <> 0 Then
       MsgBox "Maaf, Setoran Tabungan Minimal : Rp. " & Format((nSetoranMinimum.Value), "#,##,###.00"), vbInformation, Me.Caption
       nMutasi.SetFocus
       Cancel = True
       Exit Sub
    End If
    
     nNilaiAkhir = nAwal.Value + IIf(cDK.Text = "K", nMutasi.Value, -nMutasi.Value)
     If lStatusBlokir = True And cDK.Text = "D" Then
      If nNilaiAkhir < nJumlahBlokir + nSaldoMinimum.Value Then
         MsgBox "Maaf, Saldo Tabungan Anda Tidak Cukup. Silahkan Mengulangi Pengisian !", vbOKOnly + vbInformation, Me.Caption
         nAkhir.Value = 0
         nMutasi.SetFocus
         Cancel = True
         Exit Sub
      End If
    Else
      If (nNilaiAkhir < nSaldoMinimum.Value) Or nNilaiAkhir < 0 Then
          MsgBox "Maaf, Setoran tunai tidak boleh lebih kecil dari SALDO MINIMUM. Silahkan Mengulangi Pengisian !", vbOKOnly + vbInformation, Me.Caption
          nAkhir.Value = 0
          nMutasi.SetFocus
          Cancel = True
          Exit Sub
      End If
    End If
    nAkhir.Value = nAwal.Value + IIf(cDK.Text = "K", nMutasi.Value, -nMutasi.Value)

    Set db = objData.Browse(GetDSN, "UserName", "Plafond", "Username", sisAssign, GetRegistry(reg_UserName))
    If GetNull(db!plafond) < nAkhir.Value And cDK.Text = "D" Then
      MsgBox "Maksimal plafond/uang yang anda keluarkan melebihi plafond yang telah ter setup.." & vbCrLf & "Untuk melanjutkan transaksi Anda memerlukan otorisasi"
      If objMenu.GetPassword("USPD", GetDSN, Me) Then
        Set dbData = objData.Browse(GetDSN, "UserName", "Plafond", "UserName", sisAssign, objMenu.UserName)
        If GetNull(dbData!plafond) < nMutasi.Value Then
          MsgBox "Maaf anda tidak berhak melakukan otorisasi!!", vbExclamation, Me.Caption
          'nPlafondCair.SetFocus
          Cancel = True
        End If
      Else
        'nPlafondCair.SetFocus
        Cancel = True
      End If
    End If
End Sub

Private Sub nNominalDeposito_Validate(Cancel As Boolean)
Dim nSal As Double

  nSal = GetSaldoTab(objData, GetRekTabungan2, Date)
  If optAsalPencairan(1).Value = True Then
    If GetValidAsalPencairan(nSal) Then
      nSaldoTabungan.Value = nSal - nNominalDeposito.Value
    Else
      Cancel = True
    End If
  End If
End Sub

Private Function GetValidAsalPencairan(ByVal NominalSaldo As Double) As Boolean
  GetValidAsalPencairan = True
  If nNominalDeposito.Value > NominalSaldo Then
    MsgBox "Nominal yang dimasukkan tidak valid!!", vbInformation
    nNominalDeposito.Value = 0
    nNominalDeposito.SetFocus
    GetValidAsalPencairan = False
  End If
End Function

Private Sub nPlafondCair_Validate(Cancel As Boolean)
Dim db As New ADODB.Recordset

  nTotalPencairan.Value = nPlafondCair.Value - nAdministrasi.Value - nMaterai.Value - nProvisi.Value - nNotaris.Value - nLainLain.Value - nSimpananWajib.Value
  Set db = objData.Browse(GetDSN, "UserName", "Plafond", "Username", sisAssign, GetRegistry(reg_UserName))
  If GetNull(db!plafond) < nTotalPencairan.Value Then
    MsgBox "Maksimal plafond/uang yang anda keluarkan melebihi plafond yang telah ter setup.." & vbCrLf & "Untuk melanjutkan transaksi Anda memerlukan otorisasi"
    If objMenu.GetPassword("USPD", GetDSN, Me) Then
      Set dbData = objData.Browse(GetDSN, "UserName", "Plafond", "UserName", sisAssign, objMenu.UserName)
      If GetNull(dbData!plafond) < nTotalPencairan.Value Then
        MsgBox "Maaf anda tidak berhak melakukan otorisasi!!", vbExclamation, Me.Caption
        nPlafondCair.SetFocus
        Cancel = True
      Else
        cusername = objMenu.UserName
      End If
    Else
      nPlafondCair.SetFocus
      Cancel = True
    End If
  End If
End Sub

Private Sub optAsalPencairan_Click(Index As Integer)
  Select Case Index
    Case 0
      GetAsalPencairan False
    Case 1
      GetAsalPencairan True
  End Select
End Sub

Private Sub GetAsalPencairan(Optional ByVal lOpt As Boolean = True)
  ccTab1.Default
  ccTab2.Default
  ccTab3.Default
  ccTab4.Default
  cKodeTransaksiDeposito.Default
  nSaldoTabungan.Value = 0
  If lOpt = False Then
    ccTab1.Enabled = False
    ccTab2.Enabled = False
    ccTab3.Enabled = False
    ccTab4.Enabled = False
    ccTab1.BackColor = &HC0C0C0
    ccTab2.BackColor = &HC0C0C0
    ccTab3.BackColor = &HC0C0C0
    ccTab4.BackColor = &HC0C0C0
    cKodeTransaksiDeposito.Enabled = False
    cKodeTransaksiDeposito.BackColor = &HC0C0C0
  Else
    ccTab1.Enabled = True
    ccTab2.Enabled = True
    ccTab3.Enabled = True
    ccTab4.Enabled = True
    ccTab1.BackColor = &H80000005
    ccTab2.BackColor = &H80000005
    ccTab3.BackColor = &H80000005
    ccTab4.BackColor = &H80000005
    cKodeTransaksiDeposito.Enabled = True
    cKodeTransaksiDeposito.BackColor = &H80000005
  End If
End Sub

Private Sub optAsalPencairan_GotFocus(Index As Integer)
  Select Case Index
    Case 0
      GetAsalPencairan False
    Case 1
      GetAsalPencairan True
  End Select
End Sub

Private Sub optAsalPencairan_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optTujuanPencairan_GotFocus(Index As Integer)
  Select Case Index
    Case 0
      GetTujuanPencairan False
    Case 1
      GetTujuanPencairan True
  End Select
End Sub

Private Sub optTujuanPencairan_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  If Value = 0 Then
    Value = ""
  Else
    Value = Format(Value, "###,###,###,###,##0.00")
  End If
End Sub

Private Sub SimpanTabungan()
Dim vaField, vaValue
Dim cFakturTabungan As String
Dim cRek As String
  
    vaField = Array("Faktur", "Tgl", "KodeTransaksi", "Rekening", "Jumlah", "UserName", "DateTime", "Keterangan")
    cFakturTabungan = GetLastFaktur(fkt_MutasiTabungan, dTgl.Value, True)
    cRek = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
            
    UpdMutasiTabungan objData, cKodeTransaksi.Text, cFakturTabungan, dTgl.Value, cRek, nMutasi.Value, True, cKeteranganTabungan.Text, , cDK.Text, cRekeningJurnal.Text
    UpdUrutFaktur objData, cFakturTabungan
'    If MsgBox("Akan mencetak Validasi Tabungan ?", vbYesNo, "Transaksi Mutasi Tabungan") = vbYes Then
'      CetakValidasiTabungan cFakturTabungan, dTgl.Value, SNow, cRek, GetRegistry(reg_UserName), cKodeTransaksi.Text, cNamaKodeTransaksi.Text, cDK.Text, nMutasi.Value
'    End If
    GetSaldoTeller
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  If cUrut.LastKey = 13 Or cUrut.LastKey = 40 Then
    cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
  End If
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cNomorRekening As String

  If Trim(cFrekuensi.Text) <> "" Then
        If CekTeller(cNomorRekening) = False Then
          MsgBox "Anda Login Tidak Sebagai Teller, " & vbCrLf & " Silahkan Konfigurasi Terlebih Dahulu Username Anda..", vbExclamation, Me.Caption
          Cancel = True
          cFrekuensi.SetFocus
          Exit Sub
        End If
        
        cNomorRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
        If CekRekening(cNomorRekening) = False Then
          MsgBox "Nomor rekening tidak ada...", vbInformation
          Cancel = True
          cFrekuensi.SetFocus
          Exit Sub
        End If
     
        If left(cJenisProduk.Text, 1) = "T" Then
            cPilihanTransaksi = "1"
            PilihSSTAB "T"
            GetSaldoTeller
            If GetValidOpenTabungan Then
              GetDataTabungan
              GetJenisTransaksi
            End If
          End If
          
          If left(cJenisProduk.Text, 1) = "D" Then
            cPilihanTransaksi = "2"
            PilihSSTAB "D"
            If GetValidOpenDeposito Then
              GetDataDeposito cNomorRekening
              GetJenisTransaksi
            End If
          End If
          
          If left(cJenisProduk.Text, 1) = "K" Then
            Set dbData = objData.Browse(GetDSN, "Debitur d", "d.*,r.nama,r.Alamat", "d.Rekening", sisAssign, cNomorRekening, , , _
                                        Array("Left Join Registernasabah r on r.Kode = d.Kode"))
            If Not dbData.eof Then
              
              'cara perhitungan
              '1 menurun
              '2 flat
              
              Select Case dbData!caraperhitungan
                Case 1
                  lbCaraPerhitungan.Visible = True
                  lbCaraPerhitungan.Caption = "Menurun"
                Case 2
                  lbCaraPerhitungan.Visible = True
                  lbCaraPerhitungan.Caption = "Flat"
              End Select
              
              If GetNull(dbData!status) = "1" Then
                MsgBox "Rekening ini sudah lunas....", vbInformation
                initvalue
                Cancel = True
                cFrekuensi.SetFocus
                Exit Sub
              End If
              
              If GetNull(dbData!statuspencairan) = "1" Then
                cPilihanTransaksi = "4"
                PilihSSTAB "A"
                GetJenisTransaksi
                cCaraAngsuran.Text = GetNull(dbData!CaraAngsuran)
                SSTab2.Tab = 0
                
                GetDataAngsuran GetNull(dbData!NoSPK, ""), GetNull(dbData!Tgl), GetNull(dbData!SukuBunga), _
                                GetNull(dbData!Lama), GetNull(dbData!plafond), GetNull(dbData!nama, ""), GetNull(dbData!alamat, ""), _
                                GetNull(dbData!PeriodeBungaMenurun), GetNull(dbData!MinimalPeriode), GetNull(dbData!KonpensasiTelat), _
                                GetNull(dbData!DendaTelatBayar), GetNull(dbData!SimpananWajib)
                
                GetGridPeriode
              Else
                cPilihanTransaksi = "3"
                PilihSSTAB "C"
                SSTab2.Tab = 0
                GetDataPencairan GetNull(dbData!NoSPK, ""), GetNull(dbData!Tgl), GetNull(dbData!SukuBunga), _
                                  GetNull(dbData!Lama), GetNull(dbData!plafond), GetNull(dbData!Administrasi), GetNull(dbData!Materai), GetNull(dbData!nama, ""), GetNull(dbData!alamat, ""), GetNull(dbData!Provisi), GetNull(dbData!Notaris), GetNull(dbData!BiayaLainLain), GetNull(dbData!SimpananWajib)
                cFaktur.Text = GetLastFaktur(fkt_Relisasi, dTgl.Value, False)
                GetGridPeriode
              End If
            End If
        End If
    End If
End Sub

Private Function GetValidOpenDeposito() As Boolean
Dim dbOpenDeposito As New ADODB.Recordset
Dim cNomorRekeningDeposito As String

  GetValidOpenDeposito = True
  cNomorRekeningDeposito = cCabang.Text & "." & cGolongan.Text & "." & cUrut.Text & "." & cFrekuensi.Text
  
  Set dbOpenDeposito = objData.Browse(GetDSN, "Deposito", "StatusBlokir,Status", "Rekening", sisAssign, cNomorRekeningDeposito, " and StatusBlokir = 'Y'")
  If Not dbOpenDeposito.eof Then
    If GetNull(dbOpenDeposito!StatusBlokir) = "Y" Then
      MsgBox "Rekening ini, " & cNomorRekeningDeposito & " Sedang di blokir", vbInformation
      GetValidOpenDeposito = False
      initvalue
      dTgl.SetFocus
      Exit Function
    End If
    
    If GetNull(dbOpenDeposito!status) = "1" Then
      MsgBox "Rekening ini, " & cNomorRekeningDeposito & " Sudah Pencairan Pokok", vbInformation
      GetValidOpenDeposito = False
      initvalue
      dTgl.SetFocus
      Exit Function
    End If
  End If
End Function

Private Function GetValidOpenTabungan() As Boolean
Dim dbOpenTabungan As New ADODB.Recordset
Dim cNomorRekeningTabungan As String

  GetValidOpenTabungan = True
  cNomorRekeningTabungan = cCabang.Text & "." & cGolongan.Text & "." & cUrut.Text & "." & cFrekuensi.Text
  Set dbOpenTabungan = objData.Browse(GetDSN, "Tabungan", "StatusBlokir,Close", "Rekening", sisAssign, cNomorRekeningTabungan, " and Close = '1'")
  If Not dbOpenTabungan.eof Then
    MsgBox "Rekening ini, " & cNomorRekeningTabungan & " Sudah Tutup", vbInformation
    GetValidOpenTabungan = False
    initvalue
  End If
End Function

Private Sub GetDataDeposito(ByVal cRekening As String)
Dim cField As String
  
  cField = "d.sistemARO,d.JumlahPerpanjangan,d.Tgl,d.JthTmp,d.lama,sukubunga,nominalDeposito,StatusPostingPokok,r.Nama,r.Alamat,d.Status,d.PersentaseFinalti"
  Set dbData = objData.Browse(GetDSN, "Deposito d", cField, "d.Rekening", sisAssign, cRekening, , , _
                              Array("Left Join GolonganDeposito g on g.Kode = d.GolonganDeposito", _
                                    "Left Join registerNasabah r on r.Kode = d.Kode"))
  If Not dbData.eof Then
    If GetNull(dbData!status) = "1" Then
      MsgBox "Rekening tsb sudah tutup..", vbInformation
      cGolongan.Text = ""
      cUrut.Default
      cFrekuensi.Default
      PilihSSTAB "T"
      cCabang.SetFocus
      Exit Sub
    End If
    If GetNull(dbData!SistemARO) = "Y" Then
      nARO.Caption = "ARO"
      nARO.Value = GetNull(dbData!jumlahperpanjangan)
    Else
      nARO.Caption = "Non ARO"
      nARO.Value = 0
    End If
    cNama.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    dValuta.Value = GetNull(dbData!Tgl)
    nLama.Value = GetNull(dbData!Lama)
    dTempo.Value = GetNull(dbData!jthtmp)
    nBunga.Value = GetNull(dbData!SukuBunga)
    nPersFinalti.Value = GetNull(dbData!PersentaseFinalti)
    nNominalDeposito.Value = GetNull(dbData!nominaldeposito)
    nNominalDeposito.BackColor = &H8000000F
    nNominalDeposito.Enabled = False
    BiSAFrame7.Enabled = True
    BiSAFrame7.Visible = True
    BiSAFrame9.Enabled = True
    BiSAFrame9.Visible = True
    lStatusNominal = False
    optAsalPencairan(0).Value = True
    optAsalPencairan(0).Enabled = False
    optAsalPencairan(1).Enabled = False
    cStatusPostingPokok = GetNull(dbData!StatusPostingPokok)
    If nNominalDeposito.Value <= 0 Then
      lStatusNominal = True
      nNominalDeposito.BackColor = &H80000005
      nNominalDeposito.Enabled = True
      BiSAFrame7.Enabled = False
      BiSAFrame7.Visible = False
      BiSAFrame9.Enabled = False
      BiSAFrame9.Visible = False
      optAsalPencairan(0).Enabled = True
      optAsalPencairan(1).Enabled = True
    End If
    OptCair_Click (0)
  End If
End Sub

Private Sub OptCair_Click(Index As Integer)
Dim nJmlFinalti As Double
Dim nBG As Double

  frmPesan.Visible = False
  If Index = 0 Then
    nPokok.Value = 0
    nFinalti.Value = 0
    nFinalti.Enabled = False
    nFinalti.BackColor = &H8000000F
    nDPMaterai.Value = 0
    nDPMaterai.Enabled = False
    nDPMaterai.BackColor = &H8000000F
    GetBunga SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  Else
    nPokok.Value = nNominalDeposito.Value
    nBahas.Value = 0
    nPajak.Value = 0
    nTotal.Value = 0
    nFinalti.Value = 0
    nFinalti.Enabled = True
    nFinalti.BackColor = &H80000005
    nDPMaterai.Value = 0
    nDPMaterai.Enabled = True
    nDPMaterai.BackColor = &H80000005
    'If DateAdd("d", 3, Date) < dTempo.Value Then
    'If DateAdd("d", -3, dTgl.Value) < dTempo.Value Then
    If dTgl.Value < dTempo.Value Then
      frmPesan.Visible = True
      'nBG = Round((nNominalDeposito.Value * 30 * nBunga.Value / 100) / 365)
      nBG = Round(nNominalDeposito.Value * nBunga.Value / 100)
      nFinalti.Value = Round(nPersFinalti.Value / 100 * nBG)
    End If
    nTotal.Value = nFinalti.Value + nPokok.Value + nDPMaterai.Value + nBahas.Value + nPajak.Value
  End If
End Sub

Private Sub OptCair_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub SimpanDeposito(ByVal cRekening As String)
Dim cRekeningDeposito As String
Dim cRekeningJT As String
Dim cRekeningKAS As String
Dim cRekeningFinalty As String
Dim cRekeningTitipanBunga As String
Dim cRekneningPajakBunga As String
Dim cRekeningMaterai As String
  
  Set dbData = objData.Browse(GetDSN, "GolonganDeposito", , "Kode", sisAssign, cGolongan.Text)
  If Not dbData.eof Then
    cRekeningDeposito = GetNull(dbData!RekeningAkuntansi, "")
    cRekeningJT = GetNull(dbData!RekeningJatuhtempo, "")
    cRekeningFinalty = GetNull(dbData!RekeningFinalty, "")
    cRekeningTitipanBunga = GetNull(dbData!Cadanganbunga, "")
    cRekneningPajakBunga = GetNull(dbData!RekeningPajakbunga, "")
    cRekeningMaterai = GetNull(dbData!rekeningmaterai, "")
  End If
  cRekeningKAS = GetKasTeller(cusername)
  
  cFaktur.Text = GetLastFaktur(fkt_Deposito, dTgl.Value, True)
  If lStatusNominal = True Then
    objData.Edit GetDSN, "Deposito", "Rekening='" & cRekening & "'", Array("NominalDeposito"), Array(nNominalDeposito.Value)
    GetSimpanMutasi cFaktur.Text, dTgl.Value, cRekening, trPembukaan, nNominalDeposito.Value, nPajak.Value, cusername, SNow
    If optAsalPencairan(1).Value = True Then
      UpdMutasiTabungan objData, aCfg(msKodeTransaksiPB), cFaktur.Text, dTgl.Value, GetRekTabungan2, nNominalDeposito.Value, True, "Pencairan Tabungan ke Deposito"
    End If
    UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningKAS, "Pembukaan Deposito a.n " & cNama.Text, nNominalDeposito.Value, 0, , SNow
        UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningDeposito, "Pembukaan Deposito a.n " & cNama.Text, 0, nNominalDeposito.Value, , SNow
    initvalue
    cCabang.SetFocus
    Exit Sub
  End If
  
  If OptCair(0).Value = True And nBahas.Value <= 0 Then
    MsgBox "Bunga kosong..., Penyimpanan dibatalkan", vbInformation
    cCabang.SetFocus
    Exit Sub
  End If
  
  If OptCair(1).Value = True And nPokok.Value <= 0 Then
    MsgBox "Pokok tidak ada..., Penyimpanan dibatalkan", vbInformation
    cCabang.SetFocus
    Exit Sub
  End If
  
  'procedure
  'pembatalan pencairan pokok atau bunga
  
  ' mutasi tabungan
  ' mutasi deposito
  ' bunga deposito
  ' jika transaksi ini adalah pencairan pokok, maka update table deposito - ubah status = 0
  ' bukubesar
  
  
    '=====================
    'JIKA PENCAIRAN BUNGA
    '=====================
    If OptCair(0).Value = True Then
      GetSimpanMutasi cFaktur.Text, dTgl.Value, cRekening, trPencairanBunga, nBahas.Value - nPajak.Value, nPajak.Value, cusername, SNow
      If optTujuanPencairan(1).Value = False Then
        'Pencairan tunai
        UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningTitipanBunga, "Pencairan Bunga Deposito a.n " & cNama.Text, nTotal.Value, 0, , SNow
          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningKAS, "Pencairan Bunga Deposito a.n " & cNama.Text, 0, nTotal.Value, , SNow
'          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningKAS, "Pencairan Bunga Deposito a.n " & cNama.Text, 0, nBahas.Value - nPajak.Value, , SNow
'          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekneningPajakBunga, "Pajak Bunga Deposito a.n " & cNama.Text, 0, nPajak.Value, , SNow
      Else
        'Pemindahbukuan (ke Tabungan)
        UpdMutasiTabungan objData, aCfg(msKodeTransaksiPB), cFaktur.Text, dTgl.Value, GetRekTabungan, nTotal.Value, True, "Pencairan Bunga Deposito ke Tabungan"
        
'        UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningTitipanBunga, "Pencairan Bunga Deposito a.n " & cNama.Text, nBahas.Value, 0, , SNow
'          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, GetRekeningKodeTransaksi(aCfg(msKodeTransaksiPB)), "Pencairan Bunga Deposito a.n " & cNama.Text, 0, nBahas.Value - nPajak.Value, , SNow
'          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekneningPajakBunga, "Pajak Bunga Deposito a.n " & cNama.Text, 0, nPajak.Value, , SNow
        
      
      End If
      'hapus di mutasibungadeposito
      objData.Delete GetDSN, "MutasiBungaDeposito", "Rekening", sisAssign, cRekening
    '====================
    'JIKA PENCAIRAN POKOK
    '====================
    Else
      GetSimpanMutasi cFaktur.Text, dTgl.Value, cRekening, trPencairanPokok, nPokok.Value, nPajak.Value, cusername, SNow
      If nFinalti.Value > 0 Then
        GetSimpanMutasi cFaktur.Text, dTgl.Value, cRekening, trPinalti, nFinalti.Value, nPajak.Value, cusername, SNow
      End If
      
      If nMaterai.Value > 0 Then
        GetSimpanMutasi cFaktur.Text, dTgl.Value, cRekening, trMaterai, nDPMaterai.Value, nPajak.Value, cusername, SNow
      End If
      
      objData.Edit GetDSN, "deposito", "rekening='" & cRekening & "'", Array("Status", "Tglcair"), Array("1", dTgl.Value)
      If cStatusPostingPokok = "1" Then
        UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningJT, "Pencairan Pokok Deposito a.n " & cNama.Text, nNominalDeposito.Value, 0, "K", SNow
          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningKAS, "Pencairan Pokok Deposito a.n " & cNama.Text, , nPokok.Value - nFinalti.Value - nDPMaterai.Value, "K", SNow
          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningFinalty, "Finalty Pencairan Pokok Deposito a.n " & cNama.Text, 0, nFinalti.Value, "K", SNow
          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningMaterai, "Materai Pencairan Pokok Deposito a.n " & cNama.Text, 0, nDPMaterai.Value, "K", SNow
      Else
          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningDeposito, "Pencairan Pokok Deposito a.n " & cNama.Text, nNominalDeposito.Value, 0, "K", SNow
            UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningKAS, "Pencairan Pokok Deposito a.n " & cNama.Text, 0, nPokok.Value - nFinalti.Value - nDPMaterai.Value, "K", SNow
            UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningFinalty, "Finalty Pencairan Pokok Deposito a.n " & cNama.Text, 0, nFinalti.Value, "K", SNow
            UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningMaterai, "Materai Pencairan Pokok Deposito a.n " & cNama.Text, 0, nDPMaterai.Value, "K", SNow
      End If
      
      If optTujuanPencairan(1).Value = True Then
          UpdMutasiTabungan objData, aCfg(msKodeTransaksiPB), cFaktur.Text, dTgl.Value, GetRekTabungan, nTotal.Value, True, "Pencairan Pokok Deposito ke Tabungan"
          UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningJT, "Pencairan Pokok Deposito a.n " & cNama.Text, nNominalDeposito.Value, 0, "K", SNow
              UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, GetRekeningKodeTransaksi(aCfg(msKodeTransaksiPB)), "Pencairan Pokok Deposito a.n " & cNama.Text, , nPokok.Value - nFinalti.Value - nDPMaterai.Value, "K", SNow
              UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningFinalty, "Finalty Pencairan Pokok Deposito a.n " & cNama.Text, 0, nFinalti.Value, "K", SNow
              UpdKodeTr objData, msDeposito, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningMaterai, "Materai Pencairan Pokok Deposito a.n " & cNama.Text, 0, nDPMaterai.Value, "K", SNow
      End If
    End If
End Sub

Private Function GetRekeningKodeTransaksi(ByVal KodeTransaksi As String) As String
Dim db As New ADODB.Recordset

  GetRekeningKodeTransaksi = ""
  Set db = objData.Browse(GetDSN, "KodeTransaksi", "Kode,Rekening", "Kode", sisAssign, KodeTransaksi)
  If Not db.eof Then
    GetRekeningKodeTransaksi = GetNull(db!Rekening, "")
  End If
End Function

Private Sub GetBunga(ByVal cRekening As String)
'  Set dbData = objData.Browse(GetDSN, "BungaDeposito", "Sum(Bunga) as Bunga,Sum(Pajak) as Pajak", "Rekening", sisAssign, cRekening)
'  If Not dbData.eof Then
'    nBahas.Value = GetNull(dbData!bunga)
'    nPajak.Value = GetNull(dbData!Pajak)
'    nTotal.Value = nBahas.Value - nPajak.Value
'  End If

  Set dbData = objData.Browse(GetDSN, "MutasiBungaDeposito", "Sum(Jumlah) as Bunga,Sum(Pajak) as Pajak", "Rekening", sisAssign, cRekening)
  If Not dbData.eof Then
    nBahas.Value = GetNull(dbData!bunga)
    nPajak.Value = GetNull(dbData!Pajak)
    nTotal.Value = nBahas.Value - nPajak.Value
  End If
End Sub
Private Sub GetSimpanMutasi(ByVal cNomorFaktur As String, ByVal dTgl As Date, ByVal cRekening As String, ByVal cKodePencairan As trDeposito, ByVal nJumlah As Double, ByVal nPajakValue As Double, ByVal cusername As String, ByVal dDateTime As Date)
Dim vaField
Dim vaValue
  
  If nJumlah > 0 Then
    vaField = Array("Faktur", "Tgl", "KodeMutasi", "Rekening", "Jumlah", "pajak", "UserName", "DateTime")
    vaValue = Array(cNomorFaktur, dTgl, cKodePencairan, cRekening, nJumlah, nPajakValue, cusername, dDateTime)
    objData.Add GetDSN, "MutasiDeposito", vaField, vaValue
  End If
End Sub

Private Sub nFinalti_Change()
  nTotal.Value = nPokok.Value - nFinalti.Value - nDPMaterai.Value
End Sub

Private Sub GetDataPencairan(ByVal cSpk As String, ByVal dTglRealisasi As Date, ByVal nSukuBunga As Double, ByVal nLama As Double, _
                             ByVal nPlafond As Double, ByVal nAdm As Double, ByVal nMat As Double, ByVal cNM As String, ByVal cAlm As String, ByVal nProv As Double, ByVal nNot As Double, ByVal nBiayalain As Double, ByVal SimpananWajib As Double)
    
  cNama.Text = cNM
  cAlamat.Text = cAlm
  cSpkCair.Text = cSpk
  dTglRealisasiCair.Value = dTglRealisasi
  nBungaCair.Value = nSukuBunga
  nLamaCair.Value = nLama
  nPlafondCair.Value = nPlafond
  nAdministrasi.Value = Round(nAdm / 100 * nPlafond)
  nProvisi.Value = Round(nProv / 100 * nPlafond)
  nMaterai.Value = nMat
  nNotaris.Value = nNot
  nLainLain.Value = nBiayalain
  nSimpananWajib.Value = SimpananWajib
  nTotalPencairan.Value = nPlafondCair.Value - nAdministrasi.Value - nMaterai.Value - nProvisi.Value - nNotaris.Value - nLainLain.Value - nSimpananWajib.Value
End Sub

Private Sub SimpanPencairanKredit()
Dim vaField
Dim vaValue
Dim cNorek As String
  
    cFaktur.Text = GetLastFaktur(fkt_Relisasi, dTgl.Value, True)
    cNorek = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)

    vaField = Array("Rekening", "Faktur", "Tgl", "Penarikan", "administrasi", "materai", "provisi", "notaris", "biayalain", "total", "UserName", "DateTime")
    vaValue = Array(cNorek, cFaktur.Text, dTgl.Value, nTotalPencairan.Value, nAdministrasi.Value, nMaterai.Value, nProvisi.Value, nNotaris.Value, nLainLain.Value, nPlafondCair.Value, cusername, SNow)
    objData.Update GetDSN, "PencairanKredit", "Rekening = '" & cNorek & "'", vaField, vaValue

    objData.Edit GetDSN, "Debitur", "Rekening = '" & cNorek & "'", Array("StatusPencairan", "Faktur"), Array("1", cFaktur.Text)
    UpdRekPencairanKredit
    'Kembalikan username kepada username yang login
    cusername = GetRegistry(reg_UserName)
    
    If MsgBox("Apakah bukti/kwitansi angsuran ingin dicetak?", vbYesNo + vbInformation) = vbYes Then
      rptPrintRealisasi.cNoValidasi = cFaktur.Text
      rptPrintRealisasi.cAnggota = cNama.Text
      rptPrintRealisasi.cNamaAnggota = cAlamat.Text
      rptPrintRealisasi.cKeterangan = "No Rek. " & SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text) & " ** Plafond: " & Format(nPlafondCair.Value, "###,###,###") & " ** Jangka Waktu: " & nLamaCair.Value & " Bulan" & " ** Bunga: " & nBungaCair.Value & "% pa"
      
      rptPrintRealisasi.nAdministrasi = nAdministrasi.Value
      rptPrintRealisasi.nProvisi = nProvisi.Value
      rptPrintRealisasi.nMaterai = nMaterai.Value
      rptPrintRealisasi.nNotaris = nNotaris.Value
      rptPrintRealisasi.nBiayalain = nLainLain.Value
      rptPrintRealisasi.nSimpananWajib = nSimpananWajib.Value
      rptPrintRealisasi.nTotalRealisasi = nTotalPencairan.Value
      rptPrintRealisasi.cNamaPeminjam = cNama.Text
      rptPrintRealisasi.dTgl = dTgl.Value
      
      Load rptPrintRealisasi
      rptPrintRealisasi.Show
    End If
End Sub

Private Sub UpdRekPencairanKredit()
Dim par As Single
Dim cRekeningKAS As String
Dim cRekeningAdministrasi As String
Dim cRekeningMaterai As String
Dim cRekeningProvisi As String
Dim cRekeningNotaris As String
Dim cRekeningBiayalain As String
Dim cRekeningKYD As String
Dim cRekeningSimpananWajib As String

  Set dbData = objData.Browse(GetDSN, "GolonganKredit", , "Kode", sisAssign, cGolongan.Text)
  If Not dbData.eof Then
    cRekeningKYD = GetNull(dbData!Rekening, "")
    cRekeningAdministrasi = GetNull(dbData!rekeningadministrasi, "")
    cRekeningMaterai = GetNull(dbData!rekeningmaterai, "")
    cRekeningProvisi = GetNull(dbData!rekeningprovisi, "")
    cRekeningNotaris = GetNull(dbData!RekeningNotaris, "")
    cRekeningBiayalain = GetNull(dbData!RekeningBiayalainLain)
    cRekeningSimpananWajib = GetNull(dbData!rekeningsimpananwajib, "")
  End If
  
  par = vbTrigger.msRealisasiKredit
  objData.Delete GetDSN, "BukuBesar", "Status", sisAssign, par, "and Faktur = '" & cFaktur.Text & "'"
  cRekeningKAS = cKasTeller
    UpdKodeTr objData, msRealisasiKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningKYD, "Pencairan Kredit an. " & cNama.Text, nPlafondCair.Value, 0, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningKAS, "Pencairan Kredit an. " & cNama.Text, 0, nTotalPencairan.Value, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningAdministrasi, "Adm. pencairan Kredit an. " & cNama.Text, 0, nAdministrasi.Value, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningMaterai, "Materai Pencairan Kredit an. " & cNama.Text, 0, nMaterai.Value, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningProvisi, "Provisi pencairan Kredit an. " & cNama.Text, 0, nProvisi.Value, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningNotaris, "Notaris Pencairan Kredit an. " & cNama.Text, 0, nNotaris.Value, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningBiayalain, "Biaya Lain Pencairan Kredit an. " & cNama.Text, 0, nLainLain.Value, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningSimpananWajib, "Simpanan Wajib Peminjam an. " & cNama.Text, 0, nSimpananWajib.Value, "K", SNow
      
End Sub

Private Sub GetDataAngsuran(ByVal cSpk As String, ByVal dTglRealisasi As Date, ByVal nSukuBunga As Double, ByVal nLama As Double, ByVal nPlafond As Double, _
                            ByVal cNM As String, ByVal cAlm As String, ByVal nPeriodeAngs As Integer, ByVal nMinPeriode As Integer, ByVal nKonp As Integer, _
                            ByVal nDendaTelat As Double, ByVal nSimpWajibPeminjam As Double)
Dim nTotalAngsur As Integer
Dim nAngsPokok As Double
Dim nAngsBunga As Double
Dim nTotalAngsPokok As Double
Dim nSisaAngsuranBunga As Double
Dim nSisaAngsuranPokok As Double
Dim dTglJatuhTempo As Date
Dim NoRek As String

  cNama.Text = cNM
  cAlamat.Text = cAlm
  cSpkAngsuran.Text = cSpk
  dTglRealisasiAngsuran.Value = dTglRealisasi
  nBungaAngsuran.Value = nSukuBunga
  nLamaAngsuran.Value = nLama
  nPlafondAngsuran.Value = nPlafond
  nPeriodeAngsuran.Value = nPeriodeAngs
  nMinimumPeriode.Value = nMinPeriode
  nKonpensasi.Value = nKonp
  nDendaKeterlamabatan = nDendaTelat
  nSimpananWajibPeminjam.Value = nSimpWajibPeminjam
  GetBukuAngsuran SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text), nTotalAngsur, nTotalAngsPokok
  NoRek = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  
  'nTotalAngsur = jumlah x angsuran
  'nTotalAngsPokok = total angsuran pokok
  
  nBakiDebet = nTotalAngsPokok
  nSisaPokok.Value = Round(nPlafondAngsuran.Value - nBakiDebet)
     
  If cCaraAngsuran.Text = "H" Then
    If GetValidPembayaranAngsuran Then
      
      GetNewBungaPokok nAngsBunga, nAngsPokok
      nAngsuranPokok.Value = Mod50(nAngsPokok + nPokokLalu.Value)
      nAngsuranBunga.Value = Mod50(nAngsBunga + nBungaLalu.Value)
      
      'jika pokok angsuran dibayarkan melebihi baki debet.
      If nAngsuranPokok.Value > nSisaPokok.Value Then
        nPokokLalu.Value = 0
        nAngsuranPokok.Value = nSisaPokok.Value
      End If
      dTglJatuhTempo = DateAdd("m", nLamaAngsuran.Value, dTglRealisasiAngsuran.Value)
      
      'jika tanggal pembayaran angsuran lewat tgl jatuh tempo
      'maka bebankan semua pokok angsuran pada transaksi sekarang
      If dTgl.Value > dTglJatuhTempo Then
        nPokokLalu.Value = Mod50(nSisaPokok.Value)
        nAngsuranPokok.Value = Mod50(nSisaPokok.Value)
        nBungaLalu.Value = GetBakiBunga(objData, NoRek)
        nAngsuranBunga.Value = nBungaLalu.Value
      End If
    End If
  End If

  If cCaraAngsuran.Text = "B" Then
    If UCase(lbCaraPerhitungan.Caption) = "MENURUN" Then
     GetAngsuranBungaPokok SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text), dTglRealisasiAngsuran.Value, nLamaAngsuran.Value, nPlafondAngsuran.Value, nBungaAngsuran.Value
    ElseIf UCase(lbCaraPerhitungan.Caption) = "FLAT" Then
     GetAngsuranBungaPokokFlat
    End If
  End If
End Sub

Private Function GetBakiBunga(ByVal obj As CodeSuiteLibrary.data, ByVal Rek As String) As Double
Dim db As New ADODB.Recordset
Dim nToBunga As Double
Dim nToBungaAngsuran As Double

  GetBakiBunga = 0
  nToBunga = 0
  nToBungaAngsuran = 0
  Set db = obj.Browse(GetDSN, "Debitur", "Rekening,TotalBunga", "Rekening", sisAssign, Rek)
  If Not db.eof Then
    nToBunga = GetNull(db!totalBunga)
  End If
  Set db = obj.Browse(GetDSN, "Angsuran", "sum(Bunga) as Bunga", "Rekening", sisAssign, Rek)
  If Not db.eof Then
    nToBungaAngsuran = GetNull(db!bunga)
  End If
  GetBakiBunga = nToBunga - nToBungaAngsuran
End Function

Private Sub GetGridPeriode()
Dim n As Single
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim a As Integer
  
  ClearGridPeriode
  dTglAwal = DateAdd("d", 1, dTglRealisasiAngsuran.Value)
  dTglAkhir = DateAdd("d", nKonpensasi.Value, DateAdd("m", 1, dTglRealisasiAngsuran.Value))
  For n = 1 To nLamaAngsuran.Value
    vaGrid.InsertRows vaGrid.UpperBound(1) + 1
    a = vaGrid.UpperBound(1)
    vaGrid(a, 0) = a + 1
    vaGrid(a, 1) = dTglAwal
    vaGrid(a, 2) = dTglAkhir
    dTglAwal = DateAdd("d", 1, dTglAkhir)
    dTglAkhir = DateAdd("m", 1, DateAdd("d", -1, dTglAwal))
  Next
  Set TDBGrid3.Array = vaGrid
  TDBGrid3.ReBind
  TDBGrid3.Refresh
End Sub

Private Function GetLastPeriodeKe(ByVal obj As CodeSuiteLibrary.data, ByVal Rekening As String, ByVal Ke As Integer) As Date
Dim db As New ADODB.Recordset
Dim n As Integer
Dim dTmp As Date

  Set db = obj.Browse(GetDSN, "debitur", "rekening,tgl,lama,konpensasitelat", "rekening", sisAssign, Rekening)
  If Not db.eof Then
    dTmp = DateAdd("d", GetNull(db!KonpensasiTelat), GetNull(db!Tgl))
    For n = 1 To Ke
      GetLastPeriodeKe = DateAdd("m", 1, dTmp)
      dTmp = GetLastPeriodeKe
    Next n
  End If
End Function

Private Sub ClearGridPeriode()
  vaGrid.ReDim 0, -1, 0, 2
  Set TDBGrid3.Array = vaGrid
  TDBGrid3.ReBind
  TDBGrid3.Refresh
End Sub

Private Function GetValidPembayaranAngsuran() As Boolean
  GetValidPembayaranAngsuran = True
  'jika pembayaran angsuran sebelum tgl realisasi
  If dTglRealisasiAngsuran.Value > dTgl.Value Then
    MsgBox "Tgl transaksi tidak valid, tgl transaksi (angsuran) sebelum tgl realisasi", vbExclamation, Me.Caption
    dTgl.SetFocus
    GetValidPembayaranAngsuran = False
  End If
End Function

Private Sub GetBukuAngsuran(ByVal cRek As String, ByRef nFrekuensiAngsuran As Integer, ByRef nTotalAngsuranPokok As Double)
Dim nTotal As Double
Dim n As Long
Dim nPokok As Double
Dim nBunga As Double
Dim nDenda As Double
Dim vaBuku As New XArrayDB
Dim nAngsur As Integer
  
  nPokok = 0
  nBunga = 0
  nDenda = 0
  nTotal = 0
  nAngsur = 0
  vaBuku.ReDim 0, -1, 0, 5
  Set dbData = objData.Browse(GetDSN, "Angsuran", , "Rekening", sisAssign, cRek, , "ID,Tgl")
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaBuku.InsertRows vaBuku.UpperBound(1) + 1
      n = vaBuku.UpperBound(1)
      vaBuku(n, 0) = n + 1
      vaBuku(n, 1) = GetNull(dbData!Tgl)
      vaBuku(n, 2) = GetNull(dbData!pokok)
      vaBuku(n, 3) = GetNull(dbData!bunga)
      vaBuku(n, 4) = GetNull(dbData!denda)
      vaBuku(n, 5) = GetNull(dbData!Total)
      
      nPokok = nPokok + vaBuku(n, 2)
      nBunga = nBunga + vaBuku(n, 3)
      nDenda = nDenda + vaBuku(n, 4)
      nTotal = nTotal + vaBuku(n, 5)
      nAngsur = nAngsur + 1
      dbData.MoveNext
    Loop
    nFrekuensiAngsuran = nAngsur
    nTotalAngsuranPokok = nPokok
    TDBGrid2.Columns(2).FooterText = Format(nPokok, "###,###,###,###,##0.00")
    TDBGrid2.Columns(3).FooterText = Format(nBunga, "###,###,###,###,##0.00")
    TDBGrid2.Columns(4).FooterText = Format(nDenda, "###,###,###,###,##0.00")
    TDBGrid2.Columns(5).FooterText = Format(nTotal, "###,###,###,###,##0.00")
    TDBGrid2.Array = vaBuku
    TDBGrid2.ReBind
  End If
End Sub

Private Sub GetBungaPokok(ByRef nAngBunga As Double, ByRef nAngPokok As Double)
Dim n As Single
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim nSukuBungaPerBulan As Double
Dim xArray As New XArrayDB
Dim dTanggalAwal As Date
Dim dTanggalAkhir As Date
Dim nBD As Double
Dim cRek As String
Dim nBulanKe As Integer
Dim nAngsPokokPerBulan As Double
Dim nPK As Double

  xArray.ReDim 0, nLamaAngsuran.Value, 0, 1
  dTglAwal = (DateAdd("d", 1, dTglRealisasiAngsuran.Value))
  dTglAkhir = (DateAdd("m", 1, dTglRealisasiAngsuran.Value))
  
  For n = 1 To nLamaAngsuran.Value
    xArray(n, 0) = dTglAwal
    xArray(n, 1) = dTglAkhir
    dTglAwal = (DateAdd("d", 1, dTglAkhir))
    dTglAkhir = (DateAdd("m", 1, dTglAwal))
  Next
  
  For n = 1 To xArray.UpperBound(1)
    If Between(dTgl.Value, xArray(n, 0), xArray(n, 1)) Then
      dTanggalAkhir = DateAdd("m", -1, xArray(n, 1)) - 1
      nBulanKe = n
      Exit For
    End If
  Next

  If nBulanKe <= 1 Then
    nBD = nPlafondAngsuran.Value
  Else
    cRek = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    Set dbData = objData.Browse(GetDSN, "Angsuran", "Sum(Pokok) as Pokok", "Rekening", sisAssign, cRek, "And tgl <='" & Format(dTanggalAkhir, "yyyy-mm-dd") & "' Group By Rekening", "Tgl")
    If Not dbData.eof Then
      nBD = nPlafondAngsuran.Value - GetNull(dbData!pokok)
    End If
  End If
  
  nAngsPokokPerBulan = Round(nPlafondAngsuran.Value / nLamaAngsuran.Value, 2)
  If nAngsPokokPerBulan > nBD Then
    nPK = nBD
  Else
    nPK = nPlafondAngsuran.Value
  End If
  nSukuBungaPerBulan = Round(nBungaAngsuran.Value / 12, 2)
  nAngPokok = Round((nPK / nLamaAngsuran.Value) / nPeriodeAngsuran.Value)
  nAngBunga = GetBungaReguler(nBD, nSukuBungaPerBulan)
  nAngBunga = Round(nAngBunga / nPeriodeAngsuran.Value)
  nPokokLalu.Value = GetTunggakanPokokHarian(nBulanKe, cRek)
  nBungaLalu.Value = GetTunggakanBungaHarian(nBulanKe, cRek)
End Sub

Private Sub GetNewBungaPokok(ByRef nAngBunga As Double, ByRef nAngPokok As Double)
Dim n As Single
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim nSukuBungaPerBulan As Double
Dim xArray As New XArrayDB
Dim dTanggalAwal As Date
Dim dTanggalAkhir As Date
Dim nBD As Double
Dim cRek As String
Dim nBulanKe As Integer
Dim nAngsPokokPerBulan As Double
Dim nPK As Double

  xArray.ReDim 0, nLamaAngsuran.Value, 0, 1
  dTglAwal = (DateAdd("d", 1, dTglRealisasiAngsuran.Value))
  dTglAkhir = (DateAdd("m", 1, dTglRealisasiAngsuran.Value))
  
  For n = 1 To nLamaAngsuran.Value
    xArray(n, 0) = dTglAwal
    xArray(n, 1) = dTglAkhir
    dTglAwal = (DateAdd("d", 1, dTglAkhir))
    dTglAkhir = (DateAdd("m", 1, dTglAwal))
  Next
  
  For n = 1 To xArray.UpperBound(1)
    If Between(dTgl.Value, xArray(n, 0), xArray(n, 1)) Then
      dTanggalAkhir = DateAdd("m", -1, xArray(n, 1)) - 1
      nBulanKe = n
      Exit For
    End If
  Next

  If nBulanKe <= 1 Then
    nBD = nPlafondAngsuran.Value
  Else
    cRek = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    Set dbData = objData.Browse(GetDSN, "Angsuran", "Sum(Pokok) as Pokok", "Rekening", sisAssign, cRek, "And tgl <='" & Format(dTanggalAkhir, "yyyy-mm-dd") & "' Group By Rekening", "Tgl")
    If Not dbData.eof Then
      nBD = nPlafondAngsuran.Value - GetNull(dbData!pokok)
    End If
  End If
  
  nAngsPokokPerBulan = Round(nPlafondAngsuran.Value / nLamaAngsuran.Value, 2)
  If nAngsPokokPerBulan > nBD Then
    nPK = nBD
  Else
    nPK = nPlafondAngsuran.Value
  End If
  
  nSukuBungaPerBulan = Round(nBungaAngsuran.Value / 12, 2)
  nAngPokok = Round((nPK / nLamaAngsuran.Value) / nPeriodeAngsuran.Value)
  nAngBunga = GetBungaReguler(nBD, nSukuBungaPerBulan)
  nAngBunga = Round(nAngBunga / nPeriodeAngsuran.Value)
  nPokokLalu.Value = GetTunggakanPokokHarian(nBulanKe, cRek)
  nBungaLalu.Value = GetTunggakanBungaHarian(nBulanKe, cRek)
End Sub

Private Function GetBungaReguler(ByVal nSisaPokok As Double, ByVal nBunga As Double) As Double
  GetBungaReguler = nSisaPokok * (nBunga / 100)
  GetBungaReguler = Mod50(GetBungaReguler)
End Function

Private Sub SimpanAngsuranKredit(ByVal cRek As String)
Dim vaField
Dim vaValue
Dim cRekeningPokok As String
Dim cRekeningBunga As String
Dim cRekeningDenda As String
Dim cRekeningSimpananWajib As String
Dim nSBunga As Double
Dim nSPokok As Double
  
  If validSimpanAngsuran Then
    cFaktur.Text = GetLastFaktur(fkt_Angsuran, dTgl.Value, True)
    Set dbData = objData.Browse(GetDSN, "GolonganKredit", , "Kode", sisAssign, cGolongan.Text)
    If Not dbData.eof Then
      cRekeningPokok = GetNull(dbData!RekeningAngsuranPokok, "")
      cRekeningBunga = GetNull(dbData!rekeningangsuranbunga, "")
      cRekeningDenda = GetNull(dbData!rekeningdenda, "")
      cRekeningSimpananWajib = GetNull(dbData!rekeningsimpananwajib, "")
    End If
    
    objData.Delete GetDSN, "Angsuran", "Faktur", sisAssign, cFaktur.Text
    vaField = Array("Faktur", "Tgl", "Rekening", "Pokok", "Bunga", "Denda", _
                    "Total", "DateTime", "UserName")
    vaValue = Array(cFaktur.Text, dTgl.Value, cRek, nAngsuranPokok.Value, nAngsuranBunga.Value, nDenda.Value, _
                    nKewajiban.Value, SNow, cusername)
    objData.Add GetDSN, "Angsuran", vaField, vaValue
    
    If cCaraAngsuran.Text = "B" Then
      objData.Delete GetDSN, "SisaBungaAngsuran", "Rekening", sisAssign, cRek
      nSBunga = nSisaAngsBunga - nAngsuranBunga.Value
      nSPokok = nSisaAngsPokok - nAngsuranPokok.Value
      If nSPokok > 0 Or nSBunga > 0 Then
        objData.Add GetDSN, "SisaBungaAngsuran", Array("Rekening", "tgl", "SisaBunga", "SisaPokok", "Username", "DateTime"), _
                                                 Array(cRek, dTgl.Value, nSBunga, nSPokok, cusername, SNow)
      End If
    End If
    
    objData.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msAngsuranKredit, "And Faktur='" & cFaktur.Text & "'"
    UpdKodeTr objData, msAngsuranKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cKasTeller, "Angsuran Kredit an. " & cNama.Text, nKewajiban.Value, 0, "K"
      UpdKodeTr objData, msAngsuranKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningPokok, "Angsuran Pokok Kredit an. " & cNama.Text, 0, nAngsuranPokok.Value, "K"
      UpdKodeTr objData, msAngsuranKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningBunga, "Angsuran Bunga Kredit an. " & cNama.Text, 0, nAngsuranBunga.Value, "K"
      UpdKodeTr objData, msAngsuranKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningDenda, "Denda Angsuran Kredit an. " & cNama.Text, 0, nDenda.Value, "K"
    
    If nSisaPokok.Value <= 50 Then
      'jika sudah lunas
      objData.Edit GetDSN, "Debitur", "Rekening='" & SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text) & "'", Array("Status"), Array("1")
      'kembalikan simpanan wajib mereka
      If MsgBox("Apakah simpanan wajib peminjam akan dikembalikan?", vbYesNo + vbInformation) = vbYes Then
        'simpanan wajib (D)
          'kas          (K)
        UpdKodeTr objData, msAngsuranKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cRekeningSimpananWajib, "Pengembalian simpanan wajib peminjam an. " & cNama.Text, nSimpananWajibPeminjam.Value, 0, "K", SNow
           UpdKodeTr objData, msAngsuranKredit, cCabang.Text, cFaktur.Text, dTgl.Value, cKasTeller, "Pengembalian simpanan wajib peminjam an. " & cNama.Text, 0, nSimpananWajibPeminjam.Value, "K", SNow

      End If
    End If
    If MsgBox("Apakah bukti/kwitansi angsuran ingin dicetak?", vbYesNo + vbInformation) = vbYes Then
      trPrintKwitansiAngsuran.cNoValidasi = cFaktur.Text
      trPrintKwitansiAngsuran.cAnggota = cNama.Text
      trPrintKwitansiAngsuran.cNamaAnggota = cAlamat.Text
      trPrintKwitansiAngsuran.cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
      trPrintKwitansiAngsuran.nAngsuranBunga = nAngsuranBunga.Value
      trPrintKwitansiAngsuran.nAngsuranDenda = nDenda.Value
      trPrintKwitansiAngsuran.nAngsuranPokok = nAngsuranPokok.Value
      trPrintKwitansiAngsuran.nJumlahAngsuran = nKewajiban.Value
      trPrintKwitansiAngsuran.cNamaPeminjam = cNama.Text
      trPrintKwitansiAngsuran.nBakiDebet = nSisaPokok.Value
      trPrintKwitansiAngsuran.dTgl = dTgl.Value
      Load trPrintKwitansiAngsuran
      trPrintKwitansiAngsuran.Show
    End If
  End If
End Sub

Private Function validSimpanAngsuran() As Boolean
  validSimpanAngsuran = True
  If nKewajiban.Value <= 0 Then
    validSimpanAngsuran = False
  End If
End Function

Private Function CekTeller(ByVal cNomorRekening As String) As Boolean
  CekTeller = True
  Set dbData = objData.Browse(GetDSN, "username", "KasTeller", "username", sisAssign, GetRegistry(reg_UserName))
  If Not dbData.eof Then
    If GetNull(dbData!KasTeller) = "0" Then
      CekTeller = False
    End If
  Else
    CekTeller = False
  End If
End Function

Private Function CekRekening(ByVal cNomorRekening As String) As Boolean
Dim cTabel As String

  CekRekening = True
  Select Case left(cJenisProduk.Text, 1)
    Case Is = "T"
      cTabel = "Tabungan"
    Case Is = "D"
      cTabel = "Deposito"
    Case Is = "K"
      cTabel = "Debitur"
  End Select
  
  Set dbData = objData.Browse(GetDSN, cTabel, "Rekening", "Rekening", sisAssign, cNomorRekening)
  If dbData.eof Then
    CekRekening = False
  End If
End Function

Private Function ValidSimpan() As Boolean
  ValidSimpan = True
  If Not CheckData(cCabang.Text, "Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSimpan = False
    cCabang.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cGolongan.Text, "Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSimpan = False
    cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cUrut.Text, "Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSimpan = False
    cUrut.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Rekening tidak valid, Ulangi Pengisian.....!") Then
    ValidSimpan = False
    cFrekuensi.SetFocus
    Exit Function
  End If
  
'  Dim dbCheck As New ADODB.Recordset
'  Dim cMsg As String
'
'  cMsg = "Transaksi sudah mencapai batas yang telah ditentukan"
'  Set dbCheck = objData.Browse(GetDSN, "mutasitabungan", "count(faktur) as limitrecord")
'  If Not dbCheck.eof Then
'    If GetNull(dbCheck!limitrecord) > 100 Then
'      MsgBox cMsg, vbCritical
'      Unload Me
'      End
'    End If
'  End If
'  Set dbCheck = objData.Browse(GetDSN, "mutasideposito", "count(faktur) as limitrecord")
'  If Not dbCheck.eof Then
'    If GetNull(dbCheck!limitrecord) > 100 Then
'      MsgBox cMsg, vbCritical
'      Unload Me
'      End
'    End If
'  End If
'  Set dbCheck = objData.Browse(GetDSN, "angsuran", "count(faktur) as limitrecord")
'  If Not dbCheck.eof Then
'    If GetNull(dbCheck!limitrecord) > 100 Then
'      MsgBox cMsg, vbCritical
'      Unload Me
'      End
'    End If
'  End If
End Function

Private Sub initvalue()
  cFaktur.Default
  dTgl.Value = Date
  cNama.Default
  cAlamat.Default
  cJenisProduk.Default
  Image1.Picture = LoadPicture(GetPicture(""))
  Image2.Picture = LoadPicture(GetPicture(""))
  cCabang.Text = aCfg(msKodeCabang)
  cGolongan.Text = ""
  cUrut.Default
  cFrekuensi.Default
  cGolTabungan.Default
  cNamaGolTabungan.Default
  nSaldoMinimum.Value = 0
  nSetoranMinimum.Value = 0
  cKodeTransaksi.Default
  cNamaKodeTransaksi.Default
  cDK.Default
  cRekeningJurnal.Default
  cNamaRekeningJurnal.Default
  cKeteranganTabungan.Default
  nAwal.Value = 0
  nMutasi.Value = 0
  nAkhir.Value = 0
  Frameblokir.Visible = False
  dValuta.Value = Date
  nLama.Value = 0
  dTempo.Value = Date
  nBunga.Value = 0
  nPersFinalti.Value = 0
  nNominalDeposito.Value = 0
  nARO.Default
  ccTab1.Default
  ccTab2.Default
  ccTab3.Default
  ccTab4.Default
  cKodeTransaksiDepositoTujuanPencairan.Default
  nSaldoTabungan.Value = 0
  nFinalti.Value = 0
  nPokok.Value = 0
  nDPMaterai.Value = 0
  nBahas.Value = 0
  nPajak.Value = 0
  nTotal.Value = 0
  frmPesan.Visible = False
  cSpkCair.Default
  dTglRealisasiCair.Value = Date
  nBungaCair.Value = 0
  nLamaCair.Value = 0
  nPlafondCair.Value = 0
  nAdministrasi.Value = 0
  nMaterai.Value = 0
  nProvisi.Value = 0
  nNotaris.Value = 0
  nTotalPencairan.Value = 0
  nLainLain.Value = 0
  nPeriodeAngsuran.Value = 0
  nMinimumPeriode.Value = 0
  nKonpensasi.Value = 0
  cSpkAngsuran.Default
  dTglRealisasiAngsuran.Value = Date
  nPlafondAngsuran.Value = 0
  nBungaAngsuran.Value = 0
  nLamaAngsuran.Value = 0
  nPokokLalu.Value = 0
  nAngsuranPokok.Value = 0
  nAngsuranBunga.Value = 0
  nDenda.Value = 0
  nKewajiban.Value = 0
  nSisaPokok.Value = 0
  nBakiDebet = 0
  xArray.Clear
  xArray.ReDim 0, -1, 0, 5
  Set TDBGrid2.Array = xArray
  TDBGrid2.ReBind
  ClearGridPeriode
  GetSaldoTeller
  SSTab1.TabEnabled(0) = False
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
  SSTab1.TabEnabled(3) = False
  SSTab1.Tab = 4
  optTujuanPencairan(0).Value = True
  GetTujuanPencairan False
  GetAsalPencairan False
  optAsalPencairan(0).Value = True
  cKodeTransaksiDeposito.Default
  lbCaraPerhitungan.Visible = False
End Sub

Private Sub GetJenisTransaksi()
  Select Case cJenisProduk.Text
    Case Is = "T"
      cFaktur.Text = GetLastFaktur(fkt_MutasiTabungan, dTgl.Value, False)
    Case Is = "D"
      cFaktur.Text = GetLastFaktur(fkt_Deposito, dTgl.Value, False)
    Case Is = "K"
      cFaktur.Text = GetLastFaktur(fkt_Angsuran, dTgl.Value, False)
  End Select
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me, True
  initvalue
  cCabang.Text = aCfg(msKodeCabang, "")
  OptCair(0).Value = True
  
  TabIndex dTgl, n
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cKodeTransaksi, n
  TabIndex cKeteranganTabungan, n
  TabIndex nMutasi, n
  TabIndex optAsalPencairan(0), n
  TabIndex optAsalPencairan(1), n
  TabIndex ccTab1, n
  TabIndex ccTab2, n
  TabIndex ccTab3, n
  TabIndex ccTab4, n
  TabIndex cKodeTransaksiDepositoTujuanPencairan, n
  TabIndex nNominalDeposito, n
  TabIndex cKodeTransaksiDeposito, n
  TabIndex optTujuanPencairan(0), n
  TabIndex optTujuanPencairan(1), n
  TabIndex cTab1, n
  TabIndex cTab2, n
  TabIndex cTab3, n
  TabIndex cTab4, n
  TabIndex OptCair(0), n
  TabIndex OptCair(1), n
  TabIndex nFinalti, n
  TabIndex nDPMaterai, n
  TabIndex nPlafondCair, n
  TabIndex nAngsuranPokok, n
  TabIndex nAngsuranBunga, n
  TabIndex nDenda, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdBatal, n
  SSTab1.Tab = 4
  GetDataProduk
  GetLock
End Sub

Private Sub PilihSSTAB(ByVal cJenis As String)
  FrameMutasiTabungan.Enabled = False
  FrameDeposito.Enabled = False
  FramePencairan.Enabled = False
  FrameAngsuran.Enabled = False
  Select Case cJenis
    Case "T"
      SSTab1.Tab = 0
      FrameMutasiTabungan.Enabled = True
      SSTab1.TabEnabled(0) = True
      '''''''''''''''''''''''''''
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = False
      SSTab1.TabEnabled(3) = False
    Case "D"
      SSTab1.Tab = 1
      FrameDeposito.Enabled = True
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = True
      SSTab1.TabEnabled(2) = False
      SSTab1.TabEnabled(3) = False
    Case "C"
      SSTab1.Tab = 2
      FramePencairan.Enabled = True
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = True
      SSTab1.TabEnabled(3) = False
    Case "A"
      SSTab1.Tab = 3
      FrameAngsuran.Enabled = True
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = False
  End Select
End Sub

Private Function GetAngsuranPeriodik(ByVal dTglRealisasi As Date, ByVal nJumlahPlafond As Double, ByVal nLamaAngs As Integer) As Double
Dim nAngsPokok As Double
Dim nJumlahHari As Double

  GetAngsuranPeriodik = 0
  nJumlahHari = DateDiff("d", dTglRealisasi, DateAdd("m", nLamaAngs, dTglRealisasi))
  GetAngsuranPeriodik = Round(nJumlahPlafond / nJumlahHari, 2)
End Function

Private Sub GetAngsuranMenurunPeriodik(ByVal nJumlahPlafond As Double, ByVal nLamaAngs As Double, ByVal dTglRealisasi As Date, ByVal nSukuBunga As Double, ByVal nPeriodik As Integer, _
                                            ByRef nAngsuranBunga As Double, ByRef nAngsuranPokok As Double)
Dim nSukuBungaPerBulan As Double
Dim dTglAkhirPeriodik As Date
Dim dTglJatuhTempo As Date
Dim nLamaAngsuran As Integer
Dim nTotalAngPokok As Double
Dim nAngsPokokPeriodik As Double
  
  dTglAkhirPeriodik = DateAdd("D", nPeriodik, dTglRealisasi)
  dTglJatuhTempo = DateAdd("m", nLamaAngs, dTglRealisasi)
  nLamaAngsuran = DateDiff("m", dTglAkhirPeriodik, dTglJatuhTempo)
  nSukuBungaPerBulan = Round(nSukuBunga / 12, 2)
  nTotalAngPokok = GetSaldoAngsuran(SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text))
  nAngsuranBunga = GetBungaReguler(nJumlahPlafond - nTotalAngPokok, nSukuBungaPerBulan)
  nAngsPokokPeriodik = GetAngsuranPeriodik(dTglRealisasi, nJumlahPlafond, nLamaAngs)
  nAngsuranPokok = (nJumlahPlafond - (nAngsPokokPeriodik * nPeriodik)) / nLamaAngsuran
End Sub


Private Function GetSaldoAngsuran(ByVal cRekening As String) As Double
  GetSaldoAngsuran = 0
  Set dbData = objData.Browse(GetDSN, "Angsuran", "Sum(Pokok) as AngsuranPokok", "Rekening", sisAssign, cRekening, "group By Rekening", "Rekening")
  If Not dbData.eof Then
    GetSaldoAngsuran = GetNull(dbData!angsuranpokok)
  End If
End Function

Private Sub GetAngsuranPokokBunga(ByVal nJumlahPlafond As Double, ByVal nLamaAngs As Double, ByVal nSukuBunga As Double, _
                                  ByRef nAngsuranBunga As Double, ByRef nAngsuranPokok As Double)
Dim nSukuBungaPerBulan As Double
Dim nTotalAngPokok As Double
  

  nSukuBungaPerBulan = Round(nSukuBunga / 12, 2)
  nTotalAngPokok = GetSaldoAngsuran(SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text))
  nAngsuranBunga = GetBungaReguler(nJumlahPlafond - nTotalAngPokok, nSukuBungaPerBulan)
  nAngsuranPokok = nJumlahPlafond / nLamaAngs
End Sub

Private Sub SisaAngsuranBunga(ByVal cRekening As String, ByRef nSisaBunga As Double, ByRef nSisaPokok As Double)
  nSisaBunga = 0
  nSisaPokok = 0
  Set dbData = objData.Browse(GetDSN, "SisaBungaAngsuran", "sum(SisaBunga) as SisaBunga,SUM(SisaPokok) as SisaPokok", "Rekening", sisAssign, cRekening)
  If Not dbData.eof Then
    nSisaBunga = GetNull(dbData!SisaBunga)
    nSisaPokok = GetNull(dbData!SisaPokok)
  End If
End Sub

Private Function GetDendaBulanan(ByVal cRek As String, ByVal nANGSBG As Double) As Double
Dim nFrek As Integer
Dim dTanggal As Date
Dim nTelat As Integer
  
  GetDendaBulanan = 0
  nFrek = 0
  Set dbData = objData.Browse(GetDSN, "Angsuran", "Count(Rekening) as Jumlah", "Rekening", sisAssign, cRek)
  If Not dbData.eof Then
    nFrek = GetNull(dbData!Jumlah)
  End If

  nFrek = nFrek + 1
  dTanggal = (DateAdd("m", nFrek, dTglRealisasiAngsuran.Value))
  dTanggal = DateAdd("d", nKonpensasi.Value, dTanggal)
  If dTgl.Value > dTanggal Then
    nTelat = DateDiff("d", dTanggal, dTgl.Value)
    GetDendaBulanan = Mod50((nDendaKeterlamabatan / 100 * nANGSBG) * nTelat)
  End If
End Function

Private Function GetTunggakanPokokHarian(ByVal nKe As Integer, ByVal cRekening As String)
Dim xArray As New XArrayDB
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim n As Integer
Dim nAngsBln As Double
Dim nTotalAngsur As Double
Dim nTotalAngsur1 As Double
Dim nSisa As Double
Dim cWhere As String
Dim nSum As Double

  If nKe > 1 Then
    nAngsBln = Round(nPlafondAngsuran.Value / nLamaAngsuran.Value, 2)
    dTglAkhir = DateAdd("m", nKe - 1, dTglRealisasiAngsuran.Value)
    nTotalAngsur = 0
    Set dbData = objData.Browse(GetDSN, "Angsuran", "Sum(Pokok) as AngsuranPokok", "Rekening", sisAssign, cRekening, "And Tgl <= '" & Format(dTglAkhir, "yyyy-mm-dd") & "'")
    If Not dbData.eof Then
      nTotalAngsur = GetNull(dbData!angsuranpokok)
    End If
    nSum = nAngsBln * (nKe - 1)
    nSisa = nSum - nTotalAngsur

    If nSisa > 0 Then
      dTglAwal = DateAdd("d", 1, dTglAkhir)
      Set dbData = objData.Browse(GetDSN, "Angsuran", "Sum(Pokok) as AngsuranPokok", "Rekening", sisAssign, cRekening, "And Tgl > '" & Format(dTglAkhir, "yyyy-mm-dd") & "' And Tgl <= '" & Format(dTgl.Value, "yyyy-mm-dd") & "' Group By rekening")
      If Not dbData.eof Then
        nTotalAngsur1 = GetNull(dbData!angsuranpokok)
      End If
      
      If (nTotalAngsur + nTotalAngsur1) >= (nSisa + nTotalAngsur) Then
        GetTunggakanPokokHarian = 0
      Else
        GetTunggakanPokokHarian = (nSum + nTotalAngsur1) - (nTotalAngsur1 + nTotalAngsur)
      End If
    End If
  End If
End Function

Private Function GetTunggakanBungaHarian(ByVal nKe As Integer, ByVal cRekening As String)
Dim xArray As New XArrayDB
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim n As Integer
Dim nTotalAngsur As Double
Dim nTotalAngsur1 As Double
Dim nSisa As Double
Dim cWhere As String
Dim nSum As Double
Dim nTotBunga As Double
Dim nBD As Double
Dim nBungaPerHari As Double
Dim nBunga1 As Double
Dim nBunga2 As Double

  If nKe > 1 Then
    nTotBunga = 0
    nBunga1 = 0
    nBunga2 = 0
    For n = 1 To nKe - 1
      nBD = nPlafondAngsuran.Value
      If n > 1 Then
        nBD = GetBD(n, cRekening)
      End If
      nTotBunga = nTotBunga + GetBungaReguler(nBD, nBungaAngsuran.Value / 12)
      nBungaPerHari = nTotBunga / nPeriodeAngsuran.Value
      nBunga1 = nBungaPerHari * nMinimumPeriode.Value
      nBunga2 = nBunga2 + nBunga1
    Next
    
    dTglAkhir = DateAdd("m", nKe - 1, dTglRealisasiAngsuran.Value)
    nTotalAngsur = 0
    
    Set dbData = objData.Browse(GetDSN, "Angsuran", "Sum(Bunga) as AngsuranBunga", "Rekening", sisAssign, cRekening, "And Tgl <= '" & Format(dTglAkhir, "yyyy-mm-dd") & "'")
    If Not dbData.eof Then
      nTotalAngsur = GetNull(dbData!AngsuranBunga)
    End If
    nSum = nBunga2
    nSisa = nSum - nTotalAngsur
    
    If nSisa > 0 Then
      dTglAwal = DateAdd("d", 1, dTglAkhir)
      Set dbData = objData.Browse(GetDSN, "Angsuran", "Sum(Bunga) as AngsuranBunga", "Rekening", sisAssign, cRekening, "And Tgl > '" & Format(dTglAkhir, "yyyy-mm-dd") & "' And Tgl <= '" & Format(dTgl.Value, "yyyy-mm-dd") & "'")
      If Not dbData.eof Then
        nTotalAngsur1 = GetNull(dbData!AngsuranBunga)
      End If
      If (nTotalAngsur + nTotalAngsur1) >= (nSisa + nTotalAngsur) Then
        GetTunggakanBungaHarian = 0
      Else
        GetTunggakanBungaHarian = (nSum + nTotalAngsur1) - (nTotalAngsur + nTotalAngsur1)
      End If
    End If
  End If
End Function

Private Function GetBD(ByVal nKe As Integer, ByVal cRekening As String) As Double
Dim dTglAwal As Date
Dim dTglAkhir As Date
  
  GetBD = 0
  dTglAkhir = DateAdd("m", nKe, dTglRealisasiAngsuran.Value)
  Set dbData = objData.Browse(GetDSN, "Angsuran", "Sum(Pokok) as Pokok", "Rekening", sisAssign, cRekening, "And Tgl <= '" & Format(dTglAkhir, "yyyy-mm-dd") & "'")
  If Not dbData.eof Then
    GetBD = nPlafondAngsuran.Value - GetNull(dbData!pokok)
  End If
End Function

Private Sub GetLock()
  cGolongan.Enabled = True
  cGolongan.Text = aCfg(msDefaultTeller)
  If aCfg(msLockTeller) = "1" Then
    cGolongan.Enabled = False
  End If
  cJenisProduk.Text = left(cGolongan.Text, 1)
End Sub

Private Sub GetAngsuranBungaPokok(ByVal cRek As String, ByVal dTglValuta As Date, ByVal nLamaAngs As Integer, ByVal nJmlAngs As Double, ByVal nSBunga As Double)
Dim nKaliAngsur As Integer
Dim dJthTmpNow As Date
Dim nAngsurNow As Double
Dim nBakiDebet As Double
Dim nTotalAngsur As Double
Dim dTglTerakhir As Date
Dim nAngPKPerBulan As Double
Dim nTotalSudahAngsur As Double
Dim nBesarBungaNow As Double
Dim nTunda As Integer
Dim nKe As Integer
  
  nAngPKPerBulan = nJmlAngs / nLamaAngs
  nKe = GetAngsKe(dTglValuta, nLamaAngs, dTgl.Value, nKonpensasi.Value)
  
  'Hitung total PK yg sudah diangsur
  nTotalSudahAngsur = 0
  Set dbData = objData.Browse(GetDSN, "Angsuran", "sum(Pokok) as Pokok", "Rekening", sisAssign, cRek)
  If Not dbData.eof Then
    nTotalSudahAngsur = GetNull(dbData!pokok)
  End If
  
  'Total Angsuran Sekarang
  nPokokLalu.Value = (IIf(nKe = 0, nLamaAngs, (nKe - 1)) * nAngPKPerBulan) - nTotalSudahAngsur
  nPokokLalu.Value = IIf(nPokokLalu.Value < 0, 0, nPokokLalu.Value)
  nAngsuranPokok.Value = (nKe * nAngPKPerBulan) - nTotalSudahAngsur
  nBakiDebet = nJmlAngs - nTotalSudahAngsur
  
  'Angsuran Bunga Sekarang
  nTunda = GetTunda(cRek, dTglValuta, dTgl.Value)
  nBesarBungaNow = Mod50(nBakiDebet * (nSBunga / 100 / 12))
  nAngsuranBunga.Value = nBesarBungaNow * nTunda
  
  nAngsuranBunga.Value = nBesarBungaNow
  
  If nTunda > 1 Then
    nBungaLalu.Value = nBesarBungaNow * (nTunda - 1)
  Else
    nBungaLalu.Value = 0
  End If
  
  nDenda.Value = GetDendaBulanan(cRek, nBesarBungaNow)
  If nAngsuranPokok.Value > nSisaPokok.Value Then
    nPokokLalu.Value = 0
    nAngsuranPokok.Value = nSisaPokok.Value
  End If
  If dTgl.Value > (DateAdd("m", nLamaAngs, dTglRealisasiAngsuran.Value)) Then
    nAngsuranPokok.Value = Mod50(nSisaPokok.Value)
  End If
  nSisaAngsBunga = nAngsuranBunga.Value
  nSisaAngsPokok = nAngsuranPokok.Value
End Sub

Private Sub GetAngsuranBungaPokokFlat()
'Cek tgl sekarang termasuk periode ke berapa?
'Pada periode tersebut berapa angsuran pokok/bunga seharusnya
'Cek pada kenyataannya di database, nilali pokok/bunga seluruhnya berapa
'Jika kenyataanya kurang maka terjadi late pokok/bunga
'Jika kenyataannya lebih maka cek lagi apakah kredit ini sudah tutup atau belum
'Jika belum tutup maka pada tagihan pokok/bunga munculkan tagihan regulernya

Dim nPeriode As Integer 'periode sebelumnya
Dim nX As Integer 'periode sekarang
Dim nLate As Integer
Dim nPokokAng As Double
Dim nBungaAng As Double
Dim nPokok As Double
Dim nBunga As Double
Dim cNomorRekening As String
Dim nTotBunga As Double
Dim db As New ADODB.Recordset

  If dTgl.Value > dTglRealisasiAngsuran.Value Then
    nPokokLalu.Default
    nBungaLalu.Default
    cNomorRekening = cCabang.Text & "." & cGolongan.Text & "." & cUrut.Text & "." & cFrekuensi.Text
    nX = fGetPeriode(objData, cNomorRekening, dTgl.Value, nPeriode, nLate)
    fGetBungaPokokPeriodeKe objData, cNomorRekening, nPeriode, nPokokAng, nBungaAng
    fGetBungaPokok objData, cNomorRekening, nPokok, nBunga
    nAngsuranPokok.Value = (Devide(nPlafondAngsuran.Value, nLamaAngsuran.Value))
    nAngsuranBunga.Value = (nPlafondAngsuran.Value * nBungaAngsuran.Value / 12 / 100)
    If nPokok < nPokokAng Or nBunga < nBungaAng Then
      nPokokLalu.Value = (nPokokAng) - (nPokok)
      nBungaLalu.Value = (nBungaAng) - (nBunga)
    End If
  End If
End Sub

Private Function isInPeriodeAngsuran(ByVal obj As CodeSuiteLibrary.data, ByVal Rekening As String) As Boolean
Dim db As New ADODB.Recordset
isInPeriodeAngsuran = False

  Set db = obj.Browse(GetDSN, "Debitur", , "Rekening", sisAssign, Rekening)
  If Not db.eof Then
  End If
End Function

''=========================
'Private Sub GetAngsuranBungaPokok(ByVal cRek As String, ByVal dTglValuta As Date, ByVal nLamaAngs As Integer, ByVal nJmlAngs As Double, ByVal nSBunga As Double)
'Dim nKaliAngsur As Integer
'Dim dJthTmpNow As Date
'Dim nAngsurNow As Double
'Dim nBakiDebet As Double
'Dim nTotalAngsur As Double
'Dim dTglTerakhir As Date
'Dim nAngPKPerBulan As Double
'Dim nTotalSudahAngsur As Double
'Dim nBesarBungaNow As Double
'Dim nTunda As Integer
'Dim nKe As Integer
'Dim nXAngsuranSeharusnya As Double
'Dim nCount As Double
'
'  nAngPKPerBulan = nJmlAngs / nLamaAngs
'  nKe = GetAngsKe(dTglValuta, nLamaAngs, dTgl.Value, nKonpensasi.Value)
'
'  'Hitung total PK yg sudah diangsur
'  nTotalSudahAngsur = 0
'  Set dbData = objData.Browse(GetDSN, "Angsuran", "sum(Pokok) as Pokok", "Rekening", sisAssign, cRek)
'  If Not dbData.eof Then
'    nTotalSudahAngsur = GetNull(dbData!Pokok)
'  End If
'
'  Set dbData = objData.Browse(GetDSN, "Angsuran", "count(pokok) as Pokok", "Rekening", sisAssign, cRek)
'  If Not dbData.eof Then
'    nCount = GetNull(dbData!Pokok)
'  End If
'
'  'Total Angsuran Sekarang Total.
'  nPokokLalu.Value = ((nKe - 1) * nAngPKPerBulan) - nTotalSudahAngsur
'  nPokokLalu.Value = IIf(nPokokLalu.Value < 0, 0, nPokokLalu.Value)
'  nAngsuranPokok.Value = (nKe * nAngPKPerBulan) - nTotalSudahAngsur
'  nBakiDebet = nJmlAngs - nTotalSudahAngsur
'
'  'Masukkan pokok lalu
'  'by KODE
'  nXAngsuranSeharusnya = DateDiff("m", dTglValuta, dTgl.Value) + 1
'  nCount = nXAngsuranSeharusnya - nCount
'  nPokokLalu.Value = nCount * nAngPKPerBulan
'  nAngsuranPokok.Value = nBakiDebet + nPokokLalu.Value
'
'  'Masukkan pokok sekarang
'
'  'Angsuran Bunga Sekarang
'
'  nTunda = GetTunda(cRek, dTglValuta, dTgl.Value)
'  nBesarBungaNow = Mod50(nBakiDebet * (nSBunga / 100 / 12))
'  nAngsuranBunga.Value = nBesarBungaNow * nTunda
'  If nTunda > 1 Then
'    nBungaLalu.Value = nBesarBungaNow * (nTunda - 1)
'  Else
'    nBungaLalu.Value = 0
'  End If
'
'  nDenda.Value = GetDendaBulanan(cRek, nBesarBungaNow)
''  If nAngsuranPokok.Value > nSisaPokok.Value Then
''    nPokokLalu.Value = 0
''    nAngsuranPokok.Value = nSisaPokok.Value
''  End If
''
''  If dTgl.Value > (DateAdd("m", nLamaAngs, dTglRealisasiAngsuran.Value)) Then
''    nAngsuranPokok.Value = Mod50(nSisaPokok.Value)
''  End If
'
'  nSisaAngsBunga = nAngsuranBunga.Value
'  nSisaAngsPokok = nAngsuranPokok.Value
'End Sub

Private Sub GetNewAngsuranBungaPokok(ByVal cRek As String, ByVal dTglValuta As Date, ByVal nLamaAngs As Integer, ByVal nJmlAngs As Double, ByVal nSBunga As Double)
Dim nJumlahAngsuran As Double
Dim nXseharusnya As Integer
Dim nBakiDebet As Double
Dim nAngsuran As Double
Dim nSisaPokok As Double

  Set dbData = objData.Browse(GetDSN, "Angsuran", "count(rekening) as jmlAngsuran", "Rekening", sisAssign, cRek)
  If Not dbData.eof Then
    nJumlahAngsuran = GetNull(dbData!jmlAngsuran)
  End If
  Set dbData = objData.Browse(GetDSN, "Angsuran", "sum(pokok) as Angsuran", "Rekening", sisAssign, cRek)
  If Not dbData.eof Then
    nAngsuran = GetNull(dbData!angsuran)
  End If
  nXseharusnya = DateDiff("m", dTglValuta, dTgl) + 1
  Set dbData = objData.Browse(GetDSN, "Debitur", "Plafond", "Rekening", sisAssign, cRek)
  If Not dbData.eof Then
    nSisaPokok = GetNull(dbData!plafond) - nJumlahAngsuran
  End If
  nAngsuranPokok.Value = nSisaPokok / (nLamaAngs - nJumlahAngsuran)
  nPokokLalu.Value = 0
End Sub

Private Function GetJthTmp(ByVal dVlt As Date, ByVal nKe As Integer) As Date
Dim n As Integer
Dim dTemp As Date

  dTemp = dVlt
  For n = 1 To nKe
    dTemp = DateAdd("m", n, dTemp)
  Next
  GetJthTmp = dTemp
End Function

Private Function GetTunda(ByVal cRekening As String, ByVal dValuta As Date, ByVal dNow As Date) As Integer
Dim lStop  As Boolean
Dim dAkhir As Date
Dim dAwal As Date
Dim cWhere As String
Dim nKe As Integer
Dim nCount As Integer
  
  nKe = 1
  nCount = 0
  dAwal = dValuta - 1
  dAkhir = DateAdd("m", nKe, dValuta) + nKonpensasi.Value
  If dTgl.Value <= dAkhir Then
    GetTunda = 1
  Else
    lStop = False
    Do While lStop <> True
      cWhere = " And Tgl >='" & Format(dAwal, "yyyy-mm-dd") & "'"
      cWhere = cWhere & "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'"
      cWhere = cWhere & "And Bunga >0"
      Set dbData = objData.Browse(GetDSN, "Angsuran", "Rekening", "Rekening", sisAssign, cRekening, cWhere)
      If dbData.eof Then
        nCount = nCount + 1
      Else
        nCount = 0
      End If
      nKe = nKe + 1
      dAwal = dAkhir
      dAkhir = DateAdd("m", nKe, dValuta) + nKonpensasi.Value
      If dAkhir >= dTgl.Value Then
        lStop = True
      End If
    Loop
    GetTunda = IIf(nCount = 0, 0, nCount + 1)
  End If
End Function

'Private Function GetAngsKe(ByVal dVlt As Date, ByVal nLM As Integer, ByVal dDate As Date, ByVal nKonp As Integer) As Integer
'Dim n As Integer
'Dim dTemp As Date
'
'  dTemp = dVlt
'  For n = 1 To nLM
'    dTemp = DateAdd("m", 1, dTemp)
'    If dDate <= DateAdd("d", nKonp, dTemp) Then
'      GetAngsKe = n + 1
'      Exit For
'    End If
'  Next
'End Function

'Private Function GetAngsKe(ByVal dVlt As Date, ByVal nLM As Integer, ByVal dDate As Date, ByVal nKonp As Integer) As Integer
'Dim n As Integer
'Dim dTemp As Date
'
'  dTemp = dVlt
'  For n = 1 To nLM
'    dTemp = DateAdd("m", 1, dTemp)
'    If dDate <= DateAdd("d", nKonp, dTemp) Then
'      GetAngsKe = n
'      Exit For
'    End If
'  Next
'End Function

Private Function GetAngsKe(ByVal dVlt As Date, ByVal nLM As Integer, ByVal dDate As Date, ByVal nKonp As Integer) As Integer
Dim n As Integer
Dim dTemp As Date
Dim dTglRealisasi As Date
  
  dTemp = dVlt
  dTglRealisasi = dTemp
  For n = 1 To nLM
    dTemp = DateAdd("m", 1, dTemp)
    If dDate <= DateAdd("d", nKonp, dTemp) And dDate < dTglRealisasi Then
      GetAngsKe = -1
    ElseIf dDate <= DateAdd("d", nKonp, dTemp) Then
      GetAngsKe = n
      Exit For
    End If
  Next
End Function

Private Sub GetTujuanPencairan(Optional ByVal lStatus As Boolean = False)
  cTab1.Default
  cTab2.Default
  cTab3.Default
  cTab4.Default
  cKodeTransaksiDepositoTujuanPencairan.Default
  If lStatus = False Then
    cTab1.Enabled = False
    cTab2.Enabled = False
    cTab3.Enabled = False
    cTab4.Enabled = False
    cKodeTransaksiDepositoTujuanPencairan.Enabled = False
    cTab1.BackColor = vbButtonFace
    cTab2.BackColor = vbButtonFace
    cTab3.BackColor = vbButtonFace
    cTab4.BackColor = vbButtonFace
    cKodeTransaksiDepositoTujuanPencairan.BackColor = vbButtonFace
  Else
    cTab1.Enabled = True
    cTab2.Enabled = True
    cTab3.Enabled = True
    cTab4.Enabled = True
    cKodeTransaksiDepositoTujuanPencairan.Enabled = True
    cTab1.BackColor = vbHighlightText
    cTab2.BackColor = vbHighlightText
    cTab3.BackColor = vbHighlightText
    cTab4.BackColor = vbHighlightText
    cKodeTransaksiDepositoTujuanPencairan.BackColor = vbHighlightText
  End If
End Sub
