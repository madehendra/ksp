VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trCetakPinjaman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetakan Surat Perjanjian Pinjaman"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10155
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   660
      Left            =   0
      Top             =   6525
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   1164
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
         Left            =   8910
         TabIndex        =   34
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
         Picture         =   "trCetakPinjaman.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   6765
         TabIndex        =   35
         Top             =   105
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   767
         Caption         =   "     &Preview/ Cetak"
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
         Picture         =   "trCetakPinjaman.frx":00A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   6495
      Left            =   30
      Top             =   45
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   11456
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
      Begin SizerOneLibCtl.TabOne TabOne1 
         Height          =   5805
         Left            =   90
         TabIndex        =   1
         Top             =   585
         Width           =   9855
         _cx             =   17383
         _cy             =   10239
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   0
         FrontTabForeColor=   -2147483630
         Caption         =   "Rekening Realisasi|Data Pendamping (Jika Ada)|Yang Menyetujui"
         Align           =   0
         CurrTab         =   1
         FirstTab        =   0
         Style           =   4
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   0   'False
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   0   'False
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin SizerOneLibCtl.ElasticOne ElasticOne3 
            Height          =   5430
            Left            =   10500
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   330
            Width           =   9765
            _cx             =   17224
            _cy             =   9578
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   4
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   6
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            Begin BiSATextBoxProject.BiSATextBox cNamaYangMenyetujui 
               Height          =   330
               Left            =   510
               TabIndex        =   23
               Top             =   210
               Width           =   6030
               _ExtentX        =   10636
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
               Appearance      =   0
               Caption         =   "Nama"
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
            Begin BiSATextBoxProject.BiSATextBox cSelaku 
               Height          =   330
               Left            =   510
               TabIndex        =   24
               Top             =   585
               Width           =   4890
               _ExtentX        =   8625
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
               Appearance      =   0
               Caption         =   "Selaku"
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
         Begin SizerOneLibCtl.ElasticOne ElasticOne2 
            Height          =   5430
            Left            =   45
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   9765
            _cx             =   17224
            _cy             =   9578
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   4
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   6
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            Begin VB.CheckBox Check1 
               Caption         =   "Dengan Pendamping"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   540
               TabIndex        =   28
               Top             =   330
               Width           =   2040
            End
            Begin BiSATextBoxProject.BiSATextBox cNamaPendamping 
               Height          =   330
               Left            =   495
               TabIndex        =   25
               Top             =   765
               Width           =   6030
               _ExtentX        =   10636
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
               Appearance      =   0
               Caption         =   "Nama"
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
            Begin BiSATextBoxProject.BiSATextBox cAlamatPendamping 
               Height          =   330
               Left            =   495
               TabIndex        =   26
               Top             =   1140
               Width           =   7260
               _ExtentX        =   12806
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
            Begin BiSATextBoxProject.BiSATextBox cSelakuPendamping 
               Height          =   330
               Left            =   495
               TabIndex        =   27
               Top             =   1530
               Width           =   4890
               _ExtentX        =   8625
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
               Appearance      =   0
               Caption         =   "Selaku"
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
         Begin SizerOneLibCtl.ElasticOne ElasticOne1 
            Height          =   5430
            Left            =   -10410
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   9765
            _cx             =   17224
            _cy             =   9578
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   4
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   3
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   6
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            Begin VB.TextBox cJaminan 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1005
               Left            =   5460
               MultiLine       =   -1  'True
               TabIndex        =   38
               Text            =   "trCetakPinjaman.frx":032C
               Top             =   1665
               Width           =   4155
            End
            Begin BiSATextBoxProject.BiSATextBox cNama 
               Height          =   330
               Left            =   750
               TabIndex        =   5
               Top             =   690
               Width           =   6030
               _ExtentX        =   10636
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
               Appearance      =   0
               Caption         =   "Nama"
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
            Begin BiSATextBoxProject.BiSATextBox cAlamat 
               Height          =   330
               Left            =   750
               TabIndex        =   6
               Top             =   1065
               Width           =   7260
               _ExtentX        =   12806
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
            Begin BiSATextBoxProject.BiSATextBox cNoSPK 
               Height          =   330
               Left            =   750
               TabIndex        =   7
               Top             =   1965
               Width           =   4365
               _ExtentX        =   7699
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
               Appearance      =   0
               Caption         =   "No SPK"
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
               Left            =   750
               TabIndex        =   8
               Top             =   2340
               Width           =   2280
               _ExtentX        =   4022
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
               BackColor       =   -2147483633
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
            Begin BiSANumberBoxProject.BiSANumberBox nSukuBungaPerBulan 
               Height          =   330
               Left            =   750
               TabIndex        =   9
               Top             =   2715
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
               BackColor       =   -2147483633
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
            Begin BiSANumberBoxProject.BiSANumberBox nPlafond 
               Height          =   330
               Left            =   750
               TabIndex        =   11
               Top             =   3090
               Width           =   3435
               _ExtentX        =   6059
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
               BackColor       =   -2147483633
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
            Begin BiSADateProject.BiSADate dTglRealisasi 
               Height          =   330
               Left            =   735
               TabIndex        =   12
               Top             =   3465
               Width           =   2985
               _ExtentX        =   5265
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
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Caption         =   "Tgl realisasi"
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
            Begin BiSADateProject.BiSADate dTglJatuhTempo 
               Height          =   330
               Left            =   3795
               TabIndex        =   13
               Top             =   3450
               Width           =   2985
               _ExtentX        =   5265
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
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Caption         =   "Tgl jatuh tempo"
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
               Left            =   735
               TabIndex        =   14
               Top             =   3840
               Width           =   2475
               _ExtentX        =   4366
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
               BackColor       =   -2147483633
               Caption         =   "Provisi"
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
               Left            =   735
               TabIndex        =   15
               Top             =   4200
               Width           =   2475
               _ExtentX        =   4366
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
               BackColor       =   -2147483633
               Caption         =   "Administrasi"
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
            Begin BiSANumberBoxProject.BiSANumberBox nGracePeriod 
               Height          =   330
               Left            =   735
               TabIndex        =   18
               Top             =   4560
               Width           =   2475
               _ExtentX        =   4366
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
               BackColor       =   -2147483633
               Caption         =   "Grace Period"
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
            Begin BiSANumberBoxProject.BiSANumberBox nDenda 
               Height          =   330
               Left            =   735
               TabIndex        =   20
               Top             =   4920
               Width           =   2475
               _ExtentX        =   4366
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
               BackColor       =   -2147483633
               Caption         =   "Denda"
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
            Begin BiSAFramProject.BiSAFrame BiSAFrame4 
               Height          =   540
               Left            =   765
               Top             =   105
               Width           =   5325
               _ExtentX        =   9393
               _ExtentY        =   953
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
                  TabIndex        =   30
                  Top             =   75
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
                  TabIndex        =   31
                  Top             =   75
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
                  TabIndex        =   32
                  Top             =   75
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
                  TabIndex        =   33
                  Top             =   75
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
            Begin BiSATextBoxProject.BiSATextBox cNoKTP 
               Height          =   330
               Left            =   750
               TabIndex        =   37
               Top             =   1440
               Width           =   4365
               _ExtentX        =   7699
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
               Appearance      =   0
               Caption         =   "KTP"
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
               Caption         =   "Jaminan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   5430
               TabIndex        =   39
               Top             =   1395
               Width           =   1110
            End
            Begin VB.Label Label7 
               Caption         =   "Label7"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4455
               TabIndex        =   36
               Top             =   2745
               Width           =   1845
            End
            Begin VB.Label Label6 
               Caption         =   "Per Bulan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3300
               TabIndex        =   29
               Top             =   2745
               Width           =   1035
            End
            Begin VB.Label Label5 
               Caption         =   "% dari tunggakan pokok dan bunga"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3270
               TabIndex        =   21
               Top             =   4950
               Width           =   2745
            End
            Begin VB.Label Label4 
               Caption         =   "hari setelah jatuh tempo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3255
               TabIndex        =   19
               Top             =   4605
               Width           =   1890
            End
            Begin VB.Label Label3 
               Caption         =   "% dari plafond pinjaman"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3255
               TabIndex        =   17
               Top             =   4230
               Width           =   1890
            End
            Begin VB.Label Label2 
               Caption         =   "% dari plafond pinjaman"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3255
               TabIndex        =   16
               Top             =   3885
               Width           =   1890
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
               Height          =   285
               Left            =   3165
               TabIndex        =   10
               Top             =   2385
               Width           =   555
            End
         End
      End
      Begin BiSADateProject.BiSADate dTglCetak 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   90
         Width           =   2985
         _ExtentX        =   5265
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
         ForeColor       =   -2147483640
         Caption         =   "Tgl Cetak"
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
      Begin BiSATextBoxProject.BiSATextBox cKota 
         Height          =   330
         Left            =   3150
         TabIndex        =   22
         Top             =   75
         Width           =   3195
         _ExtentX        =   5636
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
         Appearance      =   0
         Caption         =   "Kota"
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
      Begin VB.PictureBox CR 
         Height          =   480
         Left            =   6585
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   40
         Top             =   105
         Width           =   1200
      End
   End
End
Attribute VB_Name = "trCetakPinjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cRekening As String

    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    Set dbData = objData.Browse(GetDSN, "Debitur", "Rekening,StatusPencairan", "Rekening", sisAssign, cRekening)
    If Not dbData.eof Then
      GetMemory
    End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "GolonganKredit", "Kode,Keterangan", "Kode", sisContent, cGolongan.Text)
  cGolongan.Text = cGolongan.Browse(dbData)
End Sub

Private Sub Check1_Click()
  If Check1.Value = 1 Then
    sShowPendamping True
  Else
    sShowPendamping False
  End If
End Sub

Private Sub sShowPendamping(cStatus As Boolean)
  cNamaPendamping.Visible = cStatus
  cAlamatPendamping.Visible = cStatus
  cSelakuPendamping.Visible = cStatus
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub
Private Sub GetMemory()
Dim n As Integer
Dim vaJoin
Dim cField As String
Dim cRekening As String

  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cField = "d.Tgl,d.faktur,d.NoSPk,d.SukuBunga,d.Plafond,d.Lama,d.AO,d.NoPengajuan,d.CaraAngsuran,d.PeriodeBungaMenurun,d.MinimalPeriode,d.KonpensasiTelat,d.BiayalainLain,d.DendaTelatBayar,d.wajibpokok,d.caraperhitungan,d.caraperhitungan,"
  cField = cField & " d.Wilayah,d.Administrasi,d.Materai,d.Provisi,d.Notaris,d.simpananwajib,"
  cField = cField & " d.Kode as KodeDebitur,r.Nama, r.Alamat,w.Keterangan as NamaWilayah,d.jatuhtempo,"
  cField = cField & " h.Nama as NamaAO,"
  cField = cField & " p.Nama as NamaPengajuan,p.Jaminan as JaminanPengajuan,p.Plafond as PlafondPengajuan"
  
  vaJoin = Array("Left Join Wilayah w on w.Kode = d.Wilayah", _
                 "Left Join RegisterNasabah r on r.Kode = d.Kode", _
                 "Left Join AO h on h.Kode = d.AO", _
                 "Left Join PengajuanKredit p on p.Kode = d.NoPengajuan")
                 
  Set dbData = objData.Browse(GetDSN, "Debitur d", cField, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    cCabang.Text = left(GetNull(dbData!KodeDebitur, ""), 2)
    dTglRealisasi.Value = GetNull(dbData!Tgl, "")
    nLama.Value = GetNull(dbData!Lama)
    cNama.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cNoSPK.Text = GetNull(dbData!NoSPK, "")
    nSukuBungaPerBulan.Value = GetNull(dbData!SukuBunga) / 12
    nPlafond.Value = GetNull(dbData!plafond)
    dTglJatuhTempo.Value = GetNull(dbData!JatuhTempo)
    nProvisi.Value = GetNull(dbData!Provisi)
    nAdministrasi.Value = GetNull(dbData!Administrasi)
    nGracePeriod.Value = GetNull(dbData!KonpensasiTelat)
    nDenda.Value = GetNull(dbData!DendaTelatBayar)
    
    Select Case GetNull(dbData!caraperhitungan)
      Case "1"
        Label7.Caption = "menurun"
      Case "2"
        Label7.Caption = "menetap"
    End Select
    
  End If
End Sub
Private Sub cmdPreview_Click()
With CR
'      .ReportFileName = App.Path & "\perjanjian pinjaman.rpt"
'
'      .ParameterFields(0) = "cNamaSelaku;" & cNamaYangMenyetujui.Text & ";True"
'      .ParameterFields(1) = "cSelaku;" & cSelaku.Text & ";True"
'      .ParameterFields(2) = "cNamaKoperasi;" & aCfg(msNama) & ";True"
'      .ParameterFields(3) = "cAlamatKoperasi;" & aCfg(msAlamat) & ";True"
'      .ParameterFields(4) = "cKotaKoperasi;" & aCfg(msKota) & ";True"
'      .ParameterFields(5) = "cNamaDebitur;" & cNama.Text & ";True"
'      .ParameterFields(6) = "cAlamatDebitur;" & cAlamat.Text & ";True"
'      .ParameterFields(7) = "cNoKTP;" & cNoKTP.Text & ";True"
'      .ParameterFields(8) = "cPlafond;" & Format(nPlafond.Value, "###,###,##0.00") & ";True"
'      .ParameterFields(9) = "cTerbilangPlafond;" & Dec2Text(nPlafond.Value) & ";True"
'      .ParameterFields(10) = "cLamaBulan;" & nLama.Value & ";True"
'      .ParameterFields(11) = "cTerbilangBulan;" & Label6.Caption & ";True"
'      .ParameterFields(12) = "cCaraPerhitungan;" & Label7.Caption & ";True"
'      .ParameterFields(13) = "cProvisi;" & nProvisi.Value & ";True"
'      .ParameterFields(14) = "cAdministrasi;" & nAdministrasi.Value & ";True"
'      .ParameterFields(15) = "cTglJatuhTempo;" & Format(dTglJatuhTempo.Value, "dd MMMM yyyy") & ";True"
'      .ParameterFields(16) = "nBunga;" & nSukuBungaPerBulan.Value & ";True"
'      .ParameterFields(17) = "cBulanDepan;" & Format(DateAdd("M", 1, dTglRealisasi.Value), "MMMM") & ";True"
'      .ParameterFields(18) = "cGracePeriod;" & nGracePeriod.Value & ";True"
'      .ParameterFields(19) = "nDenda;" & nDenda.Value & ";True"
'      .ParameterFields(20) = "cJaminan;" & cJaminan.Text & ";True"
'      .ParameterFields(22) = "cKota;" & cKota.Text & ";True"
'      .ParameterFields(23) = "dTglCetak;" & Format(dTglCetak.Value, "dd MMMM yyyy") & ";True"
'      .ParameterFields(24) = "dTglRealisasi;" & Format(dTglRealisasi.Value, "dd MMMM yyyy") & ";True"
'      .ParameterFields(25) = "cNoSPK;" & cNoSPK.Text & ";True"
'      .ParameterFields(28) = "cHeaderPendamping;" & " " & ";True"
'      .ParameterFields(26) = "cCaptionPendamping;" & " " & ";True"
'      .ParameterFields(27) = "cNamaPendamping;" & " " & ";True"
'
'
'      If Check1.Value = 1 Then
'        .ParameterFields(28) = "cHeaderPendamping;" & "III. " & cNamaPendamping.Text & " yang beralamat di " & cAlamatPendamping.Text & " selaku " & cSelakuPendamping.Text & " dari " & cNama.Text & ";True"
'        .ParameterFields(26) = "cCaptionPendamping;" & cSelakuPendamping.Text & ";True"
'        .ParameterFields(27) = "cNamaPendamping;" & cNamaPendamping.Text & ";True"
'
'      End If
'      .Action = 1
  End With
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  Check1.Value = 0
  sShowPendamping False
  TabOne1 = 0
  cCabang.Text = aCfg(msKodeCabang)
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  
  cJaminan.Text = ""
  cSelaku.Text = "Ketua"
  cNamaYangMenyetujui.Text = aCfg(msNamaDirut)
  TabIndex dTglCetak, n
  TabIndex cKota, n
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNoKTP, n
  TabIndex cJaminan, n
  TabIndex Check1, n
  TabIndex cNamaPendamping, n
  TabIndex cAlamatPendamping, n
  TabIndex cSelakuPendamping, n
  TabIndex cNamaYangMenyetujui, n
  TabIndex cSelaku, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub Text1_Change()

End Sub
