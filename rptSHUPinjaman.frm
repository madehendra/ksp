VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form rptSHUPinjaman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM SHU"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11190
   WindowState     =   2  'Maximized
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7110
      Width           =   11190
      _cx             =   19738
      _cy             =   556
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
      Align           =   2
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
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   270
         Left            =   75
         TabIndex        =   2
         Top             =   30
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   476
         Picture         =   "rptSHUPinjaman.frx":0000
         ForeColor       =   0
         BarPicture      =   "rptSHUPinjaman.frx":001C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   7110
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11190
      _cx             =   19738
      _cy             =   12541
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
      Align           =   5
      AutoSizeChildren=   7
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
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin SizerOneLibCtl.TabOne TabOne1 
         Height          =   7110
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   11190
         _cx             =   19738
         _cy             =   12541
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
         Appearance      =   2
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "POSTING|SHU ANGGOTA"
         Align           =   5
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin SizerOneLibCtl.ElasticOne ElasticOne6 
            Height          =   6735
            Left            =   11835
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   330
            Width           =   11100
            _cx             =   19579
            _cy             =   11880
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
            AutoSizeChildren=   7
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
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            Begin VB.CheckBox chkTampil 
               Caption         =   "Tampilkan Hanya yg MENDAPAT SHU"
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
               Left            =   75
               TabIndex        =   27
               Top             =   6465
               Width           =   4935
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Tampilkan Hanya yg TIDAK Mendapat SHU"
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
               Left            =   75
               TabIndex        =   26
               Top             =   6240
               Width           =   4890
            End
            Begin vbAcceleratorSGrid6.vbalGrid sgrid 
               Height          =   6165
               Left            =   45
               TabIndex        =   6
               Top             =   45
               Width           =   11010
               _ExtentX        =   19420
               _ExtentY        =   10874
               BackgroundPictureHeight=   0
               BackgroundPictureWidth=   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   2
               DisableIcons    =   -1  'True
            End
         End
         Begin SizerOneLibCtl.ElasticOne ElasticOne5 
            Height          =   6735
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   330
            Width           =   11100
            _cx             =   19579
            _cy             =   11880
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
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            Begin SizerOneLibCtl.ElasticOne ElasticOne8 
               Height          =   1335
               Left            =   6000
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   30
               Width           =   5040
               _cx             =   8890
               _cy             =   2355
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
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   ""
               Begin BiSANumberBoxProject.BiSANumberBox nPersenSHU1 
                  Height          =   300
                  Left            =   1110
                  TabIndex        =   30
                  Top             =   555
                  Width           =   900
                  _ExtentX        =   1588
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
                  Caption         =   " "
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
               Begin BiSANumberBoxProject.BiSANumberBox nPersenSHU2 
                  Height          =   300
                  Left            =   1110
                  TabIndex        =   31
                  Top             =   870
                  Width           =   900
                  _ExtentX        =   1588
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
                  Caption         =   " "
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
                  Caption         =   "%"
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
                  Left            =   2070
                  TabIndex        =   35
                  Top             =   930
                  Width           =   285
               End
               Begin VB.Label Label7 
                  Caption         =   "%"
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
                  Left            =   2085
                  TabIndex        =   34
                  Top             =   585
                  Width           =   285
               End
               Begin VB.Label Label6 
                  Caption         =   "SHU II"
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
                  Left            =   435
                  TabIndex        =   33
                  Top             =   885
                  Width           =   600
               End
               Begin VB.Label Label5 
                  Caption         =   "SHU I"
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
                  Left            =   435
                  TabIndex        =   32
                  Top             =   615
                  Width           =   600
               End
               Begin VB.Label Label4 
                  Caption         =   "Prosentase Bahas untuk SHU"
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
                  Left            =   435
                  TabIndex        =   29
                  Top             =   270
                  Width           =   2490
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "SHU II"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4305
               Left            =   6000
               TabIndex        =   20
               Top             =   1395
               Width           =   5055
               Begin SizerOneLibCtl.ElasticOne ElasticOne7 
                  Height          =   1560
                  Left            =   765
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   1245
                  Width           =   4230
                  _cx             =   7461
                  _cy             =   2752
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
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   ""
                  Begin BiSANumberBoxProject.BiSANumberBox nSimpananPokok 
                     Height          =   330
                     Left            =   90
                     TabIndex        =   24
                     Top             =   210
                     Width           =   4080
                     _ExtentX        =   7197
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
                     Caption         =   "Simp. Pokok"
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
                  Begin BiSANumberBoxProject.BiSANumberBox nSimpananWajib 
                     Height          =   330
                     Left            =   90
                     TabIndex        =   25
                     Top             =   570
                     Width           =   4080
                     _ExtentX        =   7197
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
                     Caption         =   "Simp. Wajib"
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
               Begin BiSANumberBoxProject.BiSANumberBox nBahas2 
                  Height          =   330
                  Left            =   105
                  TabIndex        =   21
                  Top             =   420
                  Width           =   3360
                  _ExtentX        =   5927
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
                  Caption         =   "Bahas"
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
                  Caption         =   "Jumlah Pengendapan"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   165
                  TabIndex        =   22
                  Top             =   990
                  Width           =   2010
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "SHU I"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4305
               Left            =   75
               TabIndex        =   13
               Top             =   1395
               Width           =   5925
               Begin SizerOneLibCtl.ElasticOne ElasticOne1 
                  Height          =   1560
                  Left            =   1170
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   1245
                  Width           =   4680
                  _cx             =   8255
                  _cy             =   2752
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
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   ""
                  Begin BiSANumberBoxProject.BiSANumberBox nKredit 
                     Height          =   330
                     Left            =   165
                     TabIndex        =   17
                     Top             =   240
                     Width           =   4395
                     _ExtentX        =   7752
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
                     Caption         =   "Pinjaman"
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
                  Begin BiSANumberBoxProject.BiSANumberBox nSimpanan 
                     Height          =   330
                     Left            =   165
                     TabIndex        =   18
                     Top             =   585
                     Width           =   4395
                     _ExtentX        =   7752
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
                     Caption         =   "Simp. Sukarela"
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
                  Begin BiSANumberBoxProject.BiSANumberBox nDeposito 
                     Height          =   330
                     Left            =   165
                     TabIndex        =   19
                     Top             =   945
                     Width           =   4395
                     _ExtentX        =   7752
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
                     Caption         =   "Simp. Berjangka"
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
               Begin BiSANumberBoxProject.BiSANumberBox nBahas 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   14
                  Top             =   450
                  Width           =   3360
                  _ExtentX        =   5927
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
                  Caption         =   "Bahas"
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
                  Caption         =   "Jumlah Pengendapan"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   150
                  TabIndex        =   15
                  Top             =   990
                  Width           =   2010
               End
            End
            Begin SizerOneLibCtl.ElasticOne ElasticOne4 
               Height          =   1335
               Left            =   30
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   30
               Width           =   5910
               _cx             =   10425
               _cy             =   2355
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
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   ""
               Begin VB.CommandButton cmdOK 
                  Caption         =   "Proses"
                  Height          =   330
                  Left            =   4650
                  TabIndex        =   8
                  Top             =   495
                  Width           =   945
               End
               Begin BiSADateProject.BiSADate dTgl 
                  Height          =   330
                  Index           =   0
                  Left            =   120
                  TabIndex        =   9
                  Top             =   495
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
                  Caption         =   "Periode"
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
                  Index           =   1
                  Left            =   2640
                  TabIndex        =   10
                  Top             =   495
                  Width           =   1980
                  _ExtentX        =   3493
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
                  Caption         =   "sd"
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
               Begin BiSANumberBoxProject.BiSANumberBox nTotalLabaRugi 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   12
                  Top             =   870
                  Width           =   3360
                  _ExtentX        =   5927
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
                  Caption         =   "Total Laba"
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
                  Caption         =   "SISA HASIL USAHA"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   165
                  TabIndex        =   11
                  Top             =   90
                  Width           =   5580
               End
            End
         End
      End
   End
End
Attribute VB_Name = "rptSHUPinjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.data

Private Sub BuatKolom()
  With sGrid
    .AddColumn , "KODE", ecgHdrTextALignCentre, , 80
    .AddColumn , "NAMA", ecgHdrTextALignCentre, , 200
    .AddColumn , "SIM SUKARELA", ecgHdrTextALignCentre, , 100, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "SIM BERJANGKA", ecgHdrTextALignCentre, , 100, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "PINJAMAN", ecgHdrTextALignCentre, , 100, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "%", ecgHdrTextALignCentre, , 50, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "SHU I", ecgHdrTextALignCentre, , 150, , , , , "###,###,##0.00", , CCLSortNumeric
  
    .AddColumn , "SIM POKOK", ecgHdrTextALignCentre, , 100, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "SIM WAJIB", ecgHdrTextALignCentre, , 100, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "%", ecgHdrTextALignCentre, , 50, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "SHU II", ecgHdrTextALignCentre, , 150, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "JUMLAH SHU", ecgHdrTextALignCentre, , 150, , , , , "###,###,##0.00", , CCLSortNumeric
  End With
End Sub

Private Sub Check1_Click()
Dim bS As Boolean
Dim lRow As Long

   bS = (Check1.Value = Unchecked)
   With sGrid
      .Redraw = False
      For lRow = 1 To .Rows
        If .CellText(lRow, 12) <> 0 Then
          .RowVisible(lRow) = bS
        End If
      Next lRow
      .Redraw = True
   End With
End Sub

Private Sub chkTampil_Click()
Dim bS As Boolean
Dim lRow As Long

   bS = (chkTampil.Value = Unchecked)
   With sGrid
      .Redraw = False
      For lRow = 1 To .Rows
        If .CellText(lRow, 12) = 0 Then
          .RowVisible(lRow) = bS
        End If
      Next lRow
      .Redraw = True
   End With
End Sub

Private Sub cmdOK_Click()
  initvalue
  GetData
End Sub

Private Sub Command1_Click()
  PostingSaldoTerendah objData, dTgl(1).Value
End Sub

Private Sub cmdPosting_Click()
End Sub

Private Sub Form_Load()
Dim n As Single
  CenterForm Me, True
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cmdOK, n
  TabIndex nKredit, n
  TabIndex nTotalLabaRugi, n
  TabIndex nSimpanan, n
  TabIndex nBahas, n
  
  CenterForm Me
  initvalue
  InitGrid sGrid
  BuatKolom
End Sub

Private Sub initvalue()

  TabOne1 = 0
  nTotalLabaRugi.Default
  
  nPersenSHU1.Value = 25
  nPersenSHU2.Value = 20
  
  nBahas.Default
  nKredit.Default
  nSimpanan.Default
  nDeposito.Default
  
  nBahas2.Default
  nSimpananPokok.Default
  nSimpananWajib.Default
  
  chkTampil.Value = Unchecked
  vbalProgressBar1.Visible = False
End Sub

Sub InitGrid(vbagrid As vbalGrid)
  With vbagrid
    .GridLines = False
    .AlternateRowBackColor = RGB(252, 252, 230)
    .RowMode = True
    .NoVerticalGridLines = True
    .DrawFocusRectangle = False
    .SelectionAlphaBlend = True
    .SelectionOutline = True
  End With
End Sub

Private Function GetTotalSimpananMengendap(ByVal obj As CodeSuiteLibrary.data, ByVal cGolonganTabungan As String) As Double
Dim db As New ADODB.Recordset

GetTotalSimpananMengendap = 0
  
  Set db = obj.Browse(GetDSN, "simpananmengendap s", "sum(s.jumlah) as jumlah", "g.kode", sisAssign, cGolonganTabungan, , , Array("left join tabungan t on t.rekening = s.rekening", "left join golongantabungan g on g.kode = t.golongantabungan"))
  If Not db.eof Then
    GetTotalSimpananMengendap = GetNull(db!Jumlah)
  End If
End Function

Private Function GetSimpananMengendapAnggota(ByVal obj As CodeSuiteLibrary.data, ByVal cRegister As String, ByVal cGolonganTabungan As String, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
Dim db As New ADODB.Recordset

  GetSimpananMengendapAnggota = 0
  Set db = obj.Browse(GetDSN, "simpananmengendap", "sum(jumlah) as jumlah", "Kode", sisAssign, cRegister, " and golongantabungan = '" & cGolonganTabungan & "' and tahun >= '" & Year(dTgl(0).Value) & "' and bulan >='" & Month(dTgl(0).Value) & "' and tahun <= '" & Year(dTgl(1).Value) & "' and bulan <= '" & Month(dTgl(1).Value) & "'")
  If Not db.eof Then
    GetSimpananMengendapAnggota = GetNull(db!Jumlah)
  End If
End Function

Private Sub GetData()
Dim db As New ADODB.Recordset

  Me.MousePointer = vbHourglass
  nTotalLabaRugi.Value = GetRugiLabaSHU(objData, dTgl(0).Value, dTgl(1).Value)
  nKredit.Value = GetTotalKreditMengendap(objData, DateAdd("m", 1, dTgl(0).Value), dTgl(1).Value)
  nBahas.Value = nTotalLabaRugi.Value * nPersenSHU1.Value / 100
  nBahas2.Value = nTotalLabaRugi.Value * nPersenSHU2.Value / 100
  nSimpananPokok.Value = GetTotalSimpananMengendap(objData, "T1")
  nSimpananWajib.Value = GetTotalSimpananMengendap(objData, "T2")
  
  'deposito
  nDeposito.Value = GetTotalDepositoMengendap(objData, dTgl(0).Value, dTgl(1).Value)
  
  Set db = objData.SQL(GetDSN, "select * from registernasabah")
  If Not db.eof Then
    Dim i  As Integer
    i = 1
    sGrid.Clear
    With vbalProgressBar1
      .Visible = True
      .Max = db.RecordCount
      .ShowText = True
    End With
    Do While Not db.eof
      vbalProgressBar1.Value = i
      vbalProgressBar1.Text = CLng(vbalProgressBar1.Percent) & "%"
      With sGrid
        .CellDetails i, 1, GetNull(db!Kode)
        .CellDetails i, 2, GetNull(db!nama)
        .CellDetails i, 3, GetSimpanan(.CellText(i, 1)), DT_RIGHT
        .CellDetails i, 4, GetTotalDepositoMengendapAnggota(objData, dTgl(0).Value, dTgl(1).Value, .CellText(i, 1)), DT_RIGHT
        .CellDetails i, 5, GetKreditMengendap(objData, GetNull(db!Kode), DateAdd("m", 1, dTgl(0).Value), dTgl(1).Value), DT_RIGHT
        .CellDetails i, 8, GetSimpananMengendapAnggota(objData, .CellText(i, 1), "T1", dTgl(0).Value, dTgl(1).Value), DT_RIGHT
        .CellDetails i, 9, GetSimpananMengendapAnggota(objData, .CellText(i, 1), "T2", dTgl(0).Value, dTgl(1).Value), DT_RIGHT
        db.MoveNext
      End With
      i = i + 1
    Loop
    nSimpanan.Value = GetTotalTabungan
    SHU
  End If
  vbalProgressBar1.Visible = False
  Me.MousePointer = vbDefault
End Sub

Private Sub SHU()
Dim lRow As Long

  With sGrid
    FrmPB.InitPB .Rows
    .Redraw = False
    For lRow = 1 To .Rows
       FrmPB.RunPB
      .CellDetails lRow, 6, Devide(.CellText(lRow, 3) + .CellText(lRow, 4) + .CellText(lRow, 5), nKredit.Value + nSimpanan.Value + nDeposito.Value) * 100, DT_RIGHT
      .CellDetails lRow, 7, .CellText(lRow, 6) * nBahas.Value / 100, DT_RIGHT
      
      .CellDetails lRow, 10, Devide(.CellText(lRow, 8) + .CellText(lRow, 9), nSimpananPokok.Value + nSimpananWajib.Value) * 100, DT_RIGHT
      .CellDetails lRow, 11, .CellText(lRow, 10) * nBahas2.Value / 100, DT_RIGHT
      .CellDetails lRow, 12, .CellText(lRow, 7) + .CellText(lRow, 11), DT_RIGHT
    Next lRow
    .Redraw = True
    FrmPB.EndPB
  End With
End Sub

Private Sub nKredit_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub sGrid_ColumnClick(ByVal lCol As Long)
Dim sTag As String
Dim iSortIndex As Long
      
   With sGrid.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lCol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lCol
      End If
   
      ' Determine which sort order to apply:
      sTag = sGrid.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      sGrid.ColumnTag(lCol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = sGrid.ColumnSortType(lCol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   sGrid.Sort
   Screen.MousePointer = vbDefault
End Sub

Private Function GetSimpanan(cRegister As String) As Double
Dim db As New ADODB.Recordset
Dim nYear As Integer
Dim nMonth As Integer
Dim nTabungan As Double
Dim n As Integer
  
  Set db = objData.Browse(GetDSN, "tabungan t", , "golongantabungan", sisAssign, "T3", " and kode = '" & cRegister & "'")
  If Not db.eof Then
    Do While Not db.eof
      n = DateDiff("m", dTgl(0).Value, dTgl(1).Value)
      nYear = Year(dTgl(0).Value)
      nMonth = Month(dTgl(0).Value)
      nTabungan = 0
      For n = 1 To n
        nTabungan = nTabungan + GetSaldo(GetNull(db!Rekening), nYear, nMonth)
        nYear = Year(DateAdd("m", n, dTgl(0).Value))
        nMonth = Month(DateAdd("m", n, dTgl(0).Value))
      Next n
      db.MoveNext
      GetSimpanan = GetSimpanan + nTabungan
    Loop
  End If
End Function

Private Function GetSaldo(ByVal cRekening As String, dTahun As Integer, dBulan As Integer) As Double
Dim dbSaldo As New ADODB.Recordset
Dim dTgl As Date
Dim dAkhirBulan As Date
Dim nSaldoAwal As Double
Dim cWhere As String
Dim vaSaldo As New XArrayDB
Dim n As Integer
Dim nTemp As Double
Dim cField As String
Dim cTgl As String

  GetSaldo = 0
  dTgl = DateSerial(dTahun, dBulan, 1)
  dAkhirBulan = EOM(DateAdd("m", -1, dTgl))
  cTgl = "Tgl <='" & Format(dAkhirBulan, "yyyy-mm-dd") & "'"
  cField = " Sum(If(DK='D' and " & cTgl & " ,Jumlah,0)) as Debet, "
  cField = cField & " Sum(If(DK='K' and " & cTgl & " ,Jumlah,0)) as Kredit"
  Set dbSaldo = objData.Browse(GetDSN, "MutasiTabungan", cField, "Rekening", sisAssign, cRekening, , "Rekening,Tgl")
  If Not dbSaldo.eof Then
    nSaldoAwal = GetNull(dbSaldo!Kredit) - GetNull(dbSaldo!Debet)
  End If

  vaSaldo.ReDim 0, 0, 0, 2
  vaSaldo(0, 0) = 0
  vaSaldo(0, 1) = 0
  vaSaldo(0, 2) = nSaldoAwal

  dAkhirBulan = EOM(dTgl)
  cWhere = "And Tgl >='" & Format(dTgl, "yyyy-mm-dd") & "'"
  cWhere = cWhere & "And Tgl <='" & Format(dAkhirBulan, "yyyy-mm-dd") & "'"
  Set dbSaldo = objData.Browse(GetDSN, "MutasiTabungan", "Jumlah,DK", "Rekening", sisAssign, cRekening, cWhere, "Tgl")
  If Not dbSaldo.eof Then
    dbSaldo.MoveFirst
    Do While Not dbSaldo.eof
      vaSaldo.InsertRows vaSaldo.UpperBound(1) + 1
      n = vaSaldo.UpperBound(1)

      vaSaldo(n, 0) = IIf(GetNull(dbSaldo!DK) = "D", GetNull(dbSaldo!Jumlah), 0)
      vaSaldo(n, 1) = IIf(GetNull(dbSaldo!DK) = "K", GetNull(dbSaldo!Jumlah), 0)
      vaSaldo(n, 2) = vaSaldo(n - 1, 2) + vaSaldo(n, 1) - vaSaldo(n, 0)
      dbSaldo.MoveNext
    Loop
  End If

  nTemp = vaSaldo(0, 2)
  For n = 1 To vaSaldo.UpperBound(1)
    If vaSaldo(n, 2) < nTemp Then
      nTemp = vaSaldo(n, 2)
    Else
      nTemp = nTemp
    End If
  Next
  GetSaldo = nTemp
End Function

Private Function GetTotalTabungan() As Double
Dim lRow As Long

  GetTotalTabungan = 0
  With sGrid
    FrmPB.InitPB .Rows
    For lRow = 1 To .Rows
      FrmPB.RunPB
      GetTotalTabungan = GetTotalTabungan + .CellText(lRow, 3)
    Next lRow
    FrmPB.EndPB
  End With
End Function

