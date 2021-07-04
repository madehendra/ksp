VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form trSaldoAwalKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo Awal Kredit"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11085
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6765
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11085
      _cx             =   19553
      _cy             =   11933
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
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "&1 Rekening Kredit"
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
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   6390
         Left            =   11730
         TabIndex        =   7
         Top             =   330
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   11271
         Picture         =   "trSaldoAwalKredit.frx":0000
         ForeColor       =   0
         BarPicture      =   "trSaldoAwalKredit.frx":001C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne3 
         Height          =   6390
         Left            =   12030
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   330
         Width           =   10995
         _cx             =   19394
         _cy             =   11271
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
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   6390
         Left            =   45
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   330
         Width           =   10995
         _cx             =   19394
         _cy             =   11271
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
         BorderWidth     =   1
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
         Begin TrueOleDBGrid70.TDBGrid DataGrid 
            Height          =   6330
            Left            =   30
            TabIndex        =   3
            Top             =   30
            Width           =   10950
            _ExtentX        =   19315
            _ExtentY        =   11165
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
            Columns(1).Caption=   "Rekening"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nama"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Plafond"
            Columns(3).DataField=   ""
            Columns(3).NumberFormat=   "###,###,###,##0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Pokok Yg Sdh Byr"
            Columns(4).DataField=   ""
            Columns(4).NumberFormat=   "###,###,###,##0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Bunga Yg Sdh Byr"
            Columns(5).DataField=   ""
            Columns(5).NumberFormat=   "###,###,###,##0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Denda Yg Sdh Byr"
            Columns(6).DataField=   ""
            Columns(6).NumberFormat=   "###,###,###,##0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Baki Debet"
            Columns(7).DataField=   ""
            Columns(7).NumberFormat=   "###,###,###,##0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1164"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3201"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3122"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=4895"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4815"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=3096"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3016"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=3043"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2963"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=3149"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3069"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2831"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2752"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(36)=   "Column(7).Width=3678"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3598"
            Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            BorderStyle     =   0
            ColumnFooters   =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14737632
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(19)  =   ":id=13,.strikethrough=0,.charset=0"
            _StyleDefs(20)  =   ":id=13,.fontname=Tahoma"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bold=0,.fontsize=825"
            _StyleDefs(23)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(24)  =   ":id=14,.fontname=Tahoma"
            _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(26)  =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(27)  =   ":id=15,.fontname=Tahoma"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   540
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6765
      Width           =   11085
      _cx             =   19553
      _cy             =   953
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
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   360
         Left            =   75
         TabIndex        =   8
         Top             =   90
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         Picture         =   "trSaldoAwalKredit.frx":0038
         ForeColor       =   0
         BarPicture      =   "trSaldoAwalKredit.frx":0054
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
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   7740
         TabIndex        =   5
         Top             =   45
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
         Picture         =   "trSaldoAwalKredit.frx":0070
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9975
         TabIndex        =   6
         Top             =   45
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
         Picture         =   "trSaldoAwalKredit.frx":060A
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   8820
         TabIndex        =   9
         Top             =   45
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Preview"
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
         Picture         =   "trSaldoAwalKredit.frx":06B0
      End
   End
End
Attribute VB_Name = "trSaldoAwalKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB

Private Function IsValidGolonganKredit(Rekening As String) As Boolean
Dim db As New ADODB.Recordset
Dim cGolongan As String
IsValidGolonganKredit = False

  cGolongan = Mid(Rekening, 1, 2)
  Set db = objData.Browse(GetDSN, "golongankredit", , "kode", sisAssign, cGolongan)
  If Not db.eof Then
    IsValidGolonganKredit = True
  End If
End Function

Private Sub PopulateKredit()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim nTotal As Double
Dim nTotalBakiDebet As Double

  ProgressBar.Visible = True
  nTotalBakiDebet = 0
  CleanGrid
  'Set db = objData.Browse(GetDSN, "debitur d", "d.rekening,d.plafond,r.nama,sa.pokok,sa.bunga,sa.denda", , , , "1=1 and d.tgl <= '2006-12-31'", "d.rekening", Array("left join registernasabah r on r.kode = d.kode", "left join saldoawalkredit sa on sa.rekening = d.rekening"))
  Set db = objData.Browse(GetDSN, "debitur d", "d.rekening,d.plafond,r.nama,sa.pokok,sa.bunga,sa.denda", , , , , "d.rekening", Array("left join registernasabah r on r.kode = d.kode", "left join saldoawalkredit sa on sa.rekening = d.rekening"))
  If Not db.eof Then
    ProgressBar.Max = GetNull(db.RecordCount)
    Do While Not db.eof
      ProgressBar.Value = ProgressBar.Value + 1
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(db!Rekening)
      vaArray(n, 2) = GetNull(db!nama)
      vaArray(n, 3) = GetNull(db!plafond)
      vaArray(n, 4) = GetNull(db!pokok)
      vaArray(n, 5) = GetNull(db!bunga)
      vaArray(n, 6) = GetNull(db!denda)
      vaArray(n, 7) = vaArray(n, 3) - vaArray(n, 4)
      nTotal = nTotal + vaArray(n, 4)
      nTotalBakiDebet = nTotalBakiDebet + vaArray(n, 7)
      db.MoveNext
    Loop
    
    DataGrid.Columns(4).FooterText = Format(nTotal, "###,###,##0.00")
    DataGrid.Columns(7).FooterText = Format(nTotalBakiDebet, "###,###,##0.00")
    Set DataGrid.Array = vaArray
    DataGrid.ReBind
    DataGrid.Refresh
  End If
  db.Close
  ProgressBar.Visible = False
End Sub

Private Sub CleanGrid()
  vaArray.ReDim 0, -1, 0, 7
  Set DataGrid.Array = vaArray
  DataGrid.ReBind
  DataGrid.Refresh
  DataGrid.Columns(4).FooterText = Format(0, "###,###,##0.00")
  DataGrid.Columns(7).FooterText = Format(0, "###,###,##0.00")

End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetRpt
End Sub

Private Sub GetRpt()
    
'    vaArray(n, 0) = n + 1
'    vaArray(n, 1) = GetNull(db!Rekening)
'    vaArray(n, 2) = GetNull(db!nama)
'    vaArray(n, 3) = GetNull(db!plafond)
'    vaArray(n, 4) = GetNull(db!pokok)
'    vaArray(n, 5) = GetNull(db!bunga)
'    vaArray(n, 6) = GetNull(db!denda)
'    vaArray(n, 7) = vaArray(n, 3) - vaArray(n, 4)

    With FrmRPT
    .AddPageHeader "LAPORAN BAKI DEBET SALDO AWAL KREDIT", tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Rekening", , , , 12
    .AddTableHeader "Nama"
    .AddTableHeader "Plafond", , , , 13
    .AddTableHeader "Pokok", , , , 13
    .AddTableHeader "Bunga", , , , 12
    .AddTableHeader "Denda", , , , 12
    .AddTableHeader "Baki Debet", , , , 13
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2

    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "GRAND TOTAL", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub


Private Sub cmdSimpan_Click()
Dim n As Integer
Dim vaField
Dim vaValue
Dim cGolonganKredit As String
Dim cFaktur As String
Dim dTgl As Date
Dim cRekeningPokok As String
Dim cRekeningBunga As String
Dim cRekeningDenda As String
Dim nSaldoAwalKredit As Double

  ProgressBar.Visible = True
  ProgressBar.Max = vaArray.UpperBound(1)
  ProgressBar.Value = 0
  cFaktur = "SAK"
  dTgl = "2006-12-31"
  objData.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msAngsuranKredit, "And Faktur='" & cFaktur & "'"
  objData.Delete GetDSN, "Angsuran", "Faktur", sisAssign, cFaktur
  objData.Delete GetDSN, "saldoawalkredit", "Faktur", sisAssign, cFaktur
  nSaldoAwalKredit = 0
  
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    ProgressBar.Value = ProgressBar.Value + 1
    
'    If vaArray(n, 4) = 0 And vaArray(n, 5) = 0 And vaArray(n, 6) = 0 Then
'      objData.Delete GetDSN, "saldoawalkredit", "rekening", sisAssign, vaArray(n, 1)
'    Else
'    End If
    
    vaField = Array("faktur", "rekening", "pokok", "bunga", "denda", "username", "datetime")
    vaValue = Array(cFaktur, vaArray(n, 1), vaArray(n, 4), vaArray(n, 5), vaArray(n, 6), GetRegistry(reg_UserName), SNow)
    objData.Update GetDSN, "saldoawalkredit", "rekening = '" & vaArray(n, 1) & "'", vaField, vaValue
    
    cGolonganKredit = Mid(vaArray(n, 1), 4, 2)
    Set dbData = objData.Browse(GetDSN, "GolonganKredit", , "Kode", sisAssign, cGolonganKredit)
    If Not dbData.eof Then
      cRekeningPokok = GetNull(dbData!RekeningAngsuranPokok, "")
      cRekeningBunga = GetNull(dbData!rekeningangsuranbunga, "")
      cRekeningDenda = GetNull(dbData!rekeningdenda, "")
    End If
    vaField = Array("Faktur", "Tgl", "Rekening", "Pokok", "Bunga", "Denda", _
                    "Total", "DateTime", "UserName")
    vaValue = Array(cFaktur, dTgl, vaArray(n, 1), vaArray(n, 4), vaArray(n, 5), vaArray(n, 6), _
                    CDbl(vaArray(n, 4)) + CDbl(vaArray(n, 5)) + CDbl(vaArray(n, 6)), SNow, GetRegistry(reg_UserName))
    
    objData.Add GetDSN, "Angsuran", vaField, vaValue

'      UpdKodeTr objData, msAngsuranKredit, aCfg(msKodeCabang), cFaktur, dTgl, cKasTeller, "Angsuran Kredit an. " & vaArray(n, 2), CDbl(vaArray(n, 4)) + CDbl(vaArray(n, 5)) + CDbl(vaArray(n, 6)), 0, "K"
'        UpdKodeTr objData, msAngsuranKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningPokok, "Angsuran Pokok Kredit an. " & vaArray(n, 2), 0, vaArray(n, 4), "K"
'        UpdKodeTr objData, msAngsuranKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningBunga, "Angsuran Bunga Kredit an. " & vaArray(n, 2), 0, vaArray(n, 5), "K"
'        UpdKodeTr objData, msAngsuranKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningDenda, "Denda Angsuran Kredit an. " & vaArray(n, 2), 0, vaArray(n, 6), "K"
    
    nSaldoAwalKredit = nSaldoAwalKredit + vaArray(n, 4)

  Next n
  
  cRekeningPokok = cRekeningPokok
'  UpdSaldoAwal objData, cRekeningPokok, nSaldoAwalKredit
  ProgressBar.Visible = False
  MsgBox "Data telah disimpan", , "Berhasil"
  PopulateKredit
End Sub

Private Sub UpdSaldoAwal(ByVal obj As CodeSuiteLibrary.data, ByVal Rek As String, nValue As Double)
Dim db As New ADODB.Recordset

  obj.Update GetDSN, "saldorekening", "rekening = '" & Rek & "'", Array("rekening", "awal"), Array(Rek, nValue)
End Sub

Private Sub DataGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If ColIndex < 3 Then
    Cancel = True
  Else
    If Not IsNumeric(DataGrid.Columns(3).Value) Or Not IsNumeric(DataGrid.Columns(4).Value) Or Not IsNumeric(DataGrid.Columns(5).Value) Then
      Cancel = 1
    End If
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  ProgressBar.Visible = False
  CenterForm Me
  CleanGrid
  PopulateKredit
  TabIndex TabOne1, n
  TabIndex DataGrid, n
  TabIndex cmdSimpan, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub
