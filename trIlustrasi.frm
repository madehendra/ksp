VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trIlustrasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ilustrasi Pinjaman"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11430
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   4230
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2460
      Width           =   11430
      _cx             =   20161
      _cy             =   7461
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4140
         Left            =   0
         TabIndex        =   7
         Top             =   15
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   7303
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KE"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "JATUH TEMPO"
         Columns(1).FooterText=   "Jumlah"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "dd-MM-yyyy"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "BUNGA"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "###,###,##0.00"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "POKOK"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "###,###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "ANGSURAN"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "SISA BUNGA"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "BAKI DEBET"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1376"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1296"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3466"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3387"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3704"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3625"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3916"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3836"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3731"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3651"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2963"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2884"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=3731"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=3651"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
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
         HeadLines       =   1.5
         FootLines       =   1.5
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=104,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFCFCED&,.fgcolor=&H80000008&"
         _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HEBDACB&"
         _StyleDefs(15)  =   ":id=3,.fgcolor=&H80000008&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(16)  =   ":id=3,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(29)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(30)  =   ":id=14,.fontname=Tahoma"
         _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(32)  =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(33)  =   ":id=15,.fontname=Tahoma"
         _StyleDefs(34)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(35)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(40)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(41)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(67)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(70)  =   "Named:id=33:Normal"
         _StyleDefs(71)  =   ":id=33,.parent=0"
         _StyleDefs(72)  =   "Named:id=34:Heading"
         _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(74)  =   ":id=34,.wraptext=-1"
         _StyleDefs(75)  =   "Named:id=35:Footing"
         _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   "Named:id=36:Selected"
         _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(79)  =   "Named:id=37:Caption"
         _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(81)  =   "Named:id=38:HighlightRow"
         _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(83)  =   "Named:id=39:EvenRow"
         _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(85)  =   "Named:id=40:OddRow"
         _StyleDefs(86)  =   ":id=40,.parent=33"
         _StyleDefs(87)  =   "Named:id=41:RecordSelector"
         _StyleDefs(88)  =   ":id=41,.parent=34"
         _StyleDefs(89)  =   "Named:id=42:FilterBar"
         _StyleDefs(90)  =   ":id=42,.parent=33"
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   675
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6690
      Width           =   11430
      _cx             =   20161
      _cy             =   1191
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
      Begin BiSAButtonProject.BiSAButton cmdPrint 
         Height          =   435
         Left            =   10140
         TabIndex        =   13
         Top             =   105
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   767
         Caption         =   "    &Cetak"
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
         Picture         =   "trIlustrasi.frx":0000
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   2460
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11430
      _cx             =   20161
      _cy             =   4339
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
      Align           =   1
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   9720
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   90
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         BorderStyle     =   0
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
      Begin VB.OptionButton optPerhitungan 
         Caption         =   "Flat"
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
         Index           =   1
         Left            =   2730
         TabIndex        =   6
         Top             =   2010
         Width           =   690
      End
      Begin VB.OptionButton optPerhitungan 
         Caption         =   "Menurun"
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
         Index           =   0
         Left            =   1725
         TabIndex        =   5
         Top             =   2010
         Width           =   1140
      End
      Begin BiSANumberBoxProject.BiSANumberBox nSukuBunga 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
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
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   165
         Width           =   5070
         _ExtentX        =   8943
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
      Begin BiSANumberBoxProject.BiSANumberBox nPlafond 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   795
         Width           =   3855
         _ExtentX        =   6800
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
      Begin BiSANumberBoxProject.BiSANumberBox nLama 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   1110
         Width           =   2415
         _ExtentX        =   4260
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
      Begin BiSANumberBoxProject.BiSANumberBox nBunga 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1425
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
      Begin VB.Label Label1 
         Caption         =   "/bulan"
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
         Left            =   2610
         TabIndex        =   12
         Top             =   510
         Width           =   555
      End
   End
End
Attribute VB_Name = "trIlustrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim dbData1 As New ADODB.Recordset
Dim xArray As New XArrayDB
Dim vaArray As New XArrayDB
Dim vaRPT As New XArrayDB
Dim nPos  As SisPos
Dim lEdit As Boolean
Dim cSQL As String
Dim NoRek As String
Dim dJatuhTempo As Date
Dim nTotalPokok As Double, nTotalBunga As Double, nTotalAngsuran As Double, nTotalTabungan As Double


Private Sub GetRpt()
    With FrmRPT
    .AddPageHeader "Ilustrasi Pinjaman", tdbHalignCenter, , , , , 10, True
    .AddPageHeader aCfg(msNama), tdbHalignCenter, , , , , 14, True, , True
    
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader "Bapak/Ibu " & vbTab & cNama.Text, tdbHalignLeft, , , True
    .AddPageHeader "Plafond " & vbTab & vbTab & Format(nPlafond.Value, "###,###,###,##0.00"), tdbHalignLeft, , , True
    .AddPageHeader "Bunga(pa) " & vbTab & Format(nSukuBunga.Value, "###,###,###,##0.00") & IIf(optPerhitungan(0).Value = True, "% Menurun", "% Menetap"), tdbHalignLeft, , , True
    .AddPageHeader "Lama " & vbTab & vbTab & Format(nLama.Value, "###,###,###,##0.00") & " Kali", tdbHalignLeft, , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Jatuh Tempo", , , , 13
    .AddTableHeader "Ag. Bunga", , , , 13
    .AddTableHeader "Ag. Pokok", , , , 13
    .AddTableHeader "Tot. Angsuran", , , , 13
    .AddTableHeader "Sisa Bunga", , , , 13
    .AddTableHeader "Baki Debet", , , , 13
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    
    .Preview xArray, True
  End With
End Sub


Private Sub GetJadwalMenurunNonPeriodik()
Dim n As Single
Dim dTanggal As Date
Dim nSukuBungaPerBulan As Double
Dim nKe As Integer

  nTotalPokok = 0
  nTotalBunga = 0
  xArray.ReDim 0, nLama.Value, 0, 6
  dTanggal = (DateAdd("m", 1, dTgl.Value))
  nSukuBungaPerBulan = nFnSukuBunga(False, nSukuBunga.Value)
  xArray(0, 5) = nBunga.Value
  xArray(0, 6) = nPlafond.Value
  nKe = 1
  For n = 1 To nLama.Value
    xArray(n, 0) = n
    xArray(n, 1) = Format(dTanggal, "dd/MM/yyyy")
    xArray(n, 2) = GetBungaReguler(xArray(n - 1, 6), nSukuBungaPerBulan)
    xArray(n, 3) = nPlafond.Value / (nLama.Value)
    xArray(n, 4) = xArray(n, 2) + xArray(n, 3)
    xArray(n, 5) = xArray(n - 1, 5) - xArray(n, 2)
    xArray(n, 6) = xArray(n - 1, 6) - xArray(n, 3)
    dTanggal = (DateAdd("m", 1, xArray(n, 1)))
  Next

  For n = 1 To xArray.UpperBound(1)
    nTotalBunga = nTotalBunga + xArray(n, 2)
    nTotalPokok = nTotalPokok + xArray(n, 3)
  Next

  TDBGrid1.Columns(2).FooterText = Format(nTotalBunga, "##,###,###,##0")
  TDBGrid1.Columns(3).FooterText = Format(nTotalPokok, "##,###,###,##0")

  TDBGrid1.Array = xArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Function nFnSukuBunga(ByVal lTahun As Boolean, nValue As Double) As Double
  nFnSukuBunga = nValue
  If lTahun Then
    nFnSukuBunga = Devide(nValue, 12)
  End If
End Function

Private Sub GetJadwalFlat()
Dim n As Single
Dim dTanggal As Date
Dim nSukuBungaPerBulan As Double
Dim nKe As Integer

  nTotalPokok = 0
  nTotalBunga = 0
  xArray.ReDim 0, nLama.Value, 0, 6
  dTanggal = (DateAdd("m", 1, dTgl.Value))
  nSukuBungaPerBulan = nFnSukuBunga(False, nSukuBunga.Value)
  xArray(0, 5) = nBunga.Value
  xArray(0, 6) = nPlafond.Value
  nKe = 1
  For n = 1 To nLama.Value
    xArray(n, 0) = n
    xArray(n, 1) = Format(dTanggal, "dd/MM/yyyy")
    xArray(n, 2) = nPlafond.Value * nFnSukuBunga(False, nSukuBunga.Value) / 100 'Devide(nBunga.Value, nLama.Value) 'GetBungaReguler(xArray(n - 1, 6), nSukuBungaPerBulan)
    xArray(n, 3) = nPlafond.Value / (nLama.Value)
    xArray(n, 4) = xArray(n, 2) + xArray(n, 3)
    xArray(n, 5) = xArray(n - 1, 5) - xArray(n, 2)
    xArray(n, 6) = xArray(n - 1, 6) - xArray(n, 3)
    dTanggal = (DateAdd("m", 1, xArray(n, 1)))
  Next

  For n = 1 To xArray.UpperBound(1)
    nTotalBunga = nTotalBunga + xArray(n, 2)
    nTotalPokok = nTotalPokok + xArray(n, 3)
  Next

  TDBGrid1.Columns(2).FooterText = Format(nTotalBunga, "##,###,###,##0")
  TDBGrid1.Columns(3).FooterText = Format(nTotalPokok, "##,###,###,##0")

  TDBGrid1.Array = xArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Function GetBungaReguler(ByVal nSisaPokok As Double, ByVal nBunga As Double) As Double
  GetBungaReguler = nSisaPokok * (nBunga / 100)
  GetBungaReguler = Mod50(GetBungaReguler)
End Function

Private Sub cmdPrint_Click()
  GetRpt
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  Total
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  dTgl.Value = Date
  TabIndex cNama, n
  TabIndex nSukuBunga, n
  TabIndex nPlafond, n
  TabIndex nLama, n
  TabIndex optPerhitungan(0), n
  TabIndex optPerhitungan(1), n
End Sub

Private Sub nLama_Validate(Cancel As Boolean)
  Total
End Sub

Private Sub nPlafond_Validate(Cancel As Boolean)
  Total
End Sub

Private Sub Total()
  nBunga.Value = Round(nPlafond.Value * nSukuBunga.Value / 100 / 12 * nLama.Value, 0)
  If optPerhitungan(0).Value = True Then
    GetJadwalMenurunNonPeriodik
  ElseIf optPerhitungan(1).Value = True Then
    GetJadwalFlat
  End If
End Sub

Private Sub nSukuBunga_Validate(Cancel As Boolean)
  Total
End Sub

Private Sub optPerhitungan_Click(Index As Integer)
  Total
End Sub
