VERSION 5.00
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form trCetakValidasiMutasiTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK VALIDASI TABUNGAN"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11595
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5640
      Left            =   -15
      Top             =   -15
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   9948
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4575
         Left            =   90
         TabIndex        =   0
         Top             =   990
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   8070
         _LayoutType     =   4
         _RowHeight      =   17
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No."
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Faktur"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tanggal"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Time"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "No. Rekening"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Nama Nasabah"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "User Name"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "FormatText Event"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Kode Trans."
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "D/K"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Nama Trans."
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Jumlah"
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   3
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   12107912
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=820"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=741"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3598"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3519"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=3387"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3307"
         Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=2434"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2355"
         Splits(0)._ColumnProps(30)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=4868"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=4789"
         Splits(0)._ColumnProps(35)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(37)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(42)=   "Column(8).Width=1005"
         Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=926"
         Splits(0)._ColumnProps(45)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(47)=   "Column(9).Width=767"
         Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=688"
         Splits(0)._ColumnProps(50)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(51)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(52)=   "Column(10).Width=5265"
         Splits(0)._ColumnProps(53)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(54)=   "Column(10)._WidthInPix=5186"
         Splits(0)._ColumnProps(55)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(56)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(57)=   "Column(11).Width=3678"
         Splits(0)._ColumnProps(58)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(11)._WidthInPix=3598"
         Splits(0)._ColumnProps(60)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(61)=   "Column(11).Order=12"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MarqueeUnique   =   0   'False
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483633
         RowDividerColor =   12107912
         RowSubDividerColor=   12107912
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.fgcolor=&H80000008&"
         _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HFFFF00&"
         _StyleDefs(10)  =   ":id=4,.fgcolor=&H800000&,.appearance=0,.ellipsis=0,.bold=-1,.fontsize=975"
         _StyleDefs(11)  =   ":id=4,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HFF8080&"
         _StyleDefs(14)  =   ":id=2,.fgcolor=&H80000014&,.transparentBmp=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(17)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.fgcolor=&H80000008&,.bold=0"
         _StyleDefs(18)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(19)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(20)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80000014&,.bold=-1"
         _StyleDefs(22)  =   ":id=6,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(23)  =   ":id=6,.fontname=Times New Roman"
         _StyleDefs(24)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(25)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(26)  =   ":id=8,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(27)  =   ":id=8,.fontname=MS Sans Serif"
         _StyleDefs(28)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H8000000F&"
         _StyleDefs(29)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(30)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(31)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(32)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(33)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(34)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(35)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(36)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(37)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.fgcolor=&H80000012&,.bold=-1"
         _StyleDefs(38)  =   ":id=18,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(39)  =   ":id=18,.fontname=Times New Roman"
         _StyleDefs(40)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(41)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80000018&"
         _StyleDefs(42)  =   ":id=19,.fgcolor=&H80000007&,.borderColor=&H80000007&"
         _StyleDefs(43)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(44)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(45)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(46)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(47)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(48)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(52)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(56)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(60)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(64)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(68)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(71)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(72)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(73)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(74)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(75)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(76)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(77)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(78)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(79)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(80)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(81)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(82)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(83)  =   "Splits(0).Columns(9).Style:id=82,.parent=13"
         _StyleDefs(84)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
         _StyleDefs(85)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
         _StyleDefs(86)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
         _StyleDefs(87)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
         _StyleDefs(88)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(89)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(90)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(91)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
         _StyleDefs(92)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(93)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(94)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(95)  =   "Named:id=33:Normal"
         _StyleDefs(96)  =   ":id=33,.parent=0"
         _StyleDefs(97)  =   "Named:id=34:Heading"
         _StyleDefs(98)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(99)  =   ":id=34,.wraptext=-1"
         _StyleDefs(100) =   "Named:id=35:Footing"
         _StyleDefs(101) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(102) =   "Named:id=36:Selected"
         _StyleDefs(103) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(104) =   "Named:id=37:Caption"
         _StyleDefs(105) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(106) =   "Named:id=38:HighlightRow"
         _StyleDefs(107) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(108) =   "Named:id=39:EvenRow"
         _StyleDefs(109) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(110) =   "Named:id=40:OddRow"
         _StyleDefs(111) =   ":id=40,.parent=33"
         _StyleDefs(112) =   "Named:id=41:RecordSelector"
         _StyleDefs(113) =   ":id=41,.parent=34"
         _StyleDefs(114) =   "Named:id=42:FilterBar"
         _StyleDefs(115) =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   5625
      Width           =   11565
      _ExtentX        =   20399
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
      BorderStyle     =   4
      BackColor       =   -2147483633
   End
End
Attribute VB_Name = "trCetakValidasiMutasiTAbungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
