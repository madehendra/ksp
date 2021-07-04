VERSION 5.00
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form rptLaporanBungaDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN TURUN BUNGA DEPOSITO"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11805
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   855
      Left            =   30
      Top             =   0
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   1508
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
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   345
         Left            =   3825
         TabIndex        =   2
         Top             =   330
         Width           =   585
      End
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   3600
         _ExtentX        =   6350
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
         Caption         =   "Nama"
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4725
      Left            =   60
      TabIndex        =   0
      Top             =   870
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   8334
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO"
      Columns(0).DataField=   ""
      Columns(0).NumberFormat=   "FormatText Event"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "REK"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NAMA"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "NOMINAL"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TGL VALUTA"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "JTH TEMPO"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "ARO"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "LAST PERPANJANGAN"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "LAMA"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "###,###,###,###,##0.00"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "SUKU BUNGA (pa)"
      Columns(9).DataField=   ""
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "BUNGA/BLN"
      Columns(10).DataField=   ""
      Columns(10).NumberFormat=   "###,###,###,###,##0.00"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "TOT BUNGA"
      Columns(11).DataField=   ""
      Columns(11).NumberFormat=   "###,###,###,###,##0.00"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "CAIR"
      Columns(12).DataField=   ""
      Columns(12).NumberFormat=   "###,###,###,###,##0.00"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "TITIPAN"
      Columns(13).DataField=   ""
      Columns(13).NumberFormat=   "Standard"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "SISA"
      Columns(14).DataField=   ""
      Columns(14).NumberFormat=   "Standard"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2461"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2381"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=5794"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5715"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2910"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2831"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2143"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2064"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2302"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2223"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1217"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1138"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=3281"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3201"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=1349"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1270"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2699"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2619"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=2408"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2328"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=3466"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=3387"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=3598"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=3519"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=514"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=514"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=514"
      Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HEBDACB&"
      _StyleDefs(15)  =   ":id=3,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
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
      _StyleDefs(26)  =   "Splits(0).Style:id=95,.parent=1"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=104,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=96,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=97,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=98,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=100,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=99,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=101,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=102,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=103,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=105,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=106,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=110,.parent=95,.alignment=1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=107,.parent=96"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=108,.parent=97"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=109,.parent=99"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=114,.parent=95"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=111,.parent=96"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=112,.parent=97"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=113,.parent=99"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=95"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=96"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=97"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=99"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=122,.parent=95,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=119,.parent=96"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=120,.parent=97"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=121,.parent=99"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=126,.parent=95,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=123,.parent=96"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=124,.parent=97"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=125,.parent=99"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=130,.parent=95,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=127,.parent=96"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=128,.parent=97"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=129,.parent=99"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=134,.parent=95,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=131,.parent=96"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=132,.parent=97"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=133,.parent=99"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=138,.parent=95,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=135,.parent=96"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=136,.parent=97"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=137,.parent=99"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=142,.parent=95,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=139,.parent=96"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=140,.parent=97"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=141,.parent=99"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=146,.parent=95,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=143,.parent=96"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=144,.parent=97"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=145,.parent=99"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=150,.parent=95,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=147,.parent=96"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=148,.parent=97"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=149,.parent=99"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=154,.parent=95,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=151,.parent=96"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=152,.parent=97"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=153,.parent=99"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=158,.parent=95,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=155,.parent=96"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=156,.parent=97"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=157,.parent=99"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=162,.parent=95,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=159,.parent=96"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=160,.parent=97"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=161,.parent=99"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=74,.parent=95,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=71,.parent=96"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=72,.parent=97"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=73,.parent=99"
      _StyleDefs(98)  =   "Named:id=33:Normal"
      _StyleDefs(99)  =   ":id=33,.parent=0"
      _StyleDefs(100) =   "Named:id=34:Heading"
      _StyleDefs(101) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(102) =   ":id=34,.wraptext=-1"
      _StyleDefs(103) =   "Named:id=35:Footing"
      _StyleDefs(104) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(105) =   "Named:id=36:Selected"
      _StyleDefs(106) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(107) =   "Named:id=37:Caption"
      _StyleDefs(108) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(109) =   "Named:id=38:HighlightRow"
      _StyleDefs(110) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(111) =   "Named:id=39:EvenRow"
      _StyleDefs(112) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(113) =   "Named:id=40:OddRow"
      _StyleDefs(114) =   ":id=40,.parent=33"
      _StyleDefs(115) =   "Named:id=41:RecordSelector"
      _StyleDefs(116) =   ":id=41,.parent=34"
      _StyleDefs(117) =   "Named:id=42:FilterBar"
      _StyleDefs(118) =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "rptLaporanBungaDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim lClick As Boolean

Private Sub cmdOK_Click()
Dim n As Integer
Dim nTmpCair As Double
Dim nTmpTitipan As Double


  Set dbData = objData.Browse(GetDSN, "Deposito d", "d.Rekening,d.Tgl,d.JthTmp,d.LastPerpanjangan,d.SistemARO,r.Nama,d.NominalDeposito,g.Lama,d.SukuBunga", "r.Nama", sisContent, cNama.Text, , , Array("Left Join RegisterNasabah r on r.Kode = d.Kode", "Left Join GolonganDeposito g on g.Kode = d.GolonganDeposito"))
  If Not dbData.eof Then
    vaArray.ReDim 0, -1, 0, 14
    Set TDBGrid1.Array = vaArray
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!Rekening, "")
      vaArray(n, 2) = GetNull(dbData!nama, "")
      vaArray(n, 3) = GetNull(dbData!nominaldeposito, "")
      vaArray(n, 4) = GetNull(dbData!Tgl)
      vaArray(n, 5) = GetNull(dbData!jthtmp)
      vaArray(n, 6) = GetNull(dbData!SistemARO, "")
      vaArray(n, 7) = GetNull(dbData!LastPerpanjangan)
      vaArray(n, 8) = GetNull(dbData!Lama)
      vaArray(n, 9) = GetNull(dbData!SukuBunga)
      vaArray(n, 10) = vaArray(n, 9) / 14 * vaArray(n, 3) / 140
      vaArray(n, 11) = vaArray(n, 8) * vaArray(n, 10)
      vaArray(n, 12) = GetCair(objData, vaArray(n, 1))
      vaArray(n, 13) = GetTitipan(objData, vaArray(n, 1))
      vaArray(n, 14) = vaArray(n, 11) - vaArray(n, 12) - vaArray(n, 13)
      nTmpCair = nTmpCair + vaArray(n, 12)
      nTmpTitipan = nTmpTitipan + vaArray(n, 13)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    Set TDBGrid1.Array = vaArray
    TDBGrid1.Columns(12).FooterText = Format(nTmpCair, "###,###,###,###,##0.00")
    TDBGrid1.Columns(13).FooterText = Format(nTmpTitipan, "###,###,###,###,##0.00")
    TDBGrid1.ReBind
    TDBGrid1.Refresh
  End If

End Sub

Private Sub Form_Load()
Dim n As Single

  lClick = True
  CenterForm Me, True
  cNama.Default
  vaArray.ReDim 0, -1, 0, 14
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  TabIndex cNama, n
  TabIndex cmdOK, n
End Sub

Private Function GetCair(ByVal obj As CodeSuiteLibrary.data, ByVal Rekening As String) As Double
Dim db As New ADODB.Recordset
  
  GetCair = 0
  Set db = obj.Browse(GetDSN, "MutasiDeposito", "sum(jumlah) as Jumlah", "Rekening", sisAssign, Rekening, " and kodemutasi = '3'")
  If Not db.eof Then
    GetCair = GetNull(db!Jumlah)
  End If
End Function

Private Function GetTitipan(ByVal obj As CodeSuiteLibrary.data, ByVal Rekening As String) As Double
Dim db As New ADODB.Recordset
  
  GetTitipan = 0
  Set db = obj.Browse(GetDSN, "BungaDeposito", "sum(Bunga) as Bunga", "Rekening", sisAssign, Rekening)
  If Not db.eof Then
    GetTitipan = GetNull(db!bunga)
  End If
End Function

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
If lClick Then
    Select Case ColIndex
      Case 1
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
        lClick = Not lClick
      Case 2
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
        lClick = Not lClick
    End Select
  Else
    Select Case ColIndex
      Case 1
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 1, XORDER_DESCEND, XTYPE_STRING
        lClick = Not lClick
      Case 2
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 2, XORDER_DESCEND, XTYPE_STRING
        lClick = Not lClick
    End Select
  End If
  TDBGrid1.ReBind
End Sub
