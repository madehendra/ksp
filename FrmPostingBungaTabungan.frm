VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form FrmPostingBungaTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POSTING BUNGA SIMPANAN"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11625
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4980
      Left            =   0
      Top             =   615
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   8784
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
         Height          =   4635
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   8176
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "FormatText Event"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No Rekening"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Tgl Buka"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nama Nasabah"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Saldo Min"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Bunga"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Pajak"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Total Bunga"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###,###,##0.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1799"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1720"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=5292"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=5212"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2461"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2381"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1852"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1773"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=1879"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1799"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2593"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2514"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
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
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H0&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
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
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
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
      Begin MSComctlLib.ProgressBar PB 
         Height          =   270
         Left            =   60
         TabIndex        =   9
         Top             =   4695
         Visible         =   0   'False
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   630
      Left            =   0
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Left            =   8205
         TabIndex        =   8
         Top             =   135
         Width           =   2970
         _ExtentX        =   5239
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
         Caption         =   "Tanggal Mutasi"
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Left            =   225
         TabIndex        =   7
         Top             =   135
         Width           =   2715
         _ExtentX        =   4789
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
         Caption         =   "Periode Bunga"
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
      Begin BiSANumberBoxProject.BiSANumberBox nBulan 
         Height          =   330
         Left            =   2985
         TabIndex        =   4
         Top             =   135
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         DecimalPoint    =   ""
         Separator       =   ""
         MaxValue        =   12
         MinValue        =   1
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
         Caption         =   "Bulan/Tahun"
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
      Begin BiSANumberBoxProject.BiSANumberBox nTahun 
         Height          =   330
         Left            =   4770
         TabIndex        =   5
         Top             =   135
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         DecimalPoint    =   ""
         Separator       =   ""
         MaxValue        =   9999
         MinValue        =   1
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   5595
      Width           =   11625
      _ExtentX        =   20505
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
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10350
         TabIndex        =   1
         Top             =   120
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "FrmPostingBungaTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   8100
         TabIndex        =   2
         Top             =   120
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "    &Preview"
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
         Picture         =   "FrmPostingBungaTabungan.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdRefresh 
         Height          =   435
         Left            =   6945
         TabIndex        =   3
         Top             =   120
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Refresh"
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
         Picture         =   "FrmPostingBungaTabungan.frx":032C
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9270
         TabIndex        =   6
         Top             =   120
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
         Picture         =   "FrmPostingBungaTabungan.frx":04D6
      End
   End
End
Attribute VB_Name = "FrmPostingBungaTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cKode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "PeriodeBungaTabungan", "Kode,Bulan,Tahun", "Kode", sisContent, cKode.Text, "And Status <> '1'", "Kode")
  cKode.Text = cKode.Browse(dbData)
  If Not dbData.eof Then
    nBulan.Value = GetNull(dbData!Bulan)
    nTahun.Value = GetNull(dbData!Tahun)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  If vaArray.UpperBound(1) >= 0 Then
    GetPreview
  End If
End Sub

Private Sub cmdRefresh_Click()
  GetData
End Sub

'Private Sub cmdSimpan_Click()
'Dim cKodeBunga As String
'Dim n As Integer
'Dim cFaktur As String
'Dim vaField, vaValue
'Dim cAwal As String
'
'  If vaArray.UpperBound(1) < 0 Then
'    MsgBox "Tidak ada data yang diposting....", vbOKOnly + vbInformation, "POSTING BUNGA TABUNGAN"
'    nBulan.SetFocus
'    Exit Sub
'  End If
'
'  cAwal = "BT-" & Format(dDate.Value, "yymmdd") & "-"
'  objData.Delete GetDSN, "PostingBungaTabungan", "Kode", sisAssign, cKode.Text
'  cKodeBunga = aCfg(msKodeBagiHasil)
'  If MsgBox("Data akan disimpan ?", vbYesNo + vbInformation, "POSTING BUNGA TABUNGAN") = vbYes Then
'    vaField = Array("Faktur", "Kode", "Tanggal", "Rekening", "SaldoMinimal", "Bunga", "Pajak", "TotalBunga")
'    InitPB vaArray.UpperBound(1) + 1
'    TDBGrid1.MoveFirst
'    For n = 0 To vaArray.UpperBound(1)
'      RunPB
'      cFaktur = cAwal & Padl(n + 1, 5, "0")
'      UpdMutasiTabungan objData, cKodeBunga, cFaktur, dDate.Value, vaArray(n, 1), vaArray(n, 7), True, "Bunga Tabungan a.n " & vaArray(n, 3), False, "k"
'      UpdKodeTr objData, msTabungan, aCfg(msKodeCabang), cFaktur, dDate.Value, vaArray(n, 8), "Biaya Bunga tabungan a.n " & vaArray(n, 3), vaArray(n, 7), 0, "K", Now
'        UpdKodeTr objData, msTabungan, aCfg(msKodeCabang), cFaktur, dDate.Value, vaArray(n, 9), "Biaya Bunga tabungan a.n " & vaArray(n, 3), 0, vaArray(n, 7), "K", Now
'
'      vaValue = Array(cFaktur, cKode.Text, dDate.Value, vaArray(n, 1), vaArray(n, 4), vaArray(n, 5), vaArray(n, 6), vaArray(n, 7))
'      objData.Add GetDSN, "PostingBungaTabungan", vaField, vaValue
'    Next
'    EndPB
'
'    objData.Edit GetDSN, "PeriodeBungaTabungan", "Kode='" & cKode.Text & "'", Array("Status"), Array("1")
'
'    MsgBox "Proses selesai...", vbOKOnly + vbInformation, "POSTING BUNGA TABUNGAN"
'    vaArray.ReDim 0, -1, 0, 9
'    Set TDBGrid1.Array = vaArray
'    TDBGrid1.ReBind
'    cKode.Default
'    cKode.SetFocus
'    Exit Sub
'  End If
'End Sub
Private Sub cmdSimpan_Click()
Dim cKodeBunga As String
Dim n As Integer
Dim cFaktur As String
Dim vaField, vaValue
Dim cAwal As String
Dim objMutasi As New CodeSuiteLibrary.data

  
  If vaArray.UpperBound(1) < 0 Then
    MsgBox "Tidak ada data yang diposting....", vbOKOnly + vbInformation, "POSTING BUNGA TABUNGAN"
    nBulan.SetFocus
    Exit Sub
  End If
  
  cAwal = "BT-" & Format(dDate.Value, "yymmdd") & "-"
  objMutasi.Delete GetDSN, "PostingBungaTabungan", "Kode", sisAssign, cKode.Text
  cKodeBunga = aCfg(msKodeTransaksiPB)
  If MsgBox("Data akan disimpan ?", vbYesNo + vbInformation, "POSTING BUNGA TABUNGAN") = vbYes Then
    vaField = Array("Faktur", "Kode", "Tanggal", "Rekening", "SaldoMinimal", "Bunga", "Pajak", "TotalBunga")
    InitPB vaArray.UpperBound(1) + 1
    TDBGrid1.MoveFirst
    For n = 0 To vaArray.UpperBound(1)
      RunPB
      cFaktur = cAwal & Padl(n + 1, 5, "0")
      UpdMutasiTabungan objMutasi, cKodeBunga, cFaktur, dDate.Value, vaArray(n, 1), vaArray(n, 7), True, "Bunga Tabungan a.n " & vaArray(n, 3), False, "k"
      
      UpdKodeTr objMutasi, msTabungan, aCfg(msKodeCabang), cFaktur, dDate.Value, vaArray(n, 8), "Biaya Bunga tabungan a.n " & vaArray(n, 3), vaArray(n, 7), 0, "K", SNow
        UpdKodeTr objMutasi, msTabungan, aCfg(msKodeCabang), cFaktur, dDate.Value, vaArray(n, 9), "Biaya Bunga tabungan a.n " & vaArray(n, 3), 0, vaArray(n, 7), "K", SNow
      
      vaValue = Array(cFaktur, cKode.Text, dDate.Value, vaArray(n, 1), vaArray(n, 4), vaArray(n, 5), vaArray(n, 6), vaArray(n, 7))
      objMutasi.Add GetDSN, "PostingBungaTabungan", vaField, vaValue
    Next
    EndPB
    
    objMutasi.Edit GetDSN, "PeriodeBungaTabungan", "Kode='" & cKode.Text & "'", Array("Status"), Array("1")
    
    MsgBox "Proses selesai...", vbOKOnly + vbInformation, "POSTING BUNGA TABUNGAN"
    vaArray.ReDim 0, -1, 0, 9
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    cKode.Default
    cKode.SetFocus
    Exit Sub
  End If
  
End Sub
Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me, True
  nBulan.Value = 0
  nTahun.Value = 0
  cKode.Default
  vaArray.ReDim 0, -1, 0, 9
  
  TabIndex cKode, n
  TabIndex cmdRefresh, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim n As Integer
Dim vaJoin As String
Dim cWhere As String
Dim d2BlnLalu As Date
Dim dAkhirBulanIni As Date
Dim cField As String
  
  vaArray.ReDim 0, -1, 0, 9
  dAkhirBulanIni = EOM(DateSerial(nTahun.Value, nBulan.Value, 1))
  d2BlnLalu = EOM(DateAdd("m", -1, dAkhirBulanIni))
  
  
  cWhere = "And t.Tgl <= '" & Format(d2BlnLalu, "yyyy-mm-dd") & "'"
  cField = "t.Rekening,t.Tgl,r.Nama,g.Bunga,g.SaldoMinimumDapatBunga,g.SaldoMinimumKenaPajak,g.PajakBunga,g.RekeningBunga,g.rekening as RekAkuntansi"
  Set dbData = objData.Browse(GetDSN, "Tabungan t", cField, "t.close", sisDifference, "1", cWhere, "t.Rekening", _
                              Array("Left Join registernasabah r on r.Kode = t.Kode", _
                                    "Left Join Golongantabungan g on g.Kode = t.Golongantabungan"))
  If Not dbData.eof Then
    dbData.MoveFirst
    InitPB dbData.RecordCount
    Do While Not dbData.eof
      RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
        
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!Rekening, "")
      vaArray(n, 2) = GetNull(dbData!Tgl)
      vaArray(n, 3) = GetNull(dbData!nama, "")
      vaArray(n, 4) = GetSaldo(vaArray(n, 1))
      If vaArray(n, 4) >= GetNull(dbData!SaldoMinimumDapatBunga) Then
        vaArray(n, 5) = Round(vaArray(n, 4) * (GetNull(dbData!bunga) / 100 / 12))
        If vaArray(n, 4) >= GetNull(dbData!SaldoMinimumKenaPajak) Then
          vaArray(n, 6) = Round(vaArray(n, 5) * GetNull(dbData!pajakbunga) / 100)
        Else
          vaArray(n, 6) = 0
        End If
        vaArray(n, 7) = vaArray(n, 5) - vaArray(n, 6)
      Else
        vaArray(n, 5) = 0
        vaArray(n, 6) = 0
        vaArray(n, 7) = 0
      End If
      vaArray(n, 8) = GetNull(dbData!Rekeningbunga, "")
      vaArray(n, 9) = GetNull(dbData!RekAkuntansi, "")
      dbData.MoveNext
    Loop
    EndPB
  End If
  
  
  n = 0
  Do While n <= vaArray.UpperBound(1)
    If vaArray(n, 7) <= 0 Then
      vaArray.DeleteRows n
      n = n - 1
    End If
    n = n + 1
  Loop
  
  'urutkan nomor kembali
  Dim nBunga As Double
  Dim nSaldoMengendap As Double
  
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    vaArray(n, 0) = n + 1
    nBunga = nBunga + vaArray(n, 7)
    nSaldoMengendap = nSaldoMengendap + vaArray(n, 4)
  Next n
  TDBGrid1.Columns(4).FooterText = Format(nSaldoMengendap, "###,###,###,##0.00")
  TDBGrid1.Columns(7).FooterText = Format(nBunga, "###,###,###,##0.00")
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub
Private Sub GetPreview()
  With FrmRPT
    .AddPageHeader "HASIL PERHITUNGAN BUNGA SIMPANAN", tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader aCfg(msNama), tdbHalignCenter, , , True
    .AddPageHeader "Bulan/Tahun : " & nBulan.Value & "/" & nTahun.Value, tdbHalignCenter, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "NO", , , , 6
    .AddTableHeader "REKENING", , , , 13
    .AddTableHeader "TGL BUKA", , , , 9
    .AddTableHeader "NAMA NASABAH"
    .AddTableHeader "SALDO MIN", , , , 12
    .AddTableHeader "BUNGA", , , , 12
    .AddTableHeader "PAJAK", , , , 10
    .AddTableHeader "TOTAL BUNGA", , , , 12
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    
    .AddTableBody Sis_Rpt_Number, tdbHalignRight
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    
    .AddTableFooter "TOTAL", , tdbHalignRight, , , , , , , , , , , , 5
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&sum", Sis_Rpt_Number2
    .AddTableFooter "&sum", Sis_Rpt_Number2
    .AddTableFooter "&sum", Sis_Rpt_Number2
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    
    .Preview vaArray
  End With
End Sub

Private Sub InitPB(ByVal nMax)
  PB.Visible = True
  PB.Min = 0
  PB.Max = nMax + 1
  PB.Value = 0
End Sub

Private Sub RunPB()
  PB.Value = PB.Value + IIf(PB.Value < PB.Max, 1, 0)
End Sub

Private Sub EndPB()
  PB.Visible = False
End Sub

Private Function GetSaldo(ByVal cRekening As String) As Double
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
  dTgl = DateSerial(nTahun.Value, nBulan.Value, 1)
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
