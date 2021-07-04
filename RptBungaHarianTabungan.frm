VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptBungaHarianTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN BUNGA HARIAN TABUNGAN"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11625
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4710
      Left            =   0
      Top             =   975
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   8308
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
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   8070
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TANGGAL"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "SALDO TABUNGAN"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "RATE"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "RUMUS"
         Columns(3).FooterText=   "                        JUMLAH"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "BUNGA"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "PAJAK"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "FormatText Event"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "BUNGA BERSIH"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3334"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3254"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=514"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1535"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1455"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=4339"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=4260"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2937"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2858"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2328"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2249"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=3149"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=3069"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(66)  =   "Named:id=33:Normal"
         _StyleDefs(67)  =   ":id=33,.parent=0"
         _StyleDefs(68)  =   "Named:id=34:Heading"
         _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(70)  =   ":id=34,.wraptext=-1"
         _StyleDefs(71)  =   "Named:id=35:Footing"
         _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   "Named:id=36:Selected"
         _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=37:Caption"
         _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(77)  =   "Named:id=38:HighlightRow"
         _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(79)  =   "Named:id=39:EvenRow"
         _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(81)  =   "Named:id=40:OddRow"
         _StyleDefs(82)  =   ":id=40,.parent=33"
         _StyleDefs(83)  =   "Named:id=41:RecordSelector"
         _StyleDefs(84)  =   ":id=41,.parent=34"
         _StyleDefs(85)  =   "Named:id=42:FilterBar"
         _StyleDefs(86)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1720
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   5520
         TabIndex        =   1
         Top             =   105
         Width           =   4575
         _ExtentX        =   8070
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
         Caption         =   "NAMA"
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
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   105
         Width           =   3165
         _ExtentX        =   5583
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
         Caption         =   "ANTARA TANGGAL"
         CaptionWidth    =   1700
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
         Left            =   3345
         TabIndex        =   3
         Top             =   105
         Width           =   2010
         _ExtentX        =   3545
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
         Caption         =   "S.D"
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
      Begin BiSATextBoxProject.BiSATextBox cCabang 
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   495
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   582
         Text            =   "1234"
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
         MaxLength       =   4
         Caption         =   "NO. REKENING"
         CaptionWidth    =   1700
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
         Left            =   2295
         TabIndex        =   5
         Top             =   495
         Width           =   780
         _ExtentX        =   1376
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
         Left            =   3120
         TabIndex        =   6
         Top             =   495
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   3915
         TabIndex        =   7
         Top             =   495
         Width           =   390
         _ExtentX        =   688
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
         Left            =   5520
         TabIndex        =   8
         Top             =   495
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
         Caption         =   "ALAMAT"
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
      Top             =   5685
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin MSComctlLib.ProgressBar PB 
         Height          =   465
         Left            =   60
         TabIndex        =   9
         Top             =   90
         Visible         =   0   'False
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   820
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10335
         TabIndex        =   10
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
         Picture         =   "RptBungaHarianTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   9165
         TabIndex        =   11
         Top             =   120
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
         Picture         =   "RptBungaHarianTabungan.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdRefresh 
         Height          =   435
         Left            =   7995
         TabIndex        =   12
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
         Picture         =   "RptBungaHarianTabungan.frx":032C
      End
   End
End
Attribute VB_Name = "RptBungaHarianTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vaarray As New XArrayDB
Dim dbData As New ADODB.Recordset
Dim objData As New BiSAMyDLL.data
Dim cRekening As String
Dim lPreview As Boolean
Dim nTotalBunga As Double
Dim nTotalPajak As Double
Dim nTotalNetto As Double
Dim nTotalSaldo As Double

Private Sub GetSaldoTabunganHarian()
Dim n As Integer
Dim nJumlahHari As Integer
Dim dTanggal As Date
    
    objData.OpenConnection GetDSN
    Set vaarray = GetBungaHarian(objData, cRekening, dTgl(0).Value, dTgl(1).Value, nTotalBunga, nTotalPajak, , nTotalSaldo, True)
    nTotalNetto = nTotalBunga - nTotalPajak
    n = 0
    Do While n <= vaarray.UpperBound(1)
      If vaarray(n, 1) = 0 Then
        vaarray.DeleteRows n
        n = n - 1
      End If
      n = n + 1
    Loop
    
    PB.Visible = False
    TDBGrid1.Columns(1).FooterText = Format(nTotalSaldo, "###,###,###,###,##0.00")
    TDBGrid1.Columns(4).FooterText = Format(nTotalBunga, "###,###,###,###,##0.00")
    TDBGrid1.Columns(5).FooterText = Format(nTotalPajak, "###,###,###,###,##0.00")
    TDBGrid1.Columns(6).FooterText = Format(nTotalNetto, "###,###,###,###,##0.00")
    
    Set TDBGrid1.Array = vaarray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
    objData.CloseConnection GetDSN
    If lPreview = True Then
      rpt
    End If
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
  If cFrekuensi.LastKey = 13 Or cFrekuensi.LastKey = 40 Then
    If cFrekuensi.Text <> "" Then
      cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
      Set dbData = objData.Browse(GetDSN, "Tabungan t", "t.Close,r.Nama,r.Alamat", "t.Rekening", sisAssign, cRekening, , , Array("Left Join RegisterNasabah r on r.Kode=t.Kode"))
      If Not dbData.eof Then
        If dbData!Close = "1" Then
          MsgBox "Rekening tersebut sudah tutup. Silahkan ulangi pengisian", vbInformation
          Cancel = True
          cFrekuensi.Default
          cFrekuensi.SetFocus
          Exit Sub
        End If
        cNama.Text = dbData!Nama
        cAlamat.Text = dbData!Alamat
      Else
        MsgBox "Data tidak ada. Silahkan Ulangi pengisian", vbInformation
        Cancel = True
        cFrekuensi.SetFocus
        Exit Sub
      End If
    Else
      MsgBox "Inputan tidak boleh kosong", vbInformation
      Cancel = True
      cFrekuensi.SetFocus
      Exit Sub
    End If
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganTabungan", "Kode", cGolongan, "Kode,Keterangan")
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  lPreview = True
  GetSaldoTabunganHarian
End Sub

Private Sub cmdRefresh_Click()
  lPreview = False
  GetSaldoTabunganHarian
End Sub

Private Sub Initvalue()
  cNama.Default
  cAlamat.Default
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = Date
  cCabang.Text = aCfg(msKodeCabang)
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Tabungan t", "r.Nama,r.alamat,t.Rekening,t.Awal,t.Akhir", "r.Nama", sisContent, cNama.Text, "And t.Close<>'1' And r.ALamat Like '" & cAlamat.Text & "%'", , Array("Left Join RegisterNasabah r on r.Kode=t.Kode"))
  cNama.Text = cNama.Browse(dbData, Array("Nama", "Alamat", "Rekening"))
  If Not dbData.eof Then
     GetRegister
  End If
End Sub

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Tabungan t", "r.alamat,r.Nama,t.Rekening,t.Awal,t.Akhir", "r.Alamat", sisContent, cAlamat.Text, "And t.Close<>'1' And r.Nama Like '" & cNama.Text & "%'", , Array("Left Join RegisterNasabah r on r.Kode=t.Kode"))
  cAlamat.Text = cAlamat.Browse(dbData, Array("Nama", "Alamat", "Rekening"))
  If Not dbData.eof Then
     GetRegister
  End If
End Sub

Private Sub GetRegister()
  cGolongan.Text = Mid(dbData!Rekening, 4, 2)
  cUrut.Text = Mid(dbData!Rekening, 7, 6)
  cFrekuensi.Text = Right(dbData!Rekening, 2)
  cNama.Text = dbData!Nama
  cAlamat.Text = dbData!Alamat
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me, True
  Initvalue
  
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cmdRefresh, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  Value = Format(Value, "###,###,###,###,##0.00")
End Sub

Private Sub rpt()
  'FrmPengesahan.GetPengesahaan Me.name
  
  With FrmRPT
    .AddPageHeader "Laporan Bunga Harian", tdbHalignCenter, , , , , 12, True
    .AddPageHeader "Antara Tanggal : " & Format(dTgl(0).Value, "dd-MM-yyyy") & " s.d " & Format(dTgl(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 12, True
    
    .AddPageHeader "No. Rekening ", , , 15, True, , , , , True, , tdbPageHeaderSect, , , , 5
    .AddPageHeader ": " & cRekening
    .AddPageHeader "Nama Nasabah", , , 15, True
    .AddPageHeader ": " & cNama.Text
    .AddPageHeader "Alamat Nasabah", , , 15, True
    .AddPageHeader ": " & cAlamat.Text
    
    .AddTableHeader "Tanggal", , tdbHalignCenter, , 10
    .AddTableHeader "Saldo Tab", , , , 15
    .AddTableHeader "RATE", , tdbHalignGeneral, , 5
    .AddTableHeader "RUMUS"
    .AddTableHeader "Bunga", , , , 13
    .AddTableHeader "Pajak Bunga", , , , 13
    .AddTableHeader "Bunga Bersih", , , , 15
    
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody , tdbHalignCenter
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody , , , , , , , , , , , , , False
    
    .AddTableFooter "Jumlah", , tdbHalignCenter, , , , , , , , , , , , 4
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter nTotalBunga, Sis_Rpt_Number2
    .AddTableFooter nTotalPajak, Sis_Rpt_Number2
    .AddTableFooter nTotalNetto, Sis_Rpt_Number2
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    
    GetRptFooter
    
    .Preview vaarray, True
  End With
End Sub



