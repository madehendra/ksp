VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form MstSARekening 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER SALDO AWAL REKENING"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9645
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4260
      Left            =   0
      Top             =   1335
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   7514
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
      Begin TrueOleDBGrid70.TDBGrid DataGrid1 
         Height          =   4125
         Left            =   75
         TabIndex        =   0
         Top             =   75
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7276
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "CABANG"
         Columns(0).DataField=   "Kode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "REKENING"
         Columns(1).DataField=   "Keterangan"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA REKENING"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "JENIS"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "SALDO AWAL"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2937"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2858"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=7038"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=6959"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1323"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1244"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3201"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3122"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
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
         HeadLines       =   1.5
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
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
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Named:id=33:Normal"
         _StyleDefs(58)  =   ":id=33,.parent=0"
         _StyleDefs(59)  =   "Named:id=34:Heading"
         _StyleDefs(60)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   ":id=34,.wraptext=-1"
         _StyleDefs(62)  =   "Named:id=35:Footing"
         _StyleDefs(63)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   "Named:id=36:Selected"
         _StyleDefs(65)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(66)  =   "Named:id=37:Caption"
         _StyleDefs(67)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(68)  =   "Named:id=38:HighlightRow"
         _StyleDefs(69)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(70)  =   "Named:id=39:EvenRow"
         _StyleDefs(71)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(72)  =   "Named:id=40:OddRow"
         _StyleDefs(73)  =   ":id=40,.parent=33"
         _StyleDefs(74)  =   "Named:id=41:RecordSelector"
         _StyleDefs(75)  =   ":id=41,.parent=34"
         _StyleDefs(76)  =   "Named:id=42:FilterBar"
         _StyleDefs(77)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   2355
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
      Begin BiSANumberBoxProject.BiSANumberBox nAwal 
         Height          =   330
         Left            =   375
         TabIndex        =   1
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "S. AWAL"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGudang 
         Height          =   330
         Left            =   2805
         TabIndex        =   2
         Top             =   105
         Width           =   4290
         _ExtentX        =   7567
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
      Begin BiSATextBoxProject.BiSABrowse cGudang 
         Height          =   330
         Left            =   375
         TabIndex        =   3
         Top             =   105
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
         MaxLength       =   4
         Button          =   -1  'True
         Caption         =   "CABANG"
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Left            =   375
         TabIndex        =   4
         Top             =   465
         Width           =   3075
         _ExtentX        =   5424
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
         MaxLength       =   20
         Button          =   -1  'True
         Caption         =   "REKENING"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekening 
         Height          =   330
         Left            =   3465
         TabIndex        =   5
         Top             =   465
         Width           =   3645
         _ExtentX        =   6429
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5580
      Width           =   9645
      _ExtentX        =   17013
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
         Left            =   2655
         TabIndex        =   6
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
         Picture         =   "MstSARekening.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3825
         TabIndex        =   7
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
         Picture         =   "MstSARekening.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   7365
         TabIndex        =   8
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
         Picture         =   "MstSARekening.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1605
         TabIndex        =   9
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
         Picture         =   "MstSARekening.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   540
         TabIndex        =   10
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
         Picture         =   "MstSARekening.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   8445
         TabIndex        =   11
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
         Picture         =   "MstSARekening.frx":0C9A
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   105
         TabIndex        =   12
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
         Picture         =   "MstSARekening.frx":0D40
      End
   End
End
Attribute VB_Name = "MstSARekening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim dbRekening As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim lEdit As Boolean
Dim nPos As SisPos
Dim cSQL As String

Private Sub cGudang_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "Cabang", "Kode", cGudang, "Kode,Keterangan", "And Kode like '" & cGudang.Text & "%'")
  If Not dbData.eof Then
    cNamaGudang.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cGudang_Validate(Cancel As Boolean)
  If cGudang.LastKey = 13 Then
    cGudang_ButtonClick
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  cGudang.SetFocus
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  lEdit = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  cGudang.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  cGudang.SetFocus
End Sub

Private Sub DeleteData()
  If MsgBox("Data Benar-benar akan Dihapus ?", vbYesNo + vbInformation) = vbYes Then
    objData.Delete GetDSN, "saldorekening", "Rekening", sisAssign, cRekening.Text, "And cabang ='" & cGudang.Text & "'"
    GetSQL
  End If
  initvalue
  GetEdit False
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    initvalue
    GetEdit False
  End If
End Sub

Private Sub cmdPreview_Click()
Dim vaJoin
Dim cField As String
  
  cField = "c.keterangan as cabang,s.Rekening,r.Keterangan,r.Jenis,s.Awal"
  vaJoin = Array("Left Join Rekening r on r.Kode = s.Rekening", _
                  "Left join Cabang c on c.kode = s.cabang")
  Set dbData = objData.Browse(GetDSN, "SaldoRekening s", cField, , , , , "s.cabang,s.Rekening", vaJoin)
  If Not dbData.eof Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    GetRpt
  Else
    MsgBox "Data Tidak Ada Untuk Ditampilkan..!!", vbInformation, "Preview"
  End If
End Sub

Private Sub GetRpt()
    With FrmRPT
    .AddPageHeader "LAPORAN SALDO AWAL REKENING", tdbHalignCenter, , , , dbArial, 12, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "CABANG", , , , 10, , , , , , , , , , , , , 5
    .AddTableHeader "REKENING", , , , 15
    .AddTableHeader "NAMA REKENING"
    .AddTableHeader "JENIS", , , , 5
    .AddTableHeader "AWAL", , , , 15
    
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .Preview vaArray, True
  End With
End Sub

Private Sub cmdSimpan_Click()
  If ValidSaving() Then
    If MsgBox("Data benar-benar sudah Valid", vbYesNo + vbInformation) = vbYes Then
      objData.Update GetDSN, "SaldoRekening", "Cabang ='" & cGudang.Text & "' And Rekening = '" & cRekening.Text & "' ", _
                              Array("Cabang", "Rekening", "AwalTahun", "Awal"), _
                              Array(cGudang.Text, cRekening.Text, nAwal.Value, nAwal.Value)
    
      GetSQL
    Else
      nAwal.SetFocus
      Exit Sub
    End If
    initvalue
    GetEdit False
  End If
End Sub

Static Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cGudang.Text, "Cabang Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cGudang.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cRekening.Text, "Rekening Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cRekening.SetFocus
    Exit Function
  End If
  
  If nAwal.Value = 0 Then
    MsgBox "Saldo awal tidak boleh nol..", vbInformation
    ValidSaving = False
    nAwal.SetFocus
    Exit Function
  End If
End Function

Private Sub cRekening_ButtonClick()
  Set dbRekening = objData.Pick(GetDSN, "Rekening", "Kode", cRekening, "Kode,Keterangan,Jenis", " and Jenis = 'D'")
  If Not dbRekening.eof Then
    cNamaRekening.Text = GetNull(dbRekening!Keterangan, "")
  End If
End Sub

Private Sub cRekening_Validate(Cancel As Boolean)
  If cRekening.LastKey = 13 Then
     cRekening_ButtonClick
      Set dbData = objData.Browse(GetDSN, "SaldoRekening", , "Cabang", sisAssign, cGudang.Text, "And Rekening='" & cRekening.Text & "'")
      If Not dbData.eof Then
        If nPos = Add Then
          MsgBox "Data Sudah ada. Silahkan Ulangi pengisian", vbOKOnly + vbInformation, "Saldo Awal Rekening"
          Cancel = True
          cRekening.Default
          cRekening.SetFocus
          Exit Sub
        End If
        GetMemory
        If nPos = Delete Then DeleteData
      ElseIf dbData.eof And nPos <> Add Then
        MsgBox "Data tidak ada. Silahkan ulangi pengisian", vbInformation
        Cancel = True
        cRekening.Default
        cRekening.SetFocus
        Exit Sub
      End If
  End If
End Sub

Private Sub GetMemory()
Dim cField As String
Dim vaJoin

  cField = "s.Rekening,s.Cabang,s.awaltahun,s.awal, r.Keterangan as NamaRekening,r.Jenis, c.Keterangan as NamaGudang"
  vaJoin = Array("Left join Rekening r on r.Kode = s.Rekening", _
                 "Left join Cabang c on c.Kode = s.Cabang")
  Set dbData = objData.Browse(GetDSN, "SaldoRekening s", cField, "s.Cabang", sisAssign, cGudang.Text, "And s.Rekening ='" & cRekening.Text & "'", "s.cabang,s.Rekening", vaJoin)
  If dbData.RecordCount > 0 Then
    cGudang.Text = GetNull(dbData!Cabang, "")
    cNamaGudang.Text = GetNull(dbData!NamaGudang, "")
    cRekening.Text = GetNull(dbData!Rekening, "")
    cNamaRekening.Text = GetNull(dbData!NamaRekening, "")
    nAwal.Value = GetNull(dbData!Awal, "")
  End If
End Sub

Private Sub DataGrid1_Click()
  lClick = True
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Then
    lClick = True
  End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If lClick And Not lEdit Then
    cGudang.Text = DataGrid1.Columns(0)
    cRekening.Text = DataGrid1.Columns(1)
    GetMemory
  End If
  lClick = False
End Sub

Private Sub initvalue()
  cGudang.Default
  cNamaGudang.Default
  cRekening.Default
  cNamaRekening.Default
  nAwal.Value = 0
End Sub

Private Sub Form_Load()
Dim n As Single

  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 4
  Set DataGrid1.Array = vaArray
  DataGrid1.ReBind
  
  CenterForm Me
  Me.Top = 0
  GetSQL
  initvalue
  GetEdit False
  
  TabIndex cGudang, n
  TabIndex cRekening, n
  TabIndex nAwal, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdPreview, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub GetSQL()
Dim cField As String
Dim n As Integer
Dim nJumlah As Double

  vaArray.ReDim 0, -1, 0, 4
  cField = "s.Rekening,s.Cabang,s.awal, r.Keterangan,r.Jenis"
  Set dbData = objData.Browse(GetDSN, "SaldoRekening s", cField, , , , " 1=1 and r.jenis = 'D'", "s.Cabang,s.Rekening", _
                              Array("Left join Rekening r on r.Kode = s.Rekening"))
  If Not dbData.eof Then
    dbData.MoveFirst
    nJumlah = 0
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull(dbData!Cabang)
      vaArray(n, 1) = GetNull(dbData!Rekening)
      vaArray(n, 2) = GetNull(dbData!Keterangan)
      vaArray(n, 3) = GetNull(dbData!Jenis)
      vaArray(n, 4) = GetNull(dbData!Awal)
      
      nJumlah = nJumlah + vaArray(n, 4)
      dbData.MoveNext
    Loop
    DataGrid1.Columns(4).FooterText = Format(nJumlah, "###,###,##0.00")
  End If
  Set DataGrid1.Array = vaArray
  DataGrid1.ReBind
  DataGrid1.Refresh
End Sub
