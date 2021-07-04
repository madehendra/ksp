VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trHapusAngsuran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HAPUS PENCAIRAN DAN ANGSURAN"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11670
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3720
      Left            =   0
      Top             =   1455
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   6562
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
         Height          =   3600
         Left            =   75
         TabIndex        =   0
         Top             =   60
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   6350
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Tanggal"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No Transaksi"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Angs Bunga"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "###,###,###,###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Angs Pokok"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Denda"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Total Angs"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "User Name"
         Columns(7).DataField=   ""
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=476"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=397"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1746"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1667"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4233"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4154"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2487"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2408"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2461"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2381"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2090"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2011"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=3598"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3519"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(25)  =   "Splits(0).Style:id=75,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=76,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=77,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=78,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=80,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=75"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=76"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=77"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=79"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=75"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=76"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=77"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=79"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=75"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=76"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=77"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=79"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=75,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=76"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=77"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=79"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=75,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=76"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=77"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=79"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=75,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=76"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=77"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=79"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=62,.parent=75,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=76"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=77"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=79"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=90,.parent=75"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=87,.parent=76"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=88,.parent=77"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=89,.parent=79"
         _StyleDefs(69)  =   "Named:id=33:Normal"
         _StyleDefs(70)  =   ":id=33,.parent=0"
         _StyleDefs(71)  =   "Named:id=34:Heading"
         _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   ":id=34,.wraptext=-1"
         _StyleDefs(74)  =   "Named:id=35:Footing"
         _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   "Named:id=36:Selected"
         _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=37:Caption"
         _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(80)  =   "Named:id=38:HighlightRow"
         _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(82)  =   "Named:id=39:EvenRow"
         _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(84)  =   "Named:id=40:OddRow"
         _StyleDefs(85)  =   ":id=40,.parent=33"
         _StyleDefs(86)  =   "Named:id=41:RecordSelector"
         _StyleDefs(87)  =   ":id=41,.parent=34"
         _StyleDefs(88)  =   "Named:id=42:FilterBar"
         _StyleDefs(89)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   2566
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
      Begin VB.OptionButton Opt 
         Caption         =   "Pencairan Kredit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3495
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   60
         Width           =   2010
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Angsuran Kredit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1710
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   60
         Width           =   2010
      End
      Begin BiSANumberBoxProject.BiSANumberBox nPlafond 
         Height          =   300
         Left            =   7785
         TabIndex        =   1
         Top             =   420
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   529
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   300
         Left            =   7785
         TabIndex        =   2
         Top             =   90
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   529
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   300
         Left            =   3690
         TabIndex        =   3
         Top             =   405
         Width           =   390
         _ExtentX        =   688
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
         Height          =   300
         Left            =   2130
         TabIndex        =   4
         Top             =   405
         Width           =   720
         _ExtentX        =   1270
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
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   405
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   "No Rekening"
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
      Begin BiSATextBoxProject.BiSATextBox cUrut 
         Height          =   300
         Left            =   2880
         TabIndex        =   6
         Top             =   405
         Width           =   795
         _ExtentX        =   1402
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   300
         Left            =   90
         TabIndex        =   7
         Top             =   735
         Width           =   5040
         _ExtentX        =   8890
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
         BackColor       =   12632256
         Enabled         =   0   'False
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   300
         Left            =   90
         TabIndex        =   8
         Top             =   1050
         Width           =   5745
         _ExtentX        =   10134
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
         BackColor       =   12632256
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
      Begin BiSANumberBoxProject.BiSANumberBox nSukuBunga 
         Height          =   300
         Left            =   7785
         TabIndex        =   9
         Top             =   750
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   529
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
      Begin BiSANumberBoxProject.BiSANumberBox nLama 
         Height          =   300
         Left            =   7785
         TabIndex        =   10
         Top             =   1080
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   529
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
      Begin VB.Label Label3 
         Caption         =   "Transaksi"
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
         Left            =   135
         TabIndex        =   15
         Top             =   120
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "% p.a"
         Height          =   210
         Left            =   10275
         TabIndex        =   12
         Top             =   795
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Bulan"
         Height          =   210
         Left            =   10230
         TabIndex        =   11
         Top             =   1110
         Width           =   900
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5160
      Width           =   11670
      _ExtentX        =   20585
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
      Begin BiSAButtonProject.BiSAButton cmdDelete 
         Height          =   435
         Left            =   9435
         TabIndex        =   13
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
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
         Picture         =   "trHapusAngsuran.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10515
         TabIndex        =   14
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
         Picture         =   "trHapusAngsuran.frx":0416
      End
   End
End
Attribute VB_Name = "trHapusAngsuran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim cRekening As String
Dim nBarisClick As Integer
Dim lClick As Boolean

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cFakturRealisasi As String

  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  Set dbData = objData.Browse(GetDSN, "Debitur", "Rekening,Status,Faktur", "Rekening", sisAssign, cRekening)
  If Not dbData.eof Then
    If GetNull(dbData!status) = "1" Then
      MsgBox "Kredit Sudah Lunas.....", vbOKOnly + vbInformation
      Cancel = True
      initvalue
      cFrekuensi.SetFocus
      Exit Sub
    End If
    cFakturRealisasi = GetNull(dbData!Faktur, "")
    GetMemory
    'Jika pencairan kredit
    If Opt(1).Value = True Then
      If CekAngsuran(cRekening) Then
        If MsgBox("Akan menghapus pencairan Kredit..?", vbYesNo + vbInformation) = vbYes Then
          DeletePencairan cRekening, cFakturRealisasi
        End If
      End If
      initvalue
      Cancel = True
      cFrekuensi.SetFocus
      Exit Sub
    End If
  Else
    MsgBox "No. Rekening Tidak Ditemukan, Ulangi Pengisian", vbOKOnly + vbInformation
    Cancel = True
    cFrekuensi.SetFocus
    Exit Sub
  End If
End Sub

Private Sub DeletePencairan(ByVal cRek As String, ByVal cFakturRealisasi As String)
  'hapus buku besar
  objData.Delete GetDSN, "BukuBesar", "Faktur", sisAssign, cFakturRealisasi, "And Status= '" & vbTrigger.msRealisasiKredit & "'"
  
  'Edit di debitur
  objData.Edit GetDSN, "Debitur", "Rekening='" & cRek & "'", Array("Faktur", "Statuspencairan"), Array("", "")
  MsgBox "Pengahapusan telah dilakukan..", vbOKOnly + vbInformation
  initvalue
  Opt(0).SetFocus
  Exit Sub
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  'Cek Register Nasabah
  If Not CheckData(cCabang.Text, "Kode Cabang Tidak Terisi..!") Then
    ValidSaving = False
    Exit Function
  End If
  
  'Cek Golongan Nasabah
  If Not CheckData(cGolongan.Text, "Kode Golongan Tidak Terisi..!") Then
    ValidSaving = False
    Exit Function
  End If
  
  If Not CheckData(cUrut.Text, "Kode Urut Tidak Terisi..!") Then
    ValidSaving = False
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Kode Frekuensi Tidak Terisi..!") Then
    ValidSaving = False
    Exit Function
  End If
End Function

Private Sub GetMemory()
Dim vaJoin
Dim cField As String

  cField = "d.Tgl,d.Plafond,d.Lama,d.SukuBunga, r.Nama, r.Alamat"
  vaJoin = Array("Left Join Registernasabah r on r.Kode=d.Kode")
  Set dbData = objData.Browse(GetDSN, "Debitur d", cField, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    cNama.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    dTgl.Value = GetNull(dbData!Tgl)
    nPlafond.Value = GetNull(dbData!plafond)
    nSukuBunga.Value = GetNull(dbData!SukuBunga)
    nLama.Value = GetNull(dbData!Lama)
    GetDataAngsuran
  End If
End Sub

Private Sub GetDataAngsuran()
Dim nTotal As Double
Dim n As Long
Dim nPokok As Double
Dim nBunga As Double
Dim nDenda As Double
  
  nPokok = 0
  nBunga = 0
  nDenda = 0
  nTotal = 0
  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "Angsuran", , "Rekening", sisAssign, cRekening, , "Tgl")
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = 0
      vaArray(n, 1) = GetNull(dbData!Tgl)
      vaArray(n, 2) = GetNull(dbData!Faktur, "")
      vaArray(n, 3) = GetNull(dbData!bunga)
      vaArray(n, 4) = GetNull(dbData!pokok)
      vaArray(n, 5) = GetNull(dbData!denda)
      vaArray(n, 6) = GetNull(dbData!Total)
      vaArray(n, 7) = GetNull(dbData!UserName, "")
      
      nPokok = nPokok + vaArray(n, 3)
      nBunga = nBunga + vaArray(n, 4)
      nDenda = nDenda + vaArray(n, 5)
      nTotal = nTotal + vaArray(n, 6)
      dbData.MoveNext
    Loop
  End If
  TDBGrid1.Columns(3).FooterText = Format(nBunga, "###,###,###,###,##0.00")
  TDBGrid1.Columns(4).FooterText = Format(nPokok, "###,###,###,###,##0.00")
  TDBGrid1.Columns(5).FooterText = Format(nDenda, "###,###,###,###,##0.00")
  TDBGrid1.Columns(6).FooterText = Format(nTotal, "###,###,###,###,##0.00")
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "GolonganKredit", "Kode,Keterangan", "Kode", sisContent, cGolongan.Text)
  cGolongan.Text = cGolongan.Browse(dbData)
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  cGolongan_ButtonClick
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim n As Integer

  If ValidSaving Then
    If MsgBox("Data benar-benar akan dihapus ?", vbYesNo + vbExclamation) = vbYes Then
      For n = 0 To vaArray.UpperBound(1)
        'If vaArray(n, 0) = -1 Then
        If vaArray(n, 0) = 1 Then
          objData.Delete GetDSN, "Angsuran", "Faktur", sisAssign, vaArray(n, 2)
          objData.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msAngsuranKredit, "And Faktur='" & vaArray(n, 2) & "'"
        End If
      Next
      GetDataAngsuran
      cCabang.SetFocus
      Exit Sub
    End If
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub initvalue()
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
  dTgl.Value = Date
  nPlafond.Value = 0
  nLama.Value = 0
  nSukuBunga.Value = 0
  
  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 7
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me, True
  initvalue
  cCabang.Text = aCfg(msKodeCabang, "")
  Opt(0).Value = True
  
  TabIndex Opt(0), n
  TabIndex Opt(1), n
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cmdDelete, n
  TabIndex cmdKeluar, n
End Sub

Private Sub Opt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub


'Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
'Dim dDate As Date
'
'  If vaArray(nBarisClick, ColIndex) = -1 Then
'    dDate = vaArray(nBarisClick, 1)
'    If Not IsInPeriod(dDate) Then
'      vaArray(nBarisClick, ColIndex) = 0
'    End If
'  End If
'  TDBGrid1.ReBind
'End Sub

Private Sub TDBGrid1_Click()
  lClick = True
End Sub

Static Function CekAngsuran(ByVal cRek As String) As Boolean
  CekAngsuran = True
  Set dbData = objData.Browse(GetDSN, "Angsuran", "Count(Rekening) as Jumlah", "Rekening", sisAssign, cRek, "Group By Rekening")
  If Not dbData.eof Then
    MsgBox "Penghapusan Pencairan Kredit tidak bisa dilakukan karena sudah pernah melakukan Transaksi Angsuran", vbOKOnly + vbInformation
    CekAngsuran = False
    Opt(0).SetFocus
    initvalue
    Exit Function
  End If
  
  Set dbData = objData.Browse(GetDSN, "Debitur", "StatusPencairan", "Rekening", sisAssign, cRek)
  If Not dbData.eof Then
    If GetNull(dbData!statuspencairan) <> "1" Then
      MsgBox "Pencairan Kredit belum dilakukan....", vbOKCancel + vbInformation
      CekAngsuran = False
      Opt(0).SetFocus
      initvalue
      Exit Function
    End If
  End If
End Function

Private Sub GetMark()
Dim n As Double

  n = GetNull(TDBGrid1.Bookmark)
  vaArray(n, 0) = IIf(GetNull(vaArray(n, 0)) <= 0, 1, 0)
  TDBGrid1.Columns(0) = vaArray(n, 0)
  lClick = False
End Sub

'Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = 32 Then
'    GetMark
'    'SendKeys "{Down}"
'  End If
'End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If lClick Then
    GetMark
  End If
End Sub
