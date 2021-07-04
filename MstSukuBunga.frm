VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form MstSukuBunga 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER SUKU BUNGA PROGRESSIF"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6795
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3810
      Left            =   0
      Top             =   990
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6720
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
      Begin BiSANumberBoxProject.BiSANumberBox nNo 
         Height          =   330
         Left            =   75
         TabIndex        =   0
         Top             =   75
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         Decimals        =   0
         DecimalPoint    =   ""
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
      Begin TrueOleDBGrid70.TDBGrid DataGrid1 
         Height          =   3330
         Left            =   75
         TabIndex        =   1
         Top             =   420
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   5874
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "MINIMUM"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "###,###,###,###,##0.00"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "MAXIMUM"
         Columns(2).DataField=   "Keterangan"
         Columns(2).NumberFormat=   "###,###,###,###,##0.00"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "RATE (%)"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=953"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4048"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3969"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=514"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3942"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3863"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2302"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2223"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(53)  =   "Named:id=33:Normal"
         _StyleDefs(54)  =   ":id=33,.parent=0"
         _StyleDefs(55)  =   "Named:id=34:Heading"
         _StyleDefs(56)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   ":id=34,.wraptext=-1"
         _StyleDefs(58)  =   "Named:id=35:Footing"
         _StyleDefs(59)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   "Named:id=36:Selected"
         _StyleDefs(61)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   "Named:id=37:Caption"
         _StyleDefs(63)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(64)  =   "Named:id=38:HighlightRow"
         _StyleDefs(65)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(66)  =   "Named:id=39:EvenRow"
         _StyleDefs(67)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(68)  =   "Named:id=40:OddRow"
         _StyleDefs(69)  =   ":id=40,.parent=33"
         _StyleDefs(70)  =   "Named:id=41:RecordSelector"
         _StyleDefs(71)  =   ":id=41,.parent=34"
         _StyleDefs(72)  =   "Named:id=42:FilterBar"
         _StyleDefs(73)  =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nMin 
         Height          =   330
         Left            =   615
         TabIndex        =   2
         Top             =   75
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         MaxValue        =   999999999999
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
      Begin BiSANumberBoxProject.BiSANumberBox nMax 
         Height          =   330
         Left            =   2925
         TabIndex        =   3
         Top             =   75
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   582
         MaxValue        =   999999999999
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
      Begin BiSANumberBoxProject.BiSANumberBox nBunga 
         Height          =   330
         Left            =   5160
         TabIndex        =   4
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   6300
         TabIndex        =   13
         Top             =   75
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
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
         BackColor       =   16579821
         Picture         =   "MstSukuBunga.frx":0000
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   6795
      _ExtentX        =   11986
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
      Begin BiSATextBoxProject.BiSATextBox cKode 
         Height          =   330
         Left            =   270
         TabIndex        =   5
         Top             =   120
         Width           =   2085
         _ExtentX        =   3678
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
         MaxLength       =   2
         Caption         =   "KODE"
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   270
         TabIndex        =   6
         Top             =   510
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   582
         Text            =   "1234567890123456789012345678901234567890"
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
         MaxLength       =   40
         Caption         =   "KETERANGAN"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   4785
      Width           =   6795
      _ExtentX        =   11986
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2190
         TabIndex        =   7
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
         Picture         =   "MstSukuBunga.frx":01AA
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3360
         TabIndex        =   8
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
         Picture         =   "MstSukuBunga.frx":0434
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   4440
         TabIndex        =   9
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
         Picture         =   "MstSukuBunga.frx":05D3
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1140
         TabIndex        =   10
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
         Picture         =   "MstSukuBunga.frx":09E9
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   75
         TabIndex        =   11
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
         Picture         =   "MstSukuBunga.frx":0B15
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5520
         TabIndex        =   12
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
         Picture         =   "MstSukuBunga.frx":0CC0
      End
   End
End
Attribute VB_Name = "MstSukuBunga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
  If cKode.LastKey = 13 Then
     Set dbData = objData.Browse(GetDSN, "SukuBunga", , "Kode", sisAssign, cKode.Text)
     If Not dbData.eof Then
      If nPos = Add Then
        MsgBox "Data sudah ada...!", vbInformation
        Cancel = True
        initvalue
        cKode.SetFocus
        Exit Sub
      End If
      GetMemory
      If nPos = Delete Then HapusData
    ElseIf dbData.eof And nPos <> Add Then
      MsgBox "Data tidak ada...!", vbInformation
      Cancel = True
      initvalue
      cKode.SetFocus
    End If
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  cKode.SetFocus
End Sub

Private Sub Initdetail()
  nMin.Value = 0
  nMax.Value = 0
  nBunga.Value = 0
  nNo.Value = vaArray.UpperBound(1) + 2
End Sub

Private Sub initvalue()
  cKode.Default
  cKeterangan.Default
  vaArray.ReDim 0, -1, 0, 3
  Set DataGrid1.Array = vaArray
  DataGrid1.ReBind
  Initdetail
End Sub

Private Sub HapusData()
  If MsgBox("Data benar-benar dihapus...?", vbQuestion + vbYesNo) = vbYes Then
    objData.Delete GetDSN, "SukuBunga", "Kode", sisAssign, cKode.Text
    objData.Delete GetDSN, "DetailSukuBunga", "Kode", sisAssign, cKode.Text
  End If
  initvalue
  GetEdit False
End Sub

Private Sub GetMemory()
Dim n As Integer
  
  cKode.Text = GetNull(dbData!Kode, "")
  cKeterangan.Text = GetNull(dbData!Keterangan, "")
  
  vaArray.ReDim 0, -1, 0, 3
  Set dbData = objData.Browse(GetDSN, "DetailSukuBunga", , "Kode", sisAssign, cKode.Text, , "ID")
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!Minimum)
      vaArray(n, 2) = GetNull(dbData!Maximum)
      vaArray(n, 3) = GetNull(dbData!SukuBunga)
      dbData.MoveNext
    Loop
  End If
  Set DataGrid1.Array = vaArray
  DataGrid1.ReBind
  nNo.Value = vaArray.UpperBound(1) + 2
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  initvalue
  GetEdit True
  cKode.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  initvalue
  GetEdit True
  cKode.SetFocus
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Sub cmdOK_Click()
Dim n As Integer

  If nNo.Value > (vaArray.UpperBound(1) + 1) Then
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = nNo.Value - 1
  ElseIf vaArray.UpperBound(1) = -1 Then
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = nNo.Value - 1
  Else
    n = nNo.Value - 1
  End If
  
  vaArray(n, 0) = n + 1
  vaArray(n, 1) = nMin.Value
  vaArray(n, 2) = nMax.Value
  vaArray(n, 3) = nBunga.Value
  
  Set DataGrid1.Array = vaArray
  DataGrid1.ReBind
  DataGrid1.Refresh
  Initdetail
  nMin.SetFocus
  Exit Sub
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim n As Integer

  If ValidSaving() Then
    vaField = Array("Kode", "Keterangan")
    vaValue = Array(cKode.Text, cKeterangan.Text)
    objData.Update GetDSN, "SukuBunga", "Kode='" & cKode.Text & "'", vaField, vaValue
      
    objData.Delete GetDSN, "DetailSukuBunga", "Kode", sisAssign, cKode.Text
    For n = 0 To vaArray.UpperBound(1)
      objData.Add GetDSN, "DetailSukuBunga", Array("Kode", "Minimum", "MAximum", "SukuBunga"), _
                                            Array(cKode.Text, vaArray(n, 1), vaArray(n, 2), vaArray(n, 3))
    Next
    initvalue
    GetEdit False
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cKode.Text, "Kode harus diisi...?") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKeterangan.Text, "Keterangan harus diisi...?") Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
End Function

Private Sub Form_Load()
Dim n As Single

  GetEdit False
  initvalue
  CenterForm Me
  
  TabIndex cKode, n
  TabIndex cKeterangan, n
  TabIndex nNo, n
  TabIndex nMin, n
  TabIndex nMax, n
  TabIndex nBunga, n
  TabIndex cmdOK, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub nNo_Validate(Cancel As Boolean)
  If vaArray.UpperBound(1) >= 0 And nNo.Value - 1 <= vaArray.UpperBound(1) Then
    nMin.Value = vaArray(nNo.Value - 1, 1)
    nMax.Value = vaArray(nNo.Value - 1, 2)
    nBunga.Value = vaArray(nNo.Value - 1, 3)
  End If
End Sub
