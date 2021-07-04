VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trJurnalUmum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI JURNAL UMUM"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   11805
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4065
      Left            =   0
      Top             =   1260
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   7170
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
      Begin BiSAButtonProject.BiSAButton cmdAddDetail 
         Height          =   330
         Left            =   11370
         TabIndex        =   0
         Top             =   90
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         Caption         =   "OK"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16579821
      End
      Begin BiSANumberBoxProject.BiSANumberBox nKredit 
         Height          =   330
         Left            =   9795
         TabIndex        =   1
         Top             =   90
         Width           =   1560
         _ExtentX        =   2752
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
      Begin BiSANumberBoxProject.BiSANumberBox nDebet 
         Height          =   330
         Left            =   8160
         TabIndex        =   2
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   4890
         TabIndex        =   3
         Top             =   90
         Width           =   3255
         _ExtentX        =   5741
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
         FontName        =   "Tahoma"
         MaxLength       =   150
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
         Height          =   315
         Left            =   2025
         TabIndex        =   4
         Top             =   90
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   556
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Left            =   480
         TabIndex        =   5
         Top             =   90
         Width           =   1545
         _ExtentX        =   2725
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
         FontName        =   "Tahoma"
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
      Begin BiSANumberBoxProject.BiSANumberBox nUrut 
         Height          =   330
         Left            =   75
         TabIndex        =   6
         Top             =   90
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   582
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3045
         Left            =   60
         TabIndex        =   7
         Top             =   465
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   5371
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No."
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Rekening"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama Rekening"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Keterangan Buku Besar"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Debet"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Kredit"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "FormatText Event"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=661"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2672"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=5159"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5080"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=5662"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=5583"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2963"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2884"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2858"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2778"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=112,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFCFCED&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=Tahoma"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
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
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Named:id=33:Normal"
         _StyleDefs(62)  =   ":id=33,.parent=0"
         _StyleDefs(63)  =   "Named:id=34:Heading"
         _StyleDefs(64)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   ":id=34,.wraptext=-1"
         _StyleDefs(66)  =   "Named:id=35:Footing"
         _StyleDefs(67)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   "Named:id=36:Selected"
         _StyleDefs(69)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(70)  =   "Named:id=37:Caption"
         _StyleDefs(71)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(72)  =   "Named:id=38:HighlightRow"
         _StyleDefs(73)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=39:EvenRow"
         _StyleDefs(75)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(76)  =   "Named:id=40:OddRow"
         _StyleDefs(77)  =   ":id=40,.parent=33"
         _StyleDefs(78)  =   "Named:id=41:RecordSelector"
         _StyleDefs(79)  =   ":id=41,.parent=34,.alignment=3"
         _StyleDefs(80)  =   "Named:id=42:FilterBar"
         _StyleDefs(81)  =   ":id=42,.parent=33,.alignment=3"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nTotDebet 
         Height          =   390
         Left            =   4905
         TabIndex        =   8
         Top             =   3570
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   688
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         ForeColor       =   -2147483635
         Caption         =   "DEBET"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nTotKredit 
         Height          =   390
         Left            =   8265
         TabIndex        =   9
         Top             =   3570
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   688
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         ForeColor       =   -2147483635
         Caption         =   "KREDIT"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1260
      Left            =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   2223
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   225
         TabIndex        =   10
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "Tanggal"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSATextBoxProject.BiSATextBox cFaktur 
         Height          =   330
         Left            =   225
         TabIndex        =   11
         Top             =   465
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   582
         Text            =   "1234567890"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         MaxLength       =   20
         Caption         =   "No Faktur"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSATextBoxProject.BiSATextBox cKeteranganJurnal 
         Height          =   330
         Left            =   225
         TabIndex        =   12
         Top             =   810
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   582
         Text            =   "1234567890123456789012345678901234567890"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         MaxLength       =   50
         Caption         =   "Keterangan"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Top             =   5310
      Width           =   11805
      _ExtentX        =   20823
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
         Left            =   2235
         TabIndex        =   13
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
         Picture         =   "trJurnalUmum.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3405
         TabIndex        =   14
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
         Picture         =   "trJurnalUmum.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9435
         TabIndex        =   15
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
         Picture         =   "trJurnalUmum.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1185
         TabIndex        =   16
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
         Picture         =   "trJurnalUmum.frx":083F
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   120
         TabIndex        =   17
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
         Picture         =   "trJurnalUmum.frx":096B
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10515
         TabIndex        =   18
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
         Picture         =   "trJurnalUmum.frx":0B16
      End
   End
End
Attribute VB_Name = "trJurnalUmum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim lEdit As Boolean
Dim nPos As SisPos

Private Sub GetFaktur()
Dim cAwal As String

  cAwal = "JU/" & aCfg(msKodeCabang) & "/" & Format(dTgl.Value, "yymmdd")
  If Trim(cFaktur.Text) = "" Then
    Set dbData = objData.Browse(GetDSN, "TotJurnal", "Max(Faktur) Faktur", "Faktur", sisPrefix, cAwal)
    If dbData.eof Then
      cFaktur.Text = "1"
    Else
      cFaktur.Text = Val(Mid(GetNull(dbData!Faktur), Len(cAwal) + 1)) + 1
    End If
  End If
  cFaktur.Text = cAwal & Padl(Trim(cFaktur.Text), cFaktur.MaxLength - Len(cAwal), "0")
End Sub

Private Sub cFaktur_Validate(Cancel As Boolean)
  If cFaktur.LastKey = 13 Then
      GetFaktur
      Set dbData = objData.Browse(GetDSN, "TOTJURNAL", , "faktur", sisAssign, cFaktur.Text)
      If Not dbData.eof Then
        If nPos = Add Then 'Add
          MsgBox "Data Sudah Ada, silahkan ulangi pengisian  !", vbInformation
          Cancel = True
          initvalue
          cFaktur.SetFocus
          Exit Sub
        End If
        GetMemory
        If nPos = Delete Then DeleteTransaksi
      ElseIf dbData.eof And nPos <> Add Then
        MsgBox "Data tidak Ada....!", vbInformation
        Cancel = True
        initvalue
        cFaktur.SetFocus
        Exit Sub
      End If
  End If
End Sub

Private Sub GetMemory()
Dim n As Single, cSQL As String
  
  cKeteranganJurnal.Text = dbData!Keterangan
  
  vaArray.ReDim 0, -1, 0, 5
  cSQL = "Select j.Rekening,j.Keterangan as KeteranganBukuBesar,j.Debet,j.Kredit,r.Keterangan as NamaRekening"
  cSQL = cSQL & " From JURNAL j"
  cSQL = cSQL & " Left join Rekening r on r.Kode = j.Rekening"
  cSQL = cSQL & " Where Faktur = '" & cFaktur.Text & "'"
  cSQL = cSQL & " Order By j.rekening"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull((dbData!Rekening), "")
      vaArray(n, 2) = GetNull((dbData!NamaRekening), "")
      vaArray(n, 3) = GetNull((dbData!KeteranganBukuBesar), "")
      vaArray(n, 4) = GetNull(dbData!Debet)
      vaArray(n, 5) = GetNull(dbData!Kredit)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  SUMJUMLAH
End Sub

Private Sub cmdAdd_Click()
  initvalue
  GetEdit True
  nPos = Add
  dTgl.SetFocus
End Sub

Private Sub cmdAddDetail_Click()
Dim n As Integer

  If Trim(cRekening.Text) = "" Then
    MsgBox "Rekening harus diisi...", vbInformation
    cRekening.SetFocus
    Exit Sub
  End If
  
'  If nDebet.Value < 0 Then
'    MsgBox "Nilai Debet tidak valid...", vbInformation
'    nDebet.SetFocus
'    Exit Sub
'  End If
'
'  If nKredit.Value < 0 Then
'    MsgBox "Nilai Kredit tidak valid...", vbInformation
'    nKredit.SetFocus
'    Exit Sub
'  End If
  
'  If nDebet.Value <= 0 And nKredit.Value <= 0 Then
'    MsgBox "Nilai Debet atau Kredit harus diisi...", vbInformation
'    nDebet.SetFocus
'    Exit Sub
'  End If
  
'  If nDebet.Value > 0 And nKredit.Value > 0 Then
'    MsgBox "Nilai Debet atau Kredit tidak boleh diisi dua-duanya (Harus salah satu !)...", vbInformation
'    nDebet.SetFocus
'    Exit Sub
'  End If
  
  If nUrut.Value > (vaArray.UpperBound(1) + 1) Then
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = nUrut.Value - 1
  ElseIf vaArray.UpperBound(1) = -1 Then
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = nUrut.Value - 1
  Else
    n = nUrut.Value - 1
  End If
  
  vaArray(n, 0) = n + 1
  vaArray(n, 1) = cRekening.Text
  vaArray(n, 2) = cNama.Text
  vaArray(n, 3) = cKeterangan.Text
  vaArray(n, 4) = nDebet.Value
  vaArray(n, 5) = nKredit.Value
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  
  SUMJUMLAH
  Initdetail
  cRekening.SetFocus
  Exit Sub
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  initvalue
  GetEdit True
  nPos = Edit
  dTgl.SetFocus
End Sub

Private Sub cmdHapus_Click()
  initvalue
  GetEdit True
  nPos = Delete
  dTgl.SetFocus
End Sub

Private Sub DeleteTransaksi()
  If MsgBox("Data Benar-benar Dihapus ?", vbQuestion + vbYesNo) = vbYes Then
    
      objData.Delete GetDSN, "jurnal", "faktur", sisAssign, cFaktur.Text
      objData.Delete GetDSN, "totjurnal", "Faktur", sisAssign, cFaktur.Text
      objData.Delete GetDSN, "BukuBesar", "Faktur", sisAssign, cFaktur.Text
    
  End If
  initvalue
  GetEdit False
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    GetEdit False
    initvalue
  End If
End Sub

Private Sub lCekFaktur()
Dim lCek As Boolean

  lCek = False
  Do While Not lCek
    Set dbData = objData.Browse(GetDSN, "TotJurnal", "Faktur", "Faktur", sisAssign, cFaktur.Text)
    If Not dbData.eof Then
      cFaktur.Default
      GetFaktur
      lCek = False
    Else
      lCek = True
    End If
  Loop
End Sub

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue
Dim n As Single

'  If nTotDebet.Value <> nTotKredit.Value Then
'    MsgBox "Jurnal Tidak Balance, Transaksi Tidak Bisa Disimpan", vbInformation
'    nUrut.SetFocus
'    Exit Sub
'  End If
  
  If ValidSaving() Then
    If MsgBox("Data benar-benar sudah VALID ?", vbYesNo) = vbYes Then
        
        
        If nPos = Add Then lCekFaktur
        objData.Delete GetDSN, "TOTJURNAL", "Faktur", sisAssign, cFaktur.Text
        objData.Update GetDSN, "TOTJURNAL", "Faktur = '" & cFaktur.Text & "'", Array("Faktur", "Tgl", "Keterangan", "Username"), _
                                             Array(cFaktur.Text, dTgl.Value, cKeteranganJurnal.Text, cusername)
        
        objData.Delete GetDSN, "JURNAL", "Faktur", sisAssign, cFaktur.Text
        vaField = Array("Faktur", "Tgl", "Rekening", "Keterangan", "Debet", "Kredit")
        For n = 0 To vaArray.UpperBound(1)
          vaValue = Array(cFaktur.Text, dTgl.Value, vaArray(n, 1), vaArray(n, 3), vaArray(n, 4), vaArray(n, 5))
          objData.Add GetDSN, "JURNAL", vaField, vaValue
        Next
    
        UpdRekJurnal objData, cFaktur.Text
        
        
      Else
        nUrut.SetFocus
        Exit Sub
      End If
    End If
    initvalue
    GetEdit False
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cFaktur.Text, "Kode Voucher Harus Diisi, Ulangi Pengisian") Then
    cFaktur.SetFocus
    ValidSaving = False
    Exit Function
  End If
End Function

'Function ExportReport(ByVal vaArray As XArrayDB)
'Dim XLA As Excel.Application
'Dim XLW As Excel.Workbook
'Dim XLS As Excel.Worksheet
'Dim i As Integer
'Dim j As Integer
'Dim cNamaFile As String
'
'    Set XLA = New Excel.Application
'    Set XLW = XLA.Workbooks.Add
'    Set XLS = XLW.Worksheets(1)
'
'    XLS.Cells(1, 1) = "No"
'    XLS.Cells(1, 2) = "Kode"
'    XLS.Cells(1, 3) = "Nama Rekening"
'    XLS.Cells(1, 4) = "Keterangan"
'    XLS.Cells(1, 5) = "Debet"
'    XLS.Cells(1, 6) = "Kredit"
'
'    For i = 3 To vaArray.UpperBound(1) + 3
'        For j = 1 To vaArray.UpperBound(2) + 1
'            XLS.Cells(i, j) = vaArray(i - 3, j - 1)
'        Next j
'    Next i
'    cNamaFile = "c:\BakiDebetPersektor" & Format(Now, "dd-mm-yy hhmmss") & ".xls"
'    XLS.SaveAs cNamaFile
'    XLW.Close False
'    Set XLW = Nothing
'    Set XLA = Nothing
'    MsgBox "Export disimpan dengan nama" & cNamaFile
'End Function

Private Sub cRekening_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Rekening", "Kode,Keterangan,Jenis", "Kode", sisContent, cRekening.Text, "And Jenis='D'", "Kode")
  cRekening.Text = cRekening.Browse(dbData)
  If Not dbData.eof Then
    cNama.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Then
    Cancel = True
    dTgl.SetFocus
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me, True
  initvalue
  GetEdit False
  
  TabIndex dTgl, n
  TabIndex cFaktur, n
  TabIndex cKeteranganJurnal, n
  
  TabIndex nUrut, n
  TabIndex cRekening, n
  TabIndex cNama, n
  TabIndex cKeterangan, n
  TabIndex nDebet, n
  TabIndex nKredit, n
  TabIndex cmdAddDetail, n

  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub SUMJUMLAH()
Dim n

  nTotDebet.Value = 0
  nTotKredit.Value = 0
  For n = 0 To vaArray.UpperBound(1)
    nTotDebet.Value = nTotDebet.Value + vaArray(n, 4)
    nTotKredit.Value = nTotKredit.Value + vaArray(n, 5)
  Next
End Sub

Private Sub Initdetail()
  cRekening.Default
  cNama.Default
  cKeterangan.Default
  nDebet.Value = 0
  nKredit.Value = 0
  nUrut.Value = vaArray.UpperBound(1) + 2
End Sub

Private Sub initvalue()
  cFaktur.Default
  dTgl.Value = Date
  cKeteranganJurnal.Default
  nTotDebet.Value = 0
  nTotKredit.Value = 0
  vaArray.ReDim 0, -1, 0, 5
  Set TDBGrid1.Array = vaArray
  TDBGrid1.Refresh
  TDBGrid1.ReBind
  
  Initdetail
End Sub

Private Sub nUrut_Change()
  If nUrut.Value <= 0 Then
    nUrut.Value = vaArray.UpperBound(1) + 2
  End If
End Sub

Private Sub nUrut_Validate(Cancel As Boolean)
  
'  If nUrut.LastKey = 13 Then
'    If nUrut.Value > 0 And vaArray.UpperBound(1) >= 0 Then
'      cRekening.Text = vaArray(nUrut.Value - 1, 1)
'      cNama.Text = vaArray(nUrut.Value - 1, 2)
'      cKeterangan.Text = vaArray(nUrut.Value - 1, 3)
'      nDebet.Value = vaArray(nUrut.Value - 1, 4)
'      nKredit.Value = vaArray(nUrut.Value - 1, 5)
'    End If
'  End If

  If nUrut.Value - 1 <= vaArray.UpperBound(1) And nUrut.Value >= 1 Then
    cRekening.Text = vaArray(nUrut.Value - 1, 1)
    cNama.Text = vaArray(nUrut.Value - 1, 2)
    cKeterangan.Text = vaArray(nUrut.Value - 1, 3)
    nDebet.Value = vaArray(nUrut.Value - 1, 4)
    nKredit.Value = vaArray(nUrut.Value - 1, 5)
  ElseIf nUrut.Value - 2 > vaArray.UpperBound(1) Or nUrut.Value <= 0 Then
    nUrut.Value = vaArray.UpperBound(1) + 2
  End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  If Val(Value) = 0 Then
    Value = ""
  Else
    Value = Format(Value, "###,###,###,###,##0.00")
  End If
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  lEdit = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    TDBGrid1.Delete
    
    For n = 0 To vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
    Next
    
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    nUrut.Value = vaArray.UpperBound(1) + 2
  End If
  SUMJUMLAH
End Sub
