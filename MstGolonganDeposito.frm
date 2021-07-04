VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form MstGolonganDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER GOLONGAN DEPOSITO"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11400
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   2490
      Left            =   0
      Top             =   3255
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4392
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
         Height          =   2385
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   4207
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE"
         Columns(0).DataField=   "Kode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KETERANGAN"
         Columns(1).DataField=   "Keterangan"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "LAMA"
         Columns(2).DataField=   "Lama"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "BUNGA"
         Columns(3).DataField=   "BUNGA"
         Columns(3).NumberFormat=   "###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "REKENING"
         Columns(4).DataField=   "RekeningAkuntansi"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "REK BUNGA"
         Columns(5).DataField=   "RekeningBunga"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "REK PAJAK BUNGA"
         Columns(6).DataField=   "RekeningPajakBunga"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1138"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1058"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=6324"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6244"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1349"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1270"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1905"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1826"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2566"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2487"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2990"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2910"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2963"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2884"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
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
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=121,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(65)  =   "Named:id=33:Normal"
         _StyleDefs(66)  =   ":id=33,.parent=0"
         _StyleDefs(67)  =   "Named:id=34:Heading"
         _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   ":id=34,.wraptext=-1"
         _StyleDefs(70)  =   "Named:id=35:Footing"
         _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   "Named:id=36:Selected"
         _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=37:Caption"
         _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(76)  =   "Named:id=38:HighlightRow"
         _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=39:EvenRow"
         _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(80)  =   "Named:id=40:OddRow"
         _StyleDefs(81)  =   ":id=40,.parent=33"
         _StyleDefs(82)  =   "Named:id=41:RecordSelector"
         _StyleDefs(83)  =   ":id=41,.parent=34"
         _StyleDefs(84)  =   "Named:id=42:FilterBar"
         _StyleDefs(85)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   5741
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
      Begin BiSANumberBoxProject.BiSANumberBox nLama 
         Height          =   330
         Left            =   7650
         TabIndex        =   1
         Top             =   810
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   582
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "JANGKA WAKTU"
         CaptionWidth    =   1600
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekPajakBahas 
         Height          =   330
         Left            =   3915
         TabIndex        =   2
         Top             =   1485
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekBahas 
         Height          =   330
         Left            =   3915
         TabIndex        =   3
         Top             =   1140
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekAkuntansi 
         Height          =   330
         Left            =   3915
         TabIndex        =   4
         Top             =   795
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningakuntansi 
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   795
         Width           =   3750
         _ExtentX        =   6615
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
         Caption         =   "Rek. Akuntansi"
         CaptionWidth    =   2000
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
      Begin BiSATextBoxProject.BiSATextBox cKode 
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   90
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   582
         Text            =   "1"
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
         MaxLength       =   1
         Caption         =   "KODE                     D"
         CaptionWidth    =   2000
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
         Left            =   150
         TabIndex        =   7
         Top             =   450
         Width           =   6570
         _ExtentX        =   11589
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
         CaptionWidth    =   2000
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningBahas 
         Height          =   330
         Left            =   150
         TabIndex        =   8
         Top             =   1140
         Width           =   3750
         _ExtentX        =   6615
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
         Caption         =   "Rek. Biaya Bunga"
         CaptionWidth    =   2000
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPajakBahas 
         Height          =   330
         Left            =   150
         TabIndex        =   9
         Top             =   1485
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         Text            =   "12345678901"
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
         Caption         =   "Rek. Pajak Bunga"
         CaptionWidth    =   2000
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningJatuhTempo 
         Height          =   330
         Left            =   3915
         TabIndex        =   10
         Top             =   1830
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningJatuhTempo 
         Height          =   330
         Left            =   150
         TabIndex        =   11
         Top             =   1830
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         Text            =   "12345678901"
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
         Caption         =   "Rek. Jatuh Tempo"
         CaptionWidth    =   2000
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
      Begin BiSATextBoxProject.BiSATextBox cNamaCadanganBahas 
         Height          =   330
         Left            =   3915
         TabIndex        =   12
         Top             =   2175
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSABrowse cCadanganBahas 
         Height          =   330
         Left            =   150
         TabIndex        =   13
         Top             =   2175
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         Text            =   "12345678901"
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
         Caption         =   "Rek. Titipan Bunga"
         CaptionWidth    =   2000
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
         Left            =   7650
         TabIndex        =   22
         Top             =   1155
         Width           =   2505
         _ExtentX        =   4419
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
         Caption         =   "SUKU BUNGA"
         CaptionWidth    =   1600
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekfinalti 
         Height          =   330
         Left            =   3915
         TabIndex        =   23
         Top             =   2520
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSABrowse cRekFinalti 
         Height          =   330
         Left            =   150
         TabIndex        =   24
         Top             =   2520
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         Text            =   "12345678901"
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
         Caption         =   "Rek. Finalty"
         CaptionWidth    =   2000
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
      Begin BiSANumberBoxProject.BiSANumberBox nMinimum 
         Height          =   330
         Left            =   7650
         TabIndex        =   25
         Top             =   1500
         Width           =   3480
         _ExtentX        =   6138
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
         Caption         =   "MIN KENA PAJAK"
         CaptionWidth    =   1600
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
      Begin BiSANumberBoxProject.BiSANumberBox nPajak 
         Height          =   330
         Left            =   7650
         TabIndex        =   26
         Top             =   1845
         Width           =   2505
         _ExtentX        =   4419
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
         Caption         =   "PAJAK BUNGA"
         CaptionWidth    =   1600
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekMaterai 
         Height          =   330
         Left            =   3915
         TabIndex        =   29
         Top             =   2865
         Width           =   3690
         _ExtentX        =   6509
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
      Begin BiSATextBoxProject.BiSABrowse cRekMaterai 
         Height          =   330
         Left            =   150
         TabIndex        =   30
         Top             =   2865
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         Text            =   "12345678901"
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
         Caption         =   "R. Materai Cair Pokok"
         CaptionWidth    =   2000
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
         Caption         =   "% / p.a"
         Height          =   240
         Left            =   10230
         TabIndex        =   28
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   225
         Left            =   10215
         TabIndex        =   27
         Top             =   1905
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "BULAN"
         Height          =   270
         Left            =   10065
         TabIndex        =   14
         Top             =   870
         Width           =   690
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5730
      Width           =   11355
      _ExtentX        =   20029
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
         TabIndex        =   15
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
         Picture         =   "MstGolonganDeposito.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3825
         TabIndex        =   16
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
         Picture         =   "MstGolonganDeposito.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9030
         TabIndex        =   17
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
         Picture         =   "MstGolonganDeposito.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1605
         TabIndex        =   18
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
         Picture         =   "MstGolonganDeposito.frx":083F
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   540
         TabIndex        =   19
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
         Picture         =   "MstGolonganDeposito.frx":096B
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10110
         TabIndex        =   20
         Top             =   120
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
         Picture         =   "MstGolonganDeposito.frx":0B16
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   105
         TabIndex        =   21
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
         Picture         =   "MstGolonganDeposito.frx":0BBC
      End
   End
End
Attribute VB_Name = "MstGolonganDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim lEdit As Boolean
Dim nPos As SisPos

Private Sub GetMemory()
Dim cField As String
Dim vaJoin

  cField = " g.*,"
  cField = cField & " r.Keterangan as KeteranganAkuntansi,a.Keterangan as KeteranganBunga,"
  cField = cField & " b.Keterangan as KeteranganPajakBunga,d.Keterangan as KeteranganCadanganBunga,"
  cField = cField & " e.Keterangan as NamaRekeningJatuhTempo,"
  cField = cField & " f.Keterangan as NamaRekeningFinalty,"
  cField = cField & " m.Keterangan as NamaRekeningMaterai"
  vaJoin = Array("Left join rekening r on r.Kode = g.rekeningakuntansi", _
                 "Left join rekening a on a.Kode = g.rekeningBunga", _
                 "Left join rekening b on b.Kode = g.rekeningPajakBunga", _
                 "Left Join Rekening d on d.Kode = g.CadanganBunga", _
                 "Left Join Rekening e on e.Kode = g.RekeningJatuhTempo", _
                 "left Join rekening f on f.Kode  = g.rekeningFinalty", _
                 "left Join rekening m on m.Kode  = g.rekeningMaterai")
  Set dbData = objData.Browse(GetDSN, "GolonganDeposito g", cField, "g.Kode", sisAssign, "D" & cKode.Text, , "g.Kode", vaJoin)
  If Not dbData.eof Then
    cKeterangan.Text = GetNull(dbData!Keterangan, "")
    nLama.Value = GetNull(dbData!Lama)
    nBunga.Value = GetNull(dbData!bunga)
    cRekeningakuntansi.Text = GetNull(dbData!RekeningAkuntansi, "")
    cNamaRekAkuntansi.Text = GetNull(dbData!Keteranganakuntansi, "")
    cRekeningBahas.Text = GetNull(dbData!Rekeningbunga, "")
    cNamaRekBahas.Text = GetNull(dbData!KeteranganBunga, "")
    cRekeningPajakBahas.Text = GetNull(dbData!RekeningPajakbunga, "")
    cNamaRekPajakBahas.Text = GetNull(dbData!KeteranganPajakBunga, "")
    cCadanganBahas.Text = GetNull(dbData!Cadanganbunga, "")
    cNamaCadanganBahas.Text = GetNull(dbData!KeteranganCadanganBunga, "")
    cRekeningJatuhTempo.Text = GetNull(dbData!RekeningJatuhtempo, "")
    cNamaRekeningJatuhTempo.Text = GetNull(dbData!NamaRekeningJatuhTempo, "")
    cRekFinalti.Text = GetNull(dbData!RekeningFinalty, "")
    cNamaRekfinalti.Text = GetNull(dbData!NamaRekeningFinalty, "")
    cRekMaterai.Text = GetNull(dbData!rekeningmaterai, "")
    cNamaRekMaterai.Text = GetNull(dbData!NamaRekeningMaterai, "")
    nMinimum.Value = GetNull(dbData!MinimumkenaPajak)
    nPajak.Value = GetNull(dbData!pajakbunga)
  End If
End Sub

Private Sub cCadanganBahas_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "rekening", "Kode", cCadanganBahas, "Kode,Keterangan,Jenis", " and Jenis = 'D'")
  If Not dbData.eof Then
    cNamaCadanganBahas.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cCadanganBahas_Validate(Cancel As Boolean)
  If cCadanganBahas.LastKey = 13 Then
    cCadanganBahas_ButtonClick
  End If
End Sub

Private Sub cRekfinalti_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "rekening", "Kode", cRekFinalti, "Kode,Keterangan,Jenis", " and Jenis = 'D'")
  If Not dbData.eof Then
    cNamaRekfinalti.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cRekFinalti_Validate(Cancel As Boolean)
  If cRekFinalti.LastKey = 13 Then
    cRekfinalti_ButtonClick
  End If
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
  If cKode.LastKey = 13 Then
    If Not dbData.eof Then dbData.MoveFirst
    dbData.Find "Kode = 'D" & cKode.Text & "'"
    If Not dbData.eof Then
      If nPos = Add Then
        MsgBox "Kode Sudah Ada, Ulangi Pengisian", vbExclamation
        Cancel = True
        cKode.SetFocus
        Exit Sub
      End If
      GetMemory
      If nPos = Delete Then DeleteData
    ElseIf dbData.eof And nPos <> Add Then
      MsgBox "Data Tidak Ada, Ulangi Pengisian", vbExclamation + vbOKOnly
      Cancel = True
      cKode.SetFocus
      Exit Sub
    End If
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  cKode.SetFocus
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
  cKode.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  cKode.SetFocus
End Sub

Private Sub DeleteData()
  If MsgBox("Data Benar-benar Dihapus ?", vbYesNo + vbExclamation) = vbYes Then
     objData.Delete GetDSN, "GolonganDeposito", "Kode", sisAssign, "D" & cKode.Text
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
Dim vaArray As New XArrayDB
Dim cJudul As String

Set dbData = objData.Browse(GetDSN, "GolonganDeposito", "Kode,Keterangan,Lama,Bunga,RekeningAkuntansi,RekeningBunga,RekeningPajakBunga,CadanganBunga,RekeningFinalty", , , , , "Kode")
  If Not dbData.eof Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  End If
  With FrmRPT
    .AddPageHeader "DAFTAR GOLONGAN DEPOSITO", tdbHalignCenter, , , True, dbArial, 12, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "KODE", , , , 6, , , , , , , , , , , , , 5
    .AddTableHeader "KETERANGAN"
    .AddTableHeader "LAMA", , , , 6
    .AddTableHeader "BUNGA(%)", , , , 10
    .AddTableHeader "REKENING", , , , 10
    .AddTableHeader "REK BUNGA", , , , 10
    .AddTableHeader "REK PAJAK BUNGA", , , , 10
    .AddTableHeader "REK TITIPAN BUNGA", , , , 10
    .AddTableHeader "REK FINALTY", , , , 10
    
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody , tdbHalignRight
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    
    .Preview vaArray, True
  End With
End Sub

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue

  If ValidSaving() Then
    If MsgBox("Data Benar-benar sudah VALID ?", vbYesNo + vbInformation) = vbYes Then
      vaField = Array("Kode", "Keterangan", "Lama", "Bunga", "Rekeningakuntansi", "RekeningBunga", "RekeningPajakBunga", "CadanganBunga", "RekeningJatuhTempo", "RekeningFinalty", "MinimumKenaPajak", "PajakBunga", "RekeningMaterai")
      vaValue = Array("D" & cKode.Text, cKeterangan.Text, nLama.Value, nBunga.Value, cRekeningakuntansi.Text, cRekeningBahas.Text, cRekeningPajakBahas.Text, cCadanganBahas.Text, cRekeningJatuhTempo.Text, cRekFinalti.Text, nMinimum.Value, nPajak.Value, cRekMaterai.Text)
      objData.Update GetDSN, "golongandeposito", "kode = 'D" & cKode.Text & "'", vaField, vaValue
      GetSQL
      initvalue
      GetEdit False
    Else
      cKode.SetFocus
      Exit Sub
    End If
  End If
End Sub

Static Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cKode.Text, "Kode Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
 If Not CheckData(cKeterangan.Text, "Keterangan Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cRekeningakuntansi.Text, "Rekening Akuntansi Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cRekeningakuntansi.SetFocus
    Exit Function
  End If
  
  If Not CheckData(nLama.Value, "Lama Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    nLama.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cRekeningBahas.Text, "Rekening Bunga Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cRekeningBahas.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cRekeningPajakBahas.Text, "Rekening Pajak Bunga Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cRekeningPajakBahas.SetFocus
    Exit Function
  End If
  
  If Not CheckData(nBunga.Value, "Suku Bunga Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    nBunga.SetFocus
    Exit Function
  End If
  
  If Not CheckData(nMinimum.Value, "Minimum Kena Pajak Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    nMinimum.SetFocus
    Exit Function
  End If
  
  If Not CheckData(nPajak.Value, "Pajak bunga Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    nPajak.SetFocus
    Exit Function
  End If
End Function

Private Sub cRekeningAkuntansi_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "Rekening", "Kode", cRekeningakuntansi, "Kode,Keterangan,Jenis", " and jenis = 'D'")
  If Not dbData.eof Then
    cRekeningakuntansi.Text = GetNull(dbData!Kode, "")
    cNamaRekAkuntansi.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cRekeningAkuntansi_Validate(Cancel As Boolean)
 If cRekeningakuntansi.LastKey = 13 Then
    cRekeningAkuntansi_ButtonClick
 End If
End Sub

Private Sub cRekeningBahas_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "rekening", "Kode", cRekeningBahas, "Kode,Keterangan,Jenis", " and jenis = 'D'")
  If Not dbData.eof Then
    cRekeningBahas.Text = GetNull(dbData!Kode, "")
    cNamaRekBahas.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cRekeningBahas_Validate(Cancel As Boolean)
  If cRekeningBahas.LastKey = 13 Then
    cRekeningBahas_ButtonClick
  End If
End Sub

Private Sub cRekeningJatuhTempo_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "rekening", "Kode", cRekeningJatuhTempo, "Kode,Keterangan,Jenis", " and Jenis = 'D'")
  If Not dbData.eof Then
    cNamaRekeningJatuhTempo.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cRekeningJatuhTempo_Validate(Cancel As Boolean)
  If cRekeningJatuhTempo.LastKey = 13 Then
    cRekeningJatuhTempo_ButtonClick
  End If
End Sub

Private Sub cRekeningPajakBahas_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "rekening", "Kode", cRekeningPajakBahas, "Kode,Keterangan,Jenis", " and jenis = 'D'")
  If Not dbData.eof Then
    cRekeningPajakBahas.Text = GetNull(dbData!Kode, "")
    cNamaRekPajakBahas.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cRekeningPajakBahas_Validate(Cancel As Boolean)
  If cRekeningPajakBahas.LastKey = 13 Then
    cRekeningPajakBahas_ButtonClick
  End If
End Sub

Private Sub cRekMaterai_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "rekening", "Kode", cRekMaterai, "Kode,Keterangan,Jenis", " and Jenis = 'D'")
  If Not dbData.eof Then
    cNamaRekMaterai.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cRekMaterai_Validate(Cancel As Boolean)
  If cRekMaterai.LastKey = 13 Or cRekMaterai.LastKey = 40 Then
    cRekMaterai_ButtonClick
  End If
End Sub

Private Sub TDBGrid1_Click()
  lClick = True
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Then
    lClick = True
  End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If lClick And Not lEdit Then
    cKode.Default
    cKode.Text = Right(TDBGrid1.Columns(0), 1)
    GetMemory
  End If
  lClick = False
End Sub

Private Sub initvalue()
  cKode.Default
  cKeterangan.Default
  cRekeningakuntansi.Default
  cNamaRekAkuntansi.Default
  cRekeningBahas.Default
  cNamaRekBahas.Default
  cRekeningPajakBahas.Default
  cNamaRekPajakBahas.Default
  cRekeningJatuhTempo.Default
  cNamaRekeningJatuhTempo.Default
  cRekMaterai.Default
  cNamaRekMaterai.Default
  nLama.Value = 0
  nBunga.Value = 0
  cCadanganBahas.Default
  cNamaCadanganBahas.Default
  cRekFinalti.Default
  cNamaRekfinalti.Default
  nMinimum.Value = 0
  nPajak.Value = 0
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  Me.Top = 0
  GetSQL
  initvalue
  GetEdit False
  
  TabIndex cKode, n
  TabIndex cKeterangan, n
  TabIndex cRekeningakuntansi, n
  TabIndex cRekeningBahas, n
  TabIndex cRekeningPajakBahas, n
  TabIndex cRekeningJatuhTempo, n
  TabIndex cCadanganBahas, n
  TabIndex cRekFinalti, n
  TabIndex cRekMaterai, n
  TabIndex nLama, n
  TabIndex nBunga, n
  TabIndex nMinimum, n
  TabIndex nPajak, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub GetSQL()
  Set dbData = objData.Browse(GetDSN, "golongandeposito", , , , , , "Kode")
  Set TDBGrid1.DataSource = dbData
End Sub
