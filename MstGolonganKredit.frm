VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form MstGolonganKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER GOLONGAN PINJAMAN"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   11775
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3285
      Left            =   0
      Top             =   2835
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5794
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
         Height          =   3120
         Left            =   60
         TabIndex        =   0
         Top             =   75
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5503
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
         Columns(2).Caption=   "REKE. AKUNTANSI"
         Columns(2).DataField=   "Rekening"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "REK DENDA"
         Columns(3).DataField=   "RekeningDenda"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "REK ADMINISTRSI"
         Columns(4).DataField=   "RekeningAdministrasi"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "REK. PROVISI"
         Columns(5).DataField=   "REKENINGPROVISI"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "REK MATERAI"
         Columns(6).DataField=   "RekeningMaterai"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "REK NOTARIS"
         Columns(7).DataField=   "REKENINGNOTARIS"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "REK BIAYA LAIN"
         Columns(8).DataField=   "RekeningBiayalainLain"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "REK ANGS POKOK"
         Columns(9).DataField=   "RekeningAngsuranPokok"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "REK ANGS BUNGA"
         Columns(10).DataField=   "RekeningAngsuranBunga"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   3
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1032"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=6059"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5980"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3122"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3043"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3069"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2990"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(51)=   "Column(10).Width=3149"
         Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=3069"
         Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=516"
         Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFCFCED&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H0&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
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
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(81)  =   "Named:id=33:Normal"
         _StyleDefs(82)  =   ":id=33,.parent=0"
         _StyleDefs(83)  =   "Named:id=34:Heading"
         _StyleDefs(84)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(85)  =   ":id=34,.wraptext=-1"
         _StyleDefs(86)  =   "Named:id=35:Footing"
         _StyleDefs(87)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(88)  =   "Named:id=36:Selected"
         _StyleDefs(89)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(90)  =   "Named:id=37:Caption"
         _StyleDefs(91)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(92)  =   "Named:id=38:HighlightRow"
         _StyleDefs(93)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(94)  =   "Named:id=39:EvenRow"
         _StyleDefs(95)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(96)  =   "Named:id=40:OddRow"
         _StyleDefs(97)  =   ":id=40,.parent=33"
         _StyleDefs(98)  =   "Named:id=41:RecordSelector"
         _StyleDefs(99)  =   ":id=41,.parent=34"
         _StyleDefs(100) =   "Named:id=42:FilterBar"
         _StyleDefs(101) =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2865
      Left            =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5054
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningDenda 
         Height          =   330
         Left            =   3090
         TabIndex        =   1
         Top             =   1215
         Width           =   3060
         _ExtentX        =   5398
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningProvisi 
         Height          =   330
         Left            =   3090
         TabIndex        =   2
         Top             =   1935
         Width           =   3060
         _ExtentX        =   5398
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekening 
         Height          =   330
         Left            =   3090
         TabIndex        =   3
         Top             =   855
         Width           =   3060
         _ExtentX        =   5398
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   855
         Width           =   2985
         _ExtentX        =   5265
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
         Caption         =   "R. Akuntansi"
         CaptionWidth    =   1300
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
         Left            =   90
         TabIndex        =   5
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
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
         Caption         =   "KODE          K"
         CaptionWidth    =   1300
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
         Left            =   90
         TabIndex        =   6
         Top             =   495
         Width           =   6060
         _ExtentX        =   10689
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
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningProvisi 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   1935
         Width           =   2985
         _ExtentX        =   5265
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
         Caption         =   "R. Provisi"
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningDenda 
         Height          =   330
         Left            =   90
         TabIndex        =   8
         Top             =   1215
         Width           =   2985
         _ExtentX        =   5265
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
         Caption         =   "R. Denda Ang"
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningAdministrasi 
         Height          =   330
         Left            =   3090
         TabIndex        =   9
         Top             =   1575
         Width           =   3060
         _ExtentX        =   5398
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningAdministrasi 
         Height          =   330
         Left            =   90
         TabIndex        =   10
         Top             =   1575
         Width           =   2985
         _ExtentX        =   5265
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
         Caption         =   "R. Admin."
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningMaterai 
         Height          =   330
         Left            =   9030
         TabIndex        =   11
         Top             =   510
         Width           =   2640
         _ExtentX        =   4657
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningMaterai 
         Height          =   330
         Left            =   6195
         TabIndex        =   12
         Top             =   510
         Width           =   2820
         _ExtentX        =   4974
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
         Caption         =   "R. Materai"
         CaptionWidth    =   1100
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningAngsuranPokok 
         Height          =   330
         Left            =   9030
         TabIndex        =   13
         Top             =   1605
         Width           =   2640
         _ExtentX        =   4657
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningAngsuranPokok 
         Height          =   330
         Left            =   6180
         TabIndex        =   14
         Top             =   1605
         Width           =   2820
         _ExtentX        =   4974
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
         Caption         =   "R. Ang Pk"
         CaptionWidth    =   1100
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningAngsuranBunga 
         Height          =   330
         Left            =   9030
         TabIndex        =   15
         Top             =   1965
         Width           =   2640
         _ExtentX        =   4657
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningAngsuranBunga 
         Height          =   330
         Left            =   6180
         TabIndex        =   16
         Top             =   1965
         Width           =   2820
         _ExtentX        =   4974
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
         Caption         =   "R. Ang Bng"
         CaptionWidth    =   1100
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningNotaris 
         Height          =   330
         Left            =   9030
         TabIndex        =   24
         Top             =   870
         Width           =   2640
         _ExtentX        =   4657
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningNotaris 
         Height          =   330
         Left            =   6195
         TabIndex        =   25
         Top             =   870
         Width           =   2820
         _ExtentX        =   4974
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
         Caption         =   "R. Notaris"
         CaptionWidth    =   1100
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningBiayaLain 
         Height          =   330
         Left            =   9030
         TabIndex        =   26
         Top             =   1230
         Width           =   2640
         _ExtentX        =   4657
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningBiayalain 
         Height          =   330
         Left            =   6195
         TabIndex        =   27
         Top             =   1230
         Width           =   2820
         _ExtentX        =   4974
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
         Caption         =   "R. By Lain"
         CaptionWidth    =   1100
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningSimpananWajib 
         Height          =   330
         Left            =   9030
         TabIndex        =   28
         Top             =   2325
         Width           =   2640
         _ExtentX        =   4657
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningSimpananWajib 
         Height          =   330
         Left            =   6180
         TabIndex        =   29
         Top             =   2325
         Width           =   2820
         _ExtentX        =   4974
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
         Caption         =   "R. Sim Wjb"
         CaptionWidth    =   1100
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
      Top             =   6090
      Width           =   11775
      _ExtentX        =   20770
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
         TabIndex        =   17
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
         Picture         =   "MstGolonganKredit.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3825
         TabIndex        =   18
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
         Picture         =   "MstGolonganKredit.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9480
         TabIndex        =   19
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
         Picture         =   "MstGolonganKredit.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1605
         TabIndex        =   20
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
         Picture         =   "MstGolonganKredit.frx":083F
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   540
         TabIndex        =   21
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
         Picture         =   "MstGolonganKredit.frx":096B
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10560
         TabIndex        =   22
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
         Picture         =   "MstGolonganKredit.frx":0B16
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   105
         TabIndex        =   23
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
         Picture         =   "MstGolonganKredit.frx":0BBC
      End
   End
End
Attribute VB_Name = "MstGolonganKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim dbRekening As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim lEdit As Boolean
Dim nPos As SisPos
Dim cSQL As String

Private Sub cKode_Validate(Cancel As Boolean)
    If cKode.LastKey = 13 Then
      If Not dbData.eof Then dbData.MoveFirst
      dbData.Find "Kode = 'K" & cKode.Text & "'"
      If Not dbData.eof Then
        If nPos = Add Then
          MsgBox "Kode Sudah Ada, Ulangi Pengisian", vbExclamation
          Cancel = True
          cKode.Default
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
  If MsgBox("Data BenarDihapus ?", vbYesNo + vbExclamation) = vbYes Then
    objData.Delete GetDSN, "GolonganKredit", "kode", sisAssign, "K" & cKode.Text
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
Dim cField As String
Dim vaArray As New XArrayDB

  cField = "Kode,Keterangan,Rekening,RekeningDenda,RekeningAdministrasi,RekeningProvisi,"
  cField = cField & " RekeningMaterai,RekeningNotaris,RekeningBiayalainLain, RekeningAngsuranPokok,RekeningAngsuranBunga"
  Set dbData = objData.Browse(GetDSN, "GolonganKredit", cField, , , , , "Kode")
  If Not dbData.eof Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  End If
  
  With FrmRPT
    .AddPageHeader "DAFTAR GOLONGAN KREDIT", tdbHalignCenter, , , True, dbArial, 12, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "KODE", , , , 6, , , , , , , , , , , , , 5
    .AddTableHeader "KETERANGAN"
    .AddTableHeader "REKENING", , , , 7
    .AddTableHeader "REK DENDA", , , , 7
    .AddTableHeader "REK ADMINS.", , , , 8
    .AddTableHeader "REK PROVISI", , , , 8
    .AddTableHeader "REK MATERAI", , , , 8
    .AddTableHeader "REK NOTARIS", , , , 8
    .AddTableHeader "REK BIAYA LAIN", , , , 9
    .AddTableHeader "REK ANG POKOK", , , , 9
    .AddTableHeader "REK ANG BUNGA", , , , 9
    

    .AddTableBody , tdbHalignCenter
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    
    .Preview vaArray, True, , True
  End With
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue

  If ValidSaving() Then
      If MsgBox("Data benar-benar sudah VALID ?", vbYesNo + vbInformation) = vbYes Then
        vaField = Array("kode", "keterangan", "Rekening", _
                        "RekeningAdministrasi", "RekeningDenda", _
                        "RekeningMaterai", "RekeningProvisi", "RekeningNotaris", _
                        "RekeningAngsuranPokok", "RekeningAngsuranBunga", "RekeningBiayalainLain", "rekeningsimpananwajib")
        vaValue = Array("K" & cKode.Text, cKeterangan.Text, cRekening.Text, _
                        cRekeningAdministrasi.Text, cRekeningDenda.Text, _
                        cRekeningMaterai.Text, cRekeningProvisi.Text, cRekeningNotaris.Text, _
                        cRekeningAngsuranPokok.Text, cRekeningAngsuranBunga.Text, cRekeningBiayalain.Text, cRekeningSimpananWajib.Text)
        objData.Update GetDSN, "GolonganKredit", "Kode = 'K" & cKode.Text & "'", vaField, vaValue
        GetSQL
        initvalue
        GetEdit False
      End If
  End If
End Sub

Static Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cKode.Text, "Kode GolonganKredit Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKeterangan.Text, "Nama GolonganKredit Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cRekening.Text, "Rekening Akuntansi Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cRekening.SetFocus
    Exit Function
  End If
End Function

Private Sub Pick(cRek, cNM)
  cNM.Text = ""
  Set dbRekening = objData.Pick(GetDSN, "Rekening", "Kode", cRek, "Kode,Keterangan,Jenis", " and jenis = 'D'")
  If dbRekening.RecordCount > 0 Then
    cNM.Text = GetNull(dbRekening!Keterangan, "")
  End If
End Sub

Private Sub cRekening_ButtonClick()
  Pick cRekening, cNamaRekening
End Sub

Private Sub cRekening_Validate(Cancel As Boolean)
  If cRekening.LastKey = 13 Then
    cRekening_ButtonClick
  End If
End Sub

Private Sub cRekeningAngsuranBunga_Validate(Cancel As Boolean)
  If cRekeningAngsuranBunga.LastKey = 13 Then
    cRekeningAngsuranBunga_ButtonClick
  End If
End Sub

Private Sub cRekeningAngsuranPokok_Validate(Cancel As Boolean)
  If cRekeningAngsuranPokok.LastKey = 13 Then
    cRekeningAngsuranPokok_ButtonClick
  End If
End Sub

Private Sub cRekeningBiayalain_Buttonclick()
  Pick cRekeningBiayalain, cNamaRekeningBiayaLain
End Sub

Private Sub cRekeningBiayalain_Validate(Cancel As Boolean)
  If cRekeningBiayalain.LastKey = 13 Then
    cRekeningBiayalain_Buttonclick
  End If
End Sub

Private Sub cRekeningProvisi_ButtonClick()
  Pick cRekeningProvisi, cNamaRekeningProvisi
End Sub

Private Sub cRekeningProvisi_Validate(Cancel As Boolean)
  If cRekeningProvisi.LastKey = 13 Then
    cRekeningProvisi_ButtonClick
  End If
End Sub

Private Sub cRekeningDenda_ButtonClick()
  Pick cRekeningDenda, cNamaRekeningDenda
End Sub

Private Sub cRekeningDenda_Validate(Cancel As Boolean)
  If cRekeningDenda.LastKey = 13 Then
    cRekeningDenda_ButtonClick
  End If
End Sub

Private Sub cRekeningAngsuranPokok_ButtonClick()
  Pick cRekeningAngsuranPokok, cNamaRekeningAngsuranPokok
End Sub

Private Sub cRekeningAngsuranBunga_ButtonClick()
  Pick cRekeningAngsuranBunga, cNamaRekeningAngsuranBunga
End Sub

Private Sub cRekeningSimpananWajib_ButtonClick()
  Pick cRekeningSimpananWajib, cNamaRekeningSimpananWajib
End Sub

Private Sub cRekeningSimpananWajib_Validate(Cancel As Boolean)
  If cRekeningSimpananWajib.LastKey = 13 Then
    cRekeningSimpananWajib_ButtonClick
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
    cKode.Text = ""
    cKode.Text = Right(TDBGrid1.Columns(0), 1)
    GetMemory
  End If
  lClick = False
End Sub

Private Sub GetMemory()
Dim cFields As String
Dim vaJoin

    cFields = "g.*, r.Keterangan as NamaRekening,"
    cFields = cFields & " m.Keterangan as NamaRekeningProvisi,"
    cFields = cFields & " n.Keterangan as NamaRekeningDenda,"
    cFields = cFields & " a.Keterangan as NamaRekeningAdministrasi,"
    cFields = cFields & " e.Keterangan as NamaRekeningMaterai,"
    cFields = cFields & " t.Keterangan as NamaRekeningNotaris,"
    cFields = cFields & " h.Keterangan as NamaRekeningAngsuranPokok,"
    cFields = cFields & " i.Keterangan as NamaRekeningAngsuranBunga,"
    cFields = cFields & " b.Keterangan as NamaRekeningBiayaLain,"
    cFields = cFields & " z.Keterangan as namarekeningsimpananwajib"
    
    vaJoin = Array(" Left Join Rekening r On r.Kode = g.Rekening", _
                 " Left Join Rekening m On m.Kode = g.RekeningProvisi", _
                 " Left Join Rekening n On n.Kode = g.RekeningDenda", _
                 " Left Join Rekening a On a.Kode = g.RekeningAdministrasi", _
                 " Left Join Rekening e On e.Kode = g.RekeningMaterai", _
                 " Left Join Rekening h On h.Kode = g.RekeningAngsuranPokok", _
                 " Left Join Rekening i On i.Kode = g.RekeningAngsuranBunga", _
                 " Left Join Rekening t on t.Kode = g.RekeningNotaris", _
                 " Left Join rekening b on b.Kode = g.RekeningBiayalainLain", _
                 " LEFT JOIN rekening z on z.kode = g.rekeningsimpananwajib")
                 
    Set dbData = objData.Browse(GetDSN, "GolonganKredit g", cFields, "g.Kode", sisAssign, "K" & cKode.Text, , , vaJoin)
    If Not dbData.eof Then
      With dbData
        cKeterangan.Text = GetNull(!Keterangan, "")
        cRekening.Text = GetNull(!Rekening, "")
        cRekeningDenda.Text = GetNull(!rekeningdenda, "")
        cRekeningAdministrasi.Text = GetNull(!rekeningadministrasi, "")
        cRekeningProvisi.Text = GetNull(!rekeningprovisi, "")
        cRekeningMaterai.Text = GetNull(!rekeningmaterai, "")
        cRekeningNotaris.Text = GetNull(!RekeningNotaris, "")
        cRekeningBiayalain.Text = GetNull(dbData!RekeningBiayalainLain)
        cRekeningAngsuranPokok.Text = GetNull(!RekeningAngsuranPokok, "")
        cRekeningAngsuranBunga.Text = GetNull(!rekeningangsuranbunga, "")
        cRekeningSimpananWajib.Text = GetNull(!rekeningsimpananwajib, "")
        
        cNamaRekening.Text = GetNull(!NamaRekening, "")
        cNamaRekeningDenda.Text = GetNull(!NamaRekeningDenda, "")
        cNamaRekeningAdministrasi.Text = GetNull(!NamaRekeningAdministrasi, "")
        cNamaRekeningProvisi.Text = GetNull(!NamaRekeningProvisi, "")
        cNamaRekeningMaterai.Text = GetNull(!NamaRekeningMaterai, "")
        cNamaRekeningNotaris.Text = GetNull(!NamaRekeningNotaris, "")
        cNamaRekeningBiayaLain.Text = GetNull(dbData!NamaRekeningBiayaLain)
        cNamaRekeningAngsuranPokok.Text = GetNull(!NamaRekeningAngsuranPokok, "")
        cNamaRekeningAngsuranBunga.Text = GetNull(!NamaRekeningAngsuranBunga, "")
        cNamaRekeningSimpananWajib.Text = GetNull(!NamaRekeningSimpananWajib, "")
      End With
    End If
End Sub

Private Sub initvalue()
  cKode.Default
  cKeterangan.Default
  cRekening.Default
  cNamaRekening.Default
  cRekeningDenda.Default
  cNamaRekeningDenda.Default
  cRekeningAdministrasi.Default
  cNamaRekeningAdministrasi.Default
  cRekeningProvisi.Default
  cNamaRekeningProvisi.Default
  cRekeningMaterai.Default
  cNamaRekeningMaterai.Default
  cRekeningNotaris.Default
  cNamaRekeningNotaris.Default
  cRekeningBiayalain.Default
  cNamaRekeningBiayaLain.Default
  cRekeningAngsuranPokok.Default
  cRekeningAngsuranBunga.Default
  cNamaRekeningAngsuranPokok.Default
  cNamaRekeningAngsuranBunga.Default
  cRekeningSimpananWajib.Default
  cNamaRekeningSimpananWajib.Default
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me, True
  Me.Top = 0
  Me.left = 0
  GetSQL
  initvalue
  GetEdit False
  
  TabIndex cKode, n
  TabIndex cKeterangan, n
  TabIndex cRekening, n
  TabIndex cRekeningDenda, n
  TabIndex cRekeningAdministrasi, n
  TabIndex cRekeningProvisi, n
  TabIndex cRekeningMaterai, n
  TabIndex cRekeningNotaris, n
  TabIndex cRekeningBiayalain, n
  TabIndex cRekeningAngsuranPokok, n
  TabIndex cRekeningAngsuranBunga, n
  TabIndex cRekeningSimpananWajib, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdPreview, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub GetSQL()
  Set dbData = objData.Browse(GetDSN, "GolonganKredit", , , , , , "Kode")
  Set TDBGrid1.DataSource = dbData
End Sub

Private Sub cRekeningAdministrasi_ButtonClick()
  Pick cRekeningAdministrasi, cNamaRekeningAdministrasi
End Sub

Private Sub cRekeningAdministrasi_Validate(Cancel As Boolean)
  If cRekeningAdministrasi.LastKey = 13 Then
    cRekeningAdministrasi_ButtonClick
  End If
End Sub

Private Sub cRekeningMaterai_ButtonClick()
  Pick cRekeningMaterai, cNamaRekeningMaterai
End Sub

Private Sub cRekeningMaterai_Validate(Cancel As Boolean)
  If cRekeningMaterai.LastKey = 13 Then
    cRekeningMaterai_ButtonClick
  End If
End Sub

Private Sub cRekeningNotaris_ButtonClick()
  Pick cRekeningNotaris, cNamaRekeningNotaris
End Sub

Private Sub cRekeningNotaris_Validate(Cancel As Boolean)
  If cRekeningNotaris.LastKey = 13 Then
    cRekeningNotaris_ButtonClick
  End If
End Sub

