VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form RptBukuAngsuran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN BUKU ANGSURAN"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   11790
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   4995
      Left            =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8811
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   300
         Left            =   3735
         TabIndex        =   0
         Top             =   90
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
         Left            =   2175
         TabIndex        =   1
         Top             =   90
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
         Left            =   135
         TabIndex        =   2
         Top             =   90
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
         Caption         =   "NO. REKENING"
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
         Left            =   2925
         TabIndex        =   3
         Top             =   90
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
         Left            =   135
         TabIndex        =   4
         Top             =   420
         Width           =   4710
         _ExtentX        =   8308
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Caption         =   "NAMA"
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
         Left            =   135
         TabIndex        =   5
         Top             =   750
         Width           =   6630
         _ExtentX        =   11695
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Caption         =   "ALAMAT"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3780
         Left            =   90
         TabIndex        =   8
         Top             =   1155
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   6668
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "FAKTUR"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "TANGGAL"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NO SPK"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "PLAFOND"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "FormatText Event"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TLG CAIR"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "ANGS POKOK"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "FormatText Event"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "ANGS BUNGA"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "DENDA"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "FormatText Event"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TOTAL"
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "FormatText Event"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "BAKI DEBET"
         Columns(9).DataField=   ""
         Columns(9).NumberFormat=   "FormatText Event"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4233"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4154"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2990"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2910"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2778"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2699"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2858"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2778"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2672"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2593"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=2672"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2593"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=2831"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2752"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=514"
         Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=172,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
         _StyleDefs(78)  =   "Named:id=33:Normal"
         _StyleDefs(79)  =   ":id=33,.parent=0"
         _StyleDefs(80)  =   "Named:id=34:Heading"
         _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(82)  =   ":id=34,.wraptext=-1"
         _StyleDefs(83)  =   "Named:id=35:Footing"
         _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(85)  =   "Named:id=36:Selected"
         _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(87)  =   "Named:id=37:Caption"
         _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(89)  =   "Named:id=38:HighlightRow"
         _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(91)  =   "Named:id=39:EvenRow"
         _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(93)  =   "Named:id=40:OddRow"
         _StyleDefs(94)  =   ":id=40,.parent=33"
         _StyleDefs(95)  =   "Named:id=41:RecordSelector"
         _StyleDefs(96)  =   ":id=41,.parent=34"
         _StyleDefs(97)  =   "Named:id=42:FilterBar"
         _StyleDefs(98)  =   ":id=42,.parent=33"
      End
      Begin TrueDBReports60Ctl.TDBReports RptFakturPenjualan 
         Height          =   570
         Left            =   8130
         TabIndex        =   11
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1005
         Caption         =   "FakturPenjualan"
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ErrorMsgCaption =   ""
         Filtered        =   0   'False
         DataMode        =   1
         DataMember      =   ""
         LinkSequence    =   1
         LinkOrder       =   0
         NameSubstitute  =   ""
         ConnectionString=   "DSN=MySalemba"
         ConnectStringType=   3
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "MySalemba"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         CursorLocation  =   3
         ConnectionTimeout=   15
         CommandTimeout  =   30
         RecordSource    =   ""
         CursorType      =   1
         CommandType     =   8
         MaxRecords      =   0
         LinkType        =   0
         Master          =   ""
         CallDataRead    =   0   'False
         ConvertNullToEmpty=   -1  'True
         DesignConnection=   -1  'True
         DesignTimeout   =   5
         UnitsOfMeasurement=   4
         Vedit_ShowGrid  =   -1  'True
         Vedit_SnapToGrid=   0   'False
         Vedit_GridUnitWidth=   2.822
         Vedit_GridUnitHeight=   2.822
         Vedit_ShowCellExpressions=   -1  'True
         Norm_rect_left  =   0
         Norm_rect_top   =   0
         Norm_rect_right =   0
         Norm_rect_bottom=   0
         Virgin          =   0   'False
         Parameters.Count=   11
         Parameters(0).Name=   "cNamaDebitur"
         Parameters(1).Name=   "cNoRekening"
         Parameters(2).Name=   "cJaminan"
         Parameters(3).Name=   "cJangkaWaktu"
         Parameters(4).Name=   "cSukuBunga"
         Parameters(5).Name=   "nKewajibanPokok"
         Parameters(6).Name=   "nKewajibanBunga"
         Parameters(7).Name=   "nJumlahKewajiban"
         Parameters(8).Name=   "dTglRealisasi"
         Parameters(9).Name=   "dTglJatuhTempo"
         Parameters(10).Name=   "nPlafond"
         Fields.Count    =   6
         Fields(0).Name  =   "dTgl"
         Fields(0).DisplayName=   "NamaBarang"
         Fields(0).Type  =   5
         Fields(1).Name  =   "nPokok"
         Fields(1).DisplayName=   "Qty"
         Fields(1).MaxLength=   40
         Fields(2).Name  =   "nBunga"
         Fields(2).DisplayName=   "Harga"
         Fields(2).MaxLength=   40
         Fields(3).Name  =   "nJumlah"
         Fields(3).DisplayName=   "Satuan"
         Fields(3).MaxLength=   3
         Fields(4).Name  =   "cParaf"
         Fields(4).DisplayName=   "Disc"
         Fields(5).Name  =   "nDenda"
         Fields(5).DisplayName=   "nDenda"
         Sections.Count  =   3
         Sections(0).Name=   "SECTION_1"
         Sections(0).Type=   3
         Sections(0).Condition=   "IsBegOfTable()"
         Sections(0).Cells.Count=   15
         Sections(0).Cells(0).Name=   "CELL_20"
         Sections(0).Cells(0).Exp=   """** KSP MITRA ABADI - KARTU ANGSURAN PINJAMAN**"""
         Sections(0).Cells(0).PrivateStyle=   -1  'True
         Sections(0).Cells(0).Style.Name=   "<private>"
         Sections(0).Cells(0).Style.ParentName=   "<null>"
         Sections(0).Cells(0).Style.Font_Name=   "Times New Roman"
         Sections(0).Cells(0).Style.Font_Size=   10
         Sections(0).Cells(0).Style.Font_Bold=   0   'False
         Sections(0).Cells(0).Style.Font_Italic=   0   'False
         Sections(0).Cells(0).Style.Font_Underline=   -1  'True
         Sections(0).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(0).Style.Font_Charset=   1
         Sections(0).Cells(0).Style.TextAlign=   1
         Sections(0).Cells(0).Style.TextVAlign=   0
         Sections(0).Cells(0).Style.TextWrap=   -1  'True
         Sections(0).Cells(0).Style.ForeColor=   0
         Sections(0).Cells(0).Style.BackColor=   16777215
         Sections(0).Cells(0).Style.NoFill=   -1  'True
         Sections(0).Cells(0).Style.BackPicFile=   ""
         Sections(0).Cells(0).Style.ForePic=   "RptBukuAngsuran.frx":0000
         Sections(0).Cells(0).Style.ForePicFile=   ""
         Sections(0).Cells(0).Style.BackPicVertPlacement=   4
         Sections(0).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(0).Style.ForePicPlacement=   5
         Sections(0).Cells(0).Style.ForePicDrawMode=   0
         Sections(0).Cells(0).Style.MarginLeft=   6
         Sections(0).Cells(0).Style.MarginTop=   6
         Sections(0).Cells(0).Style.MarginRight=   6
         Sections(0).Cells(0).Style.MarginBottom=   6
         Sections(0).Cells(0).Style.HasBorders=   -1  'True
         Sections(0).Cells(0).Style.BorderHT=   ""
         Sections(0).Cells(0).Style.BorderHI=   ""
         Sections(0).Cells(0).Style.BorderHB=   ""
         Sections(0).Cells(0).Style.BorderVL=   ""
         Sections(0).Cells(0).Style.BorderVI=   ""
         Sections(0).Cells(0).Style.BorderVR=   ""
         Sections(0).Cells(0).Style.NoClipping=   0   'False
         Sections(0).Cells(0).Style.RTF=   0   'False
         Sections(0).Cells(0).Style.fprops=   603980545
         Sections(0).Cells(1).Name=   "CELL_3"
         Sections(0).Cells(1).Exp=   """Nama"""
         Sections(0).Cells(1).NewLine=   -1  'True
         Sections(0).Cells(1).Width=   25
         Sections(0).Cells(2).Name=   "CELL_6"
         Sections(0).Cells(2).Exp=   "cNamaDebitur"
         Sections(0).Cells(3).Name=   "CELL_14"
         Sections(0).Cells(4).Name=   "CELL_7"
         Sections(0).Cells(4).Exp=   """No. Rekening"""
         Sections(0).Cells(4).NewLine=   -1  'True
         Sections(0).Cells(4).Width=   25
         Sections(0).Cells(5).Name=   "CELL_8"
         Sections(0).Cells(5).Exp=   "cNoRekening"
         Sections(0).Cells(5).Width=   30
         Sections(0).Cells(6).Name=   "CELL_15"
         Sections(0).Cells(6).Exp=   """Pokok : "" & nKewajibanPokok"
         Sections(0).Cells(7).Name=   "CELL_9"
         Sections(0).Cells(7).Exp=   """Jaminan"""
         Sections(0).Cells(7).NewLine=   -1  'True
         Sections(0).Cells(7).Width=   25
         Sections(0).Cells(8).Name=   "CELL_10"
         Sections(0).Cells(8).Exp=   "cJaminan"
         Sections(0).Cells(8).Width=   30
         Sections(0).Cells(9).Name=   "CELL_16"
         Sections(0).Cells(9).Exp=   """Bunga : "" & nKewajibanBunga"
         Sections(0).Cells(10).Name=   "CELL_11"
         Sections(0).Cells(10).Exp=   """Jangka Waktu"""
         Sections(0).Cells(10).NewLine=   -1  'True
         Sections(0).Cells(10).Width=   25
         Sections(0).Cells(11).Name=   "CELL_12"
         Sections(0).Cells(11).Exp=   "cJangkaWaktu"
         Sections(0).Cells(11).Width=   30
         Sections(0).Cells(12).Name=   "CELL_18"
         Sections(0).Cells(12).Exp=   """Jumlah :"" & nJumlahKewajiban"
         Sections(0).Cells(13).Name=   "CELL_17"
         Sections(0).Cells(13).Exp=   """Plafond : "" & nPlafond & "" Bunga : "" & cSukuBunga & "" %/bln  Tgl Kredit : "" & dTglRealisasi & "" sd "" & dTglJatuhTempo"
         Sections(0).Cells(13).NewLine=   -1  'True
         Sections(0).Cells(14).Name=   "CELL_19"
         Sections(0).Cells(14).NewLine=   -1  'True
         Sections(1).Name=   "DetailHeader"
         Sections(1).Type=   3
         Sections(1).StyleExp=   "'Tdb_Header'"
         Sections(1).Tabulator=   "Detail"
         Sections(1).AutoHeight=   0   'False
         Sections(1).Height=   6
         Sections(1).Cells.Count=   6
         Sections(1).Cells(0).Name=   "CELL_0"
         Sections(1).Cells(0).Exp=   """Tanggal"""
         Sections(1).Cells(0).PrivateStyle=   -1  'True
         Sections(1).Cells(0).Style.Name=   "<private>"
         Sections(1).Cells(0).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(0).Style.Font_Name=   "Times New Roman"
         Sections(1).Cells(0).Style.Font_Size=   8.25
         Sections(1).Cells(0).Style.Font_Bold=   -1  'True
         Sections(1).Cells(0).Style.Font_Italic=   0   'False
         Sections(1).Cells(0).Style.Font_Underline=   0   'False
         Sections(1).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(0).Style.Font_Charset=   0
         Sections(1).Cells(0).Style.TextAlign=   1
         Sections(1).Cells(0).Style.TextVAlign=   1
         Sections(1).Cells(0).Style.TextWrap=   -1  'True
         Sections(1).Cells(0).Style.ForeColor=   0
         Sections(1).Cells(0).Style.BackColor=   16777215
         Sections(1).Cells(0).Style.NoFill=   -1  'True
         Sections(1).Cells(0).Style.BackPicFile=   ""
         Sections(1).Cells(0).Style.ForePicFile=   ""
         Sections(1).Cells(0).Style.BackPicVertPlacement=   0
         Sections(1).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(0).Style.ForePicPlacement=   0
         Sections(1).Cells(0).Style.ForePicDrawMode=   0
         Sections(1).Cells(0).Style.MarginLeft=   6
         Sections(1).Cells(0).Style.MarginTop=   1
         Sections(1).Cells(0).Style.MarginRight=   6
         Sections(1).Cells(0).Style.MarginBottom=   1
         Sections(1).Cells(0).Style.HasBorders=   -1  'True
         Sections(1).Cells(0).Style.BorderHT=   "Single"
         Sections(1).Cells(0).Style.BorderHI=   "Single"
         Sections(1).Cells(0).Style.BorderHB=   "Single"
         Sections(1).Cells(0).Style.BorderVL=   "Single"
         Sections(1).Cells(0).Style.BorderVI=   "Single"
         Sections(1).Cells(0).Style.BorderVR=   "Single"
         Sections(1).Cells(0).Style.NoClipping=   0   'False
         Sections(1).Cells(0).Style.RTF=   0   'False
         Sections(1).Cells(0).Style.fprops=   1835009
         Sections(1).Cells(1).Name=   "CELL_2"
         Sections(1).Cells(1).Exp=   """Pokok"""
         Sections(1).Cells(1).PrivateStyle=   -1  'True
         Sections(1).Cells(1).Style.Name=   "<private>"
         Sections(1).Cells(1).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(1).Style.Font_Name=   "Times New Roman"
         Sections(1).Cells(1).Style.Font_Size=   8.25
         Sections(1).Cells(1).Style.Font_Bold=   -1  'True
         Sections(1).Cells(1).Style.Font_Italic=   0   'False
         Sections(1).Cells(1).Style.Font_Underline=   0   'False
         Sections(1).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(1).Style.Font_Charset=   0
         Sections(1).Cells(1).Style.TextAlign=   1
         Sections(1).Cells(1).Style.TextVAlign=   1
         Sections(1).Cells(1).Style.TextWrap=   -1  'True
         Sections(1).Cells(1).Style.ForeColor=   0
         Sections(1).Cells(1).Style.BackColor=   16777215
         Sections(1).Cells(1).Style.NoFill=   -1  'True
         Sections(1).Cells(1).Style.BackPicFile=   ""
         Sections(1).Cells(1).Style.ForePicFile=   ""
         Sections(1).Cells(1).Style.BackPicVertPlacement=   0
         Sections(1).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(1).Style.ForePicPlacement=   0
         Sections(1).Cells(1).Style.ForePicDrawMode=   0
         Sections(1).Cells(1).Style.MarginLeft=   6
         Sections(1).Cells(1).Style.MarginTop=   1
         Sections(1).Cells(1).Style.MarginRight=   6
         Sections(1).Cells(1).Style.MarginBottom=   1
         Sections(1).Cells(1).Style.HasBorders=   -1  'True
         Sections(1).Cells(1).Style.BorderHT=   "Single"
         Sections(1).Cells(1).Style.BorderHI=   "Single"
         Sections(1).Cells(1).Style.BorderHB=   "Single"
         Sections(1).Cells(1).Style.BorderVL=   "Single"
         Sections(1).Cells(1).Style.BorderVI=   "Single"
         Sections(1).Cells(1).Style.BorderVR=   "Single"
         Sections(1).Cells(1).Style.NoClipping=   0   'False
         Sections(1).Cells(1).Style.RTF=   0   'False
         Sections(1).Cells(1).Style.fprops=   1835009
         Sections(1).Cells(2).Name=   "CELL_3"
         Sections(1).Cells(2).Exp=   """Bunga"""
         Sections(1).Cells(2).PrivateStyle=   -1  'True
         Sections(1).Cells(2).Style.Name=   "<private>"
         Sections(1).Cells(2).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(2).Style.Font_Name=   "Times New Roman"
         Sections(1).Cells(2).Style.Font_Size=   8.25
         Sections(1).Cells(2).Style.Font_Bold=   -1  'True
         Sections(1).Cells(2).Style.Font_Italic=   0   'False
         Sections(1).Cells(2).Style.Font_Underline=   0   'False
         Sections(1).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(2).Style.Font_Charset=   0
         Sections(1).Cells(2).Style.TextAlign=   1
         Sections(1).Cells(2).Style.TextVAlign=   1
         Sections(1).Cells(2).Style.TextWrap=   -1  'True
         Sections(1).Cells(2).Style.ForeColor=   0
         Sections(1).Cells(2).Style.BackColor=   16777215
         Sections(1).Cells(2).Style.NoFill=   -1  'True
         Sections(1).Cells(2).Style.BackPicFile=   ""
         Sections(1).Cells(2).Style.ForePicFile=   ""
         Sections(1).Cells(2).Style.BackPicVertPlacement=   0
         Sections(1).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(2).Style.ForePicPlacement=   0
         Sections(1).Cells(2).Style.ForePicDrawMode=   0
         Sections(1).Cells(2).Style.MarginLeft=   6
         Sections(1).Cells(2).Style.MarginTop=   1
         Sections(1).Cells(2).Style.MarginRight=   6
         Sections(1).Cells(2).Style.MarginBottom=   1
         Sections(1).Cells(2).Style.HasBorders=   -1  'True
         Sections(1).Cells(2).Style.BorderHT=   "Single"
         Sections(1).Cells(2).Style.BorderHI=   "Single"
         Sections(1).Cells(2).Style.BorderHB=   "Single"
         Sections(1).Cells(2).Style.BorderVL=   "Single"
         Sections(1).Cells(2).Style.BorderVI=   "Single"
         Sections(1).Cells(2).Style.BorderVR=   "Single"
         Sections(1).Cells(2).Style.NoClipping=   0   'False
         Sections(1).Cells(2).Style.RTF=   0   'False
         Sections(1).Cells(2).Style.fprops=   1835009
         Sections(1).Cells(3).Name=   "CELL_6"
         Sections(1).Cells(3).Exp=   """Denda"""
         Sections(1).Cells(4).Name=   "CELL_4"
         Sections(1).Cells(4).Exp=   """Jumlah"""
         Sections(1).Cells(4).PrivateStyle=   -1  'True
         Sections(1).Cells(4).Style.Name=   "<private>"
         Sections(1).Cells(4).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(4).Style.Font_Name=   "Times New Roman"
         Sections(1).Cells(4).Style.Font_Size=   8.25
         Sections(1).Cells(4).Style.Font_Bold=   -1  'True
         Sections(1).Cells(4).Style.Font_Italic=   0   'False
         Sections(1).Cells(4).Style.Font_Underline=   0   'False
         Sections(1).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(4).Style.Font_Charset=   0
         Sections(1).Cells(4).Style.TextAlign=   1
         Sections(1).Cells(4).Style.TextVAlign=   1
         Sections(1).Cells(4).Style.TextWrap=   -1  'True
         Sections(1).Cells(4).Style.ForeColor=   0
         Sections(1).Cells(4).Style.BackColor=   16777215
         Sections(1).Cells(4).Style.NoFill=   -1  'True
         Sections(1).Cells(4).Style.BackPicFile=   ""
         Sections(1).Cells(4).Style.ForePicFile=   ""
         Sections(1).Cells(4).Style.BackPicVertPlacement=   0
         Sections(1).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(4).Style.ForePicPlacement=   0
         Sections(1).Cells(4).Style.ForePicDrawMode=   0
         Sections(1).Cells(4).Style.MarginLeft=   6
         Sections(1).Cells(4).Style.MarginTop=   1
         Sections(1).Cells(4).Style.MarginRight=   6
         Sections(1).Cells(4).Style.MarginBottom=   1
         Sections(1).Cells(4).Style.HasBorders=   -1  'True
         Sections(1).Cells(4).Style.BorderHT=   "Single"
         Sections(1).Cells(4).Style.BorderHI=   "Single"
         Sections(1).Cells(4).Style.BorderHB=   "Single"
         Sections(1).Cells(4).Style.BorderVL=   "Single"
         Sections(1).Cells(4).Style.BorderVI=   "Single"
         Sections(1).Cells(4).Style.BorderVR=   "Single"
         Sections(1).Cells(4).Style.NoClipping=   0   'False
         Sections(1).Cells(4).Style.RTF=   0   'False
         Sections(1).Cells(4).Style.fprops=   1835009
         Sections(1).Cells(5).Name=   "CELL_1"
         Sections(1).Cells(5).Exp=   """Paraf"""
         Sections(1).Cells(5).PrivateStyle=   -1  'True
         Sections(1).Cells(5).Style.Name=   "<private>"
         Sections(1).Cells(5).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(5).Style.Font_Name=   "Times New Roman"
         Sections(1).Cells(5).Style.Font_Size=   8.25
         Sections(1).Cells(5).Style.Font_Bold=   -1  'True
         Sections(1).Cells(5).Style.Font_Italic=   0   'False
         Sections(1).Cells(5).Style.Font_Underline=   0   'False
         Sections(1).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(5).Style.Font_Charset=   0
         Sections(1).Cells(5).Style.TextAlign=   1
         Sections(1).Cells(5).Style.TextVAlign=   1
         Sections(1).Cells(5).Style.TextWrap=   -1  'True
         Sections(1).Cells(5).Style.ForeColor=   0
         Sections(1).Cells(5).Style.BackColor=   16777215
         Sections(1).Cells(5).Style.NoFill=   -1  'True
         Sections(1).Cells(5).Style.BackPicFile=   ""
         Sections(1).Cells(5).Style.ForePicFile=   ""
         Sections(1).Cells(5).Style.BackPicVertPlacement=   0
         Sections(1).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(5).Style.ForePicPlacement=   0
         Sections(1).Cells(5).Style.ForePicDrawMode=   0
         Sections(1).Cells(5).Style.MarginLeft=   6
         Sections(1).Cells(5).Style.MarginTop=   1
         Sections(1).Cells(5).Style.MarginRight=   6
         Sections(1).Cells(5).Style.MarginBottom=   1
         Sections(1).Cells(5).Style.HasBorders=   -1  'True
         Sections(1).Cells(5).Style.BorderHT=   "Single"
         Sections(1).Cells(5).Style.BorderHI=   "Single"
         Sections(1).Cells(5).Style.BorderHB=   "Single"
         Sections(1).Cells(5).Style.BorderVL=   "Single"
         Sections(1).Cells(5).Style.BorderVI=   "Single"
         Sections(1).Cells(5).Style.BorderVR=   "Single"
         Sections(1).Cells(5).Style.NoClipping=   0   'False
         Sections(1).Cells(5).Style.RTF=   0   'False
         Sections(1).Cells(5).Style.fprops=   1835009
         Sections(2).Name=   "Detail"
         Sections(2).Type=   4
         Sections(2).StyleExp=   "'Tdb_Body'"
         Sections(2).AutoHeight=   0   'False
         Sections(2).Height=   7
         Sections(2).Cells.Count=   6
         Sections(2).Cells(0).Name=   "CELL_0"
         Sections(2).Cells(0).Exp=   "dTgl"
         Sections(2).Cells(0).Width=   18
         Sections(2).Cells(0).Height=   26
         Sections(2).Cells(0).PrivateStyle=   -1  'True
         Sections(2).Cells(0).Style.Name=   "<private>"
         Sections(2).Cells(0).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(0).Style.Font_Name=   "Times New Roman"
         Sections(2).Cells(0).Style.Font_Size=   8.25
         Sections(2).Cells(0).Style.Font_Bold=   0   'False
         Sections(2).Cells(0).Style.Font_Italic=   0   'False
         Sections(2).Cells(0).Style.Font_Underline=   0   'False
         Sections(2).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(0).Style.Font_Charset=   0
         Sections(2).Cells(0).Style.TextAlign=   0
         Sections(2).Cells(0).Style.TextVAlign=   1
         Sections(2).Cells(0).Style.TextWrap=   0   'False
         Sections(2).Cells(0).Style.ForeColor=   0
         Sections(2).Cells(0).Style.BackColor=   16777215
         Sections(2).Cells(0).Style.NoFill=   -1  'True
         Sections(2).Cells(0).Style.BackPicFile=   ""
         Sections(2).Cells(0).Style.ForePicFile=   ""
         Sections(2).Cells(0).Style.BackPicVertPlacement=   0
         Sections(2).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(0).Style.ForePicPlacement=   0
         Sections(2).Cells(0).Style.ForePicDrawMode=   0
         Sections(2).Cells(0).Style.MarginLeft=   6
         Sections(2).Cells(0).Style.MarginTop=   0
         Sections(2).Cells(0).Style.MarginRight=   6
         Sections(2).Cells(0).Style.MarginBottom=   0
         Sections(2).Cells(0).Style.HasBorders=   -1  'True
         Sections(2).Cells(0).Style.BorderHT=   "Single"
         Sections(2).Cells(0).Style.BorderHI=   "Single"
         Sections(2).Cells(0).Style.BorderHB=   "Single"
         Sections(2).Cells(0).Style.BorderVL=   "Single"
         Sections(2).Cells(0).Style.BorderVI=   "Single"
         Sections(2).Cells(0).Style.BorderVR=   "Single"
         Sections(2).Cells(0).Style.NoClipping=   0   'False
         Sections(2).Cells(0).Style.RTF=   0   'False
         Sections(2).Cells(0).Style.fprops=   2084869
         Sections(2).Cells(1).Name=   "CELL_2"
         Sections(2).Cells(1).Exp=   "nPokok"
         Sections(2).Cells(1).Width=   18
         Sections(2).Cells(1).Height=   26
         Sections(2).Cells(1).PrivateStyle=   -1  'True
         Sections(2).Cells(1).Style.Name=   "<private>"
         Sections(2).Cells(1).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(1).Style.Font_Name=   "Times New Roman"
         Sections(2).Cells(1).Style.Font_Size=   8.25
         Sections(2).Cells(1).Style.Font_Bold=   0   'False
         Sections(2).Cells(1).Style.Font_Italic=   0   'False
         Sections(2).Cells(1).Style.Font_Underline=   0   'False
         Sections(2).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(1).Style.Font_Charset=   0
         Sections(2).Cells(1).Style.TextAlign=   3
         Sections(2).Cells(1).Style.TextVAlign=   1
         Sections(2).Cells(1).Style.TextWrap=   0   'False
         Sections(2).Cells(1).Style.ForeColor=   0
         Sections(2).Cells(1).Style.BackColor=   16777215
         Sections(2).Cells(1).Style.NoFill=   -1  'True
         Sections(2).Cells(1).Style.BackPicFile=   ""
         Sections(2).Cells(1).Style.ForePicFile=   ""
         Sections(2).Cells(1).Style.BackPicVertPlacement=   0
         Sections(2).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(1).Style.ForePicPlacement=   0
         Sections(2).Cells(1).Style.ForePicDrawMode=   0
         Sections(2).Cells(1).Style.MarginLeft=   6
         Sections(2).Cells(1).Style.MarginTop=   0
         Sections(2).Cells(1).Style.MarginRight=   6
         Sections(2).Cells(1).Style.MarginBottom=   0
         Sections(2).Cells(1).Style.HasBorders=   -1  'True
         Sections(2).Cells(1).Style.BorderHT=   "Single"
         Sections(2).Cells(1).Style.BorderHI=   "Single"
         Sections(2).Cells(1).Style.BorderHB=   "Single"
         Sections(2).Cells(1).Style.BorderVL=   "Single"
         Sections(2).Cells(1).Style.BorderVI=   "Single"
         Sections(2).Cells(1).Style.BorderVR=   "Single"
         Sections(2).Cells(1).Style.NoClipping=   0   'False
         Sections(2).Cells(1).Style.RTF=   0   'False
         Sections(2).Cells(1).Style.fprops=   2064388
         Sections(2).Cells(2).Name=   "CELL_3"
         Sections(2).Cells(2).Exp=   "nBunga"
         Sections(2).Cells(2).Width=   18
         Sections(2).Cells(2).Height=   26
         Sections(2).Cells(2).PrivateStyle=   -1  'True
         Sections(2).Cells(2).Style.Name=   "<private>"
         Sections(2).Cells(2).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(2).Style.Font_Name=   "Times New Roman"
         Sections(2).Cells(2).Style.Font_Size=   8.25
         Sections(2).Cells(2).Style.Font_Bold=   0   'False
         Sections(2).Cells(2).Style.Font_Italic=   0   'False
         Sections(2).Cells(2).Style.Font_Underline=   0   'False
         Sections(2).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(2).Style.Font_Charset=   0
         Sections(2).Cells(2).Style.TextAlign=   2
         Sections(2).Cells(2).Style.TextVAlign=   1
         Sections(2).Cells(2).Style.TextWrap=   0   'False
         Sections(2).Cells(2).Style.ForeColor=   0
         Sections(2).Cells(2).Style.BackColor=   16777215
         Sections(2).Cells(2).Style.NoFill=   -1  'True
         Sections(2).Cells(2).Style.BackPicFile=   ""
         Sections(2).Cells(2).Style.ForePicFile=   ""
         Sections(2).Cells(2).Style.BackPicVertPlacement=   0
         Sections(2).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(2).Style.ForePicPlacement=   0
         Sections(2).Cells(2).Style.ForePicDrawMode=   0
         Sections(2).Cells(2).Style.MarginLeft=   6
         Sections(2).Cells(2).Style.MarginTop=   0
         Sections(2).Cells(2).Style.MarginRight=   6
         Sections(2).Cells(2).Style.MarginBottom=   0
         Sections(2).Cells(2).Style.HasBorders=   -1  'True
         Sections(2).Cells(2).Style.BorderHT=   "Single"
         Sections(2).Cells(2).Style.BorderHI=   "Single"
         Sections(2).Cells(2).Style.BorderHB=   "Single"
         Sections(2).Cells(2).Style.BorderVL=   "Single"
         Sections(2).Cells(2).Style.BorderVI=   "Single"
         Sections(2).Cells(2).Style.BorderVR=   "Single"
         Sections(2).Cells(2).Style.NoClipping=   0   'False
         Sections(2).Cells(2).Style.RTF=   0   'False
         Sections(2).Cells(2).Style.fprops=   2064389
         Sections(2).Cells(3).Name=   "CELL_6"
         Sections(2).Cells(3).Exp=   "nDenda"
         Sections(2).Cells(3).Width=   18
         Sections(2).Cells(3).PrivateStyle=   -1  'True
         Sections(2).Cells(3).Style.Name=   "<private>"
         Sections(2).Cells(3).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(3).Style.Font_Name=   "Times New Roman"
         Sections(2).Cells(3).Style.Font_Size=   8.25
         Sections(2).Cells(3).Style.Font_Bold=   0   'False
         Sections(2).Cells(3).Style.Font_Italic=   0   'False
         Sections(2).Cells(3).Style.Font_Underline=   0   'False
         Sections(2).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(3).Style.Font_Charset=   0
         Sections(2).Cells(3).Style.TextAlign=   3
         Sections(2).Cells(3).Style.TextVAlign=   1
         Sections(2).Cells(3).Style.TextWrap=   -1  'True
         Sections(2).Cells(3).Style.ForeColor=   0
         Sections(2).Cells(3).Style.BackColor=   16777215
         Sections(2).Cells(3).Style.NoFill=   -1  'True
         Sections(2).Cells(3).Style.BackPicFile=   ""
         Sections(2).Cells(3).Style.ForePicFile=   ""
         Sections(2).Cells(3).Style.BackPicVertPlacement=   0
         Sections(2).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(3).Style.ForePicPlacement=   0
         Sections(2).Cells(3).Style.ForePicDrawMode=   0
         Sections(2).Cells(3).Style.MarginLeft=   6
         Sections(2).Cells(3).Style.MarginTop=   0
         Sections(2).Cells(3).Style.MarginRight=   6
         Sections(2).Cells(3).Style.MarginBottom=   0
         Sections(2).Cells(3).Style.HasBorders=   -1  'True
         Sections(2).Cells(3).Style.BorderHT=   "Single"
         Sections(2).Cells(3).Style.BorderHI=   "Single"
         Sections(2).Cells(3).Style.BorderHB=   "Single"
         Sections(2).Cells(3).Style.BorderVL=   "Single"
         Sections(2).Cells(3).Style.BorderVI=   "Single"
         Sections(2).Cells(3).Style.BorderVR=   "Single"
         Sections(2).Cells(3).Style.NoClipping=   0   'False
         Sections(2).Cells(3).Style.RTF=   0   'False
         Sections(2).Cells(3).Style.fprops=   2064384
         Sections(2).Cells(4).Name=   "CELL_4"
         Sections(2).Cells(4).Exp=   "nJumlah"
         Sections(2).Cells(4).Width=   18
         Sections(2).Cells(4).Height=   26
         Sections(2).Cells(4).CallExpression=   -1  'True
         Sections(2).Cells(4).PrivateStyle=   -1  'True
         Sections(2).Cells(4).Format=   "###,###,###,###,###,##0.00"
         Sections(2).Cells(4).Style.Name=   "<private>"
         Sections(2).Cells(4).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(4).Style.Font_Name=   "Times New Roman"
         Sections(2).Cells(4).Style.Font_Size=   8.25
         Sections(2).Cells(4).Style.Font_Bold=   0   'False
         Sections(2).Cells(4).Style.Font_Italic=   0   'False
         Sections(2).Cells(4).Style.Font_Underline=   0   'False
         Sections(2).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(4).Style.Font_Charset=   0
         Sections(2).Cells(4).Style.TextAlign=   2
         Sections(2).Cells(4).Style.TextVAlign=   1
         Sections(2).Cells(4).Style.TextWrap=   0   'False
         Sections(2).Cells(4).Style.ForeColor=   0
         Sections(2).Cells(4).Style.BackColor=   16777215
         Sections(2).Cells(4).Style.NoFill=   -1  'True
         Sections(2).Cells(4).Style.BackPicFile=   ""
         Sections(2).Cells(4).Style.ForePicFile=   ""
         Sections(2).Cells(4).Style.BackPicVertPlacement=   0
         Sections(2).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(4).Style.ForePicPlacement=   0
         Sections(2).Cells(4).Style.ForePicDrawMode=   0
         Sections(2).Cells(4).Style.MarginLeft=   6
         Sections(2).Cells(4).Style.MarginTop=   0
         Sections(2).Cells(4).Style.MarginRight=   6
         Sections(2).Cells(4).Style.MarginBottom=   0
         Sections(2).Cells(4).Style.HasBorders=   -1  'True
         Sections(2).Cells(4).Style.BorderHT=   "Single"
         Sections(2).Cells(4).Style.BorderHI=   "Single"
         Sections(2).Cells(4).Style.BorderHB=   "Single"
         Sections(2).Cells(4).Style.BorderVL=   "Single"
         Sections(2).Cells(4).Style.BorderVI=   "Single"
         Sections(2).Cells(4).Style.BorderVR=   "Single"
         Sections(2).Cells(4).Style.NoClipping=   0   'False
         Sections(2).Cells(4).Style.RTF=   0   'False
         Sections(2).Cells(4).Style.fprops=   2064389
         Sections(2).Cells(5).Name=   "CELL_1"
         Sections(2).Cells(5).Exp=   "cParaf"
         Sections(2).Cells(5).Width=   18
         Sections(2).Cells(5).Height=   26
         Sections(2).Cells(5).PrivateStyle=   -1  'True
         Sections(2).Cells(5).Style.Name=   "<private>"
         Sections(2).Cells(5).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(5).Style.Font_Name=   "Times New Roman"
         Sections(2).Cells(5).Style.Font_Size=   8.25
         Sections(2).Cells(5).Style.Font_Bold=   0   'False
         Sections(2).Cells(5).Style.Font_Italic=   0   'False
         Sections(2).Cells(5).Style.Font_Underline=   0   'False
         Sections(2).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(5).Style.Font_Charset=   0
         Sections(2).Cells(5).Style.TextAlign=   1
         Sections(2).Cells(5).Style.TextVAlign=   1
         Sections(2).Cells(5).Style.TextWrap=   0   'False
         Sections(2).Cells(5).Style.ForeColor=   0
         Sections(2).Cells(5).Style.BackColor=   16777215
         Sections(2).Cells(5).Style.NoFill=   -1  'True
         Sections(2).Cells(5).Style.BackPicFile=   ""
         Sections(2).Cells(5).Style.ForePicFile=   ""
         Sections(2).Cells(5).Style.BackPicVertPlacement=   0
         Sections(2).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(5).Style.ForePicPlacement=   0
         Sections(2).Cells(5).Style.ForePicDrawMode=   0
         Sections(2).Cells(5).Style.MarginLeft=   6
         Sections(2).Cells(5).Style.MarginTop=   0
         Sections(2).Cells(5).Style.MarginRight=   6
         Sections(2).Cells(5).Style.MarginBottom=   0
         Sections(2).Cells(5).Style.HasBorders=   -1  'True
         Sections(2).Cells(5).Style.BorderHT=   "Single"
         Sections(2).Cells(5).Style.BorderHI=   "Single"
         Sections(2).Cells(5).Style.BorderHB=   "Single"
         Sections(2).Cells(5).Style.BorderVL=   "Single"
         Sections(2).Cells(5).Style.BorderVI=   "Single"
         Sections(2).Cells(5).Style.BorderVR=   "Single"
         Sections(2).Cells(5).Style.NoClipping=   0   'False
         Sections(2).Cells(5).Style.RTF=   0   'False
         Sections(2).Cells(5).Style.fprops=   2064389
         Styles.Count    =   6
         Styles(0).Name  =   "Tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Size=   8.25
         Styles(0).Font_Bold=   -1  'True
         Styles(0).Font_Charset=   0
         Styles(0).TextVAlign=   1
         Styles(0).MarginTop=   1
         Styles(0).MarginBottom=   1
         Styles(1).Name  =   "STYLE_1"
         Styles(1).ParentName=   "Tdb_Base"
         Styles(1).Font_Size=   8.25
         Styles(1).Font_Charset=   0
         Styles(1).TextVAlign=   1
         Styles(1).MarginTop=   1
         Styles(1).MarginBottom=   1
         Styles(1).fprops=   18087936
         Styles(2).Name  =   "Tdb_Body"
         Styles(2).ParentName=   "Tdb_Base"
         Styles(2).Font_Size=   8.25
         Styles(2).Font_Charset=   0
         Styles(2).TextVAlign=   1
         Styles(2).MarginTop=   0
         Styles(2).MarginBottom=   0
         Styles(2).fprops=   18862080
         Styles(3).Name  =   "Tdb_Header"
         Styles(3).ParentName=   "Tdb_Base"
         Styles(3).Font_Size=   8.25
         Styles(3).Font_Bold=   -1  'True
         Styles(3).Font_Charset=   0
         Styles(3).TextAlign=   0
         Styles(3).TextVAlign=   1
         Styles(3).MarginTop=   1
         Styles(3).MarginBottom=   1
         Styles(3).BorderHT=   "Single"
         Styles(3).BorderHI=   "Single"
         Styles(3).BorderHB=   "Single"
         Styles(3).fprops=   2064385
         Styles(4).Name  =   "Tdb_PageFooter"
         Styles(4).ParentName=   "Tdb_Base"
         Styles(4).Font_Size=   8.25
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).TextVAlign=   1
         Styles(4).MarginTop=   1
         Styles(4).MarginBottom=   1
         Styles(4).BorderHT=   "Single"
         Styles(4).fprops=   163840
         Styles(5).Name  =   "Garis"
         Styles(5).ParentName=   "Tdb_Base"
         Styles(5).Font_Size=   8.25
         Styles(5).Font_Bold=   -1  'True
         Styles(5).Font_Charset=   0
         Styles(5).TextAlign=   2
         Styles(5).TextVAlign=   1
         Styles(5).MarginTop=   1
         Styles(5).MarginBottom=   1
         Styles(5).BorderHT=   "Single"
         Styles(5).fprops=   32769
         Mappings.Count  =   1
         Mappings(0).Name=   "MAPPING_0"
         Lines.Count     =   4
         Lines(0).Name   =   "Single"
         Lines(0).Thickness=   4
         Lines(1).Name   =   "Double"
         Lines(1).Thickness=   5
         Lines(2).Name   =   "Quarter"
         Lines(2).Thickness=   1
         Lines(2).Color  =   8421504
         Lines(3).Name   =   "None"
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginLeft=   10
         Profiles(0).PrinterMarginTop=   5
         Profiles(0).PrinterMarginRight=   10
         Profiles(0).PrinterMarginBottom=   5
         Profiles(0).PrinterPaperSize=   256
         Profiles(0).PrinterPaperHeight=   210
         Profiles(0).PrinterPaperWidth=   160
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterPaperSize_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   4980
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
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak Buku Angsuran Kosong"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         TabIndex        =   10
         Top             =   120
         Width           =   3015
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Height          =   435
         Left            =   10530
         TabIndex        =   6
         Top             =   105
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
         Picture         =   "RptBukuAngsuran.frx":0EAA
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   9360
         TabIndex        =   7
         Top             =   105
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
         Picture         =   "RptBukuAngsuran.frx":0F50
      End
      Begin BiSAButtonProject.BiSAButton cmdRefresh 
         Height          =   435
         Left            =   8190
         TabIndex        =   9
         Top             =   105
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
         Picture         =   "RptBukuAngsuran.frx":11D6
      End
   End
End
Attribute VB_Name = "RptBukuAngsuran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim dbRekening As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim xArray As New XArrayDB
Dim nAwal As Double
Dim lPreview As Boolean
Dim nPlafond As Double
Dim nLama As Integer
Dim cRekening As String
Dim vaRPT As New XArrayDB
Dim cJaminan As String

Dim dTglRealisasi As Date
Dim dTglJatuhTempo As Date
Dim nSukuBunga As Double
Dim nKewajibanPokok As Double
Dim nKewajibanBunga As Double
Dim nTotalKewajiban As Double

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganKredit", "Kode", cGolongan, "Kode,Keterangan")
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
  GetSQL
End Sub

Private Sub cmdRefresh_Click()
  lPreview = False
  GetSQL
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
  If cFrekuensi.LastKey = 13 Then
    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    Set dbData = objData.Browse(GetDSN, "Debitur t", "t.*,r.Nama,r.Alamat", "t.Rekening", sisAssign, cRekening, , , Array("Left Join RegisterNasabah r on r.Kode=t.Kode"))
    If Not dbData.eof Then
      GetData
    Else
      MsgBox "Rekening dengan nomor: " & cRekening & " Tidak ada !", vbOKOnly, "Laporan Buku Tabungan"
      xArray.Clear
      xArray.ReDim 0, -1, 0, 6
      Set TDBGrid1.Array = xArray
      TDBGrid1.ReBind
      cFrekuensi.Default
      Cancel = True
      cFrekuensi.SetFocus
      Exit Sub
    End If
  End If
End Sub

Private Sub GetData()
  cNama.Text = GetNull(dbData!nama, "")
  cAlamat.Text = GetNull(dbData!alamat, "")
  nPlafond = GetNull(dbData!plafond)
  nLama = GetNull(dbData!Lama)
  dTglRealisasi = GetNull(dbData!Tgl)
  dTglJatuhTempo = GetNull(dbData!JatuhTempo)
  nSukuBunga = Devide(GetNull(dbData!SukuBunga), 12)
  nKewajibanPokok = Devide(GetNull(dbData!plafond), nLama)
  nKewajibanBunga = nPlafond * nSukuBunga / 100
  nTotalKewajiban = nKewajibanBunga + nKewajibanPokok
End Sub

Private Function GetBungaReguler(ByVal nSisaPokok As Double, ByVal nBunga As Double) As Double
  GetBungaReguler = nSisaPokok * (nBunga / 100)
  GetBungaReguler = Mod50(GetBungaReguler)
End Function

Private Sub Command1_Click()
    PrintSQL
    With RptFakturPenjualan
      .Parameters("cNamaDebitur").ValueExpression = "'" & cNama.Text & "'"
      .Parameters("cNoRekening").ValueExpression = "'" & cRekening & "'"
      
      .Parameters("cJaminan").ValueExpression = "'" & GetNamaJaminan(objData, cRekening) & "'"
      .Parameters("cJangkaWaktu").ValueExpression = "'" & nLama & " Bulan'"
      .Parameters("cSukuBunga").ValueExpression = "'" & nSukuBunga & "'"
      .Parameters("dTglRealisasi").ValueExpression = "'" & Format(dTglRealisasi, "dd/MM/yyyy") & "'"
      .Parameters("dTglJatuhTempo").ValueExpression = "'" & Format(dTglJatuhTempo, "dd/MM/yyyy") & "'"
      .Parameters("nKewajibanPokok").ValueExpression = "'" & Format(nKewajibanPokok, "###,###,###,##0.00") & "'"
      .Parameters("nKewajibanBunga").ValueExpression = "'" & Format(nKewajibanBunga, "###,###,###,##0.00") & "'"
      .Parameters("nJumlahKewajiban").ValueExpression = "'" & Format(nTotalKewajiban, "###,###,###,##0.00") & "'"
      .Parameters("nPlafond").ValueExpression = "'" & Format(nPlafond, "###,###,###,##0.00") & "'"
      Set .Array = vaRPT
      .Refresh
      .PrintPreview
    End With
End Sub

Private Function GetNamaJaminan(ByVal obj As CodeSuiteLibrary.data, ByVal Rek As String) As String
Dim db As New ADODB.Recordset
  
  GetNamaJaminan = ""
  Set db = objData.Browse(GetDSN, "agunan a", "a.kode,g.keterangan", "a.rekening", sisAssign, cRekening, , , Array("left join gagunan g on g.kode = a.kode"))
  If Not db.eof Then
    GetNamaJaminan = GetNull(db!Keterangan)
  End If
End Function

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub initvalue()
  cCabang.Text = aCfg(msKodeCabang, "")
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
End Sub

Private Sub Form_Load()
Dim n As Single

  InitGrid TDBGrid1
  CenterForm Me, True
  initvalue
  
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n
  TabIndex cmdRefresh, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSQL()
Dim nTotal As Double
Dim n As Long
Dim nPokok As Double
Dim nBunga As Double
Dim nDenda As Double
Dim nPlafond As Double
Dim nTmpBakiDebet As Double

  xArray.ReDim 0, -1, 0, 9
  Set dbData = objData.Browse(GetDSN, "Angsuran a", "a.*,d.Plafond,d.NoSPK", "a.Rekening", sisAssign, UCase(cRekening), , "Tgl", _
                              Array("left Join Debitur d on d.Rekening=a.Rekening"))
  If Not dbData.eof Then
    dbData.MoveFirst
    nPokok = 0
    nBunga = 0
    nDenda = 0
    nTotal = 0
    nPlafond = 0
    n = 0
    'xArray.ReDim 0, dbData.RecordCount - 1, 0, 9
    xArray.ReDim 0, -1, 0, 9
    nTmpBakiDebet = GetNull(dbData!plafond)
    Do While Not dbData.eof
        xArray.InsertRows xArray.UpperBound(1) + 1
        n = xArray.UpperBound(1)
'        xarray(n, 0) = IIf((dbData!Pokok) < 0, (dbData!Faktur) & " Tarik", (dbData!Faktur))
'        xarray(n, 1) = (dbData!Tgl)
'
'        xarray(n, 2) = (dbData!Pokok)
'        xarray(n, 3) = (dbData!Bunga)
'        xarray(n, 4) = (dbData!Denda)
'        xarray(n, 5) = (dbData!Total)
'        xarray(n, 6) = GetBakiDebet(objData, cRekening, GetNull(dbData!Plafond), GetNull(dbData!Tgl))
'
'        nPokok = nPokok + xarray(n, 2)
'        nBunga = nBunga + xarray(n, 3)
'        nDenda = nDenda + xarray(n, 4)
'        nTotal = nTotal + xarray(n, 5)
        xArray(n, 0) = IIf((dbData!pokok) < 0, (dbData!Faktur) & " Tarik", (dbData!Faktur))
        xArray(n, 1) = (dbData!Tgl)
        xArray(n, 2) = (dbData!NoSPK)
        xArray(n, 3) = (dbData!plafond)
        xArray(n, 4) = (GetTglPencairan(objData, cRekening))
        xArray(n, 5) = (dbData!pokok)
        xArray(n, 6) = (dbData!bunga)
        xArray(n, 7) = (dbData!denda)
        xArray(n, 8) = (dbData!Total)
        xArray(n, 9) = nTmpBakiDebet - xArray(n, 5) 'GetBakiDebet(objData, cRekening, GetNull(dbData!Plafond), GetNull(dbData!Tgl))
        nTmpBakiDebet = xArray(n, 9)
        
        nPokok = nPokok + xArray(n, 5)
        nBunga = nBunga + xArray(n, 6)
        nDenda = nDenda + xArray(n, 7)
        nTotal = nTotal + xArray(n, 8)
        nPlafond = nPlafond + GetNull(xArray(n, 3))
        
        dbData.MoveNext
        'n = n + 1
    Loop
    
    TDBGrid1.Columns(1).Merge = True
    TDBGrid1.Columns(2).Merge = True
    TDBGrid1.Columns(3).Merge = True
    TDBGrid1.Columns(4).Merge = True
    'TDBGrid1.Columns(3).FooterText = Format(nPlafond, "###,###,###,###,##0.00")
    TDBGrid1.Columns(5).FooterText = Format(nPokok, "###,###,###,###,##0.00")
    TDBGrid1.Columns(6).FooterText = Format(nBunga, "###,###,###,###,##0.00")
    TDBGrid1.Columns(7).FooterText = Format(nDenda, "###,###,###,###,##0.00")
    TDBGrid1.Columns(8).FooterText = Format(nTotal, "###,###,###,###,##0.00")
    TDBGrid1.Array = xArray
    TDBGrid1.ReBind
    If lPreview = True Then
      rpt
    End If
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Sub
  End If
End Sub

Private Sub PrintSQL()
Dim n As Integer
Dim i As Integer
  
  vaRPT.ReDim 0, -1, 0, 5
  i = 0
  Do While i < 48
    vaRPT.InsertRows vaRPT.UpperBound(1) + 1
    n = vaRPT.UpperBound(1)
    If n - ((n \ 2) * 2) > 0 Then 'ganjil
      vaRPT(n, 0) = "SISA"
      vaRPT(n, 1) = ""
      vaRPT(n, 2) = ""
      vaRPT(n, 3) = ""
      vaRPT(n, 4) = ""
      vaRPT(n, 5) = ""
    Else
      vaRPT(n, 0) = " "
      vaRPT(n, 1) = ""
      vaRPT(n, 2) = ""
      vaRPT(n, 3) = ""
      vaRPT(n, 4) = ""
      vaRPT(n, 5) = ""
    End If
    i = i + 1
  Loop
End Sub

Private Function GetTglPencairan(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String) As Date
Dim dbPencairan As New ADODB.Recordset

GetTglPencairan = Date
Set dbPencairan = obj.Browse(GetDSN, "PencairanKredit", "Tgl", "Rekening", sisAssign, cRekening)
  If Not dbPencairan.eof Then
    GetTglPencairan = GetNull(dbPencairan!Tgl)
  End If
End Function

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  Select Case ColIndex
    Case 0
      Value = Format(Value, "dd-mm-yyyy")
    Case Else
      If Value = 0 Then
        Value = ""
      Else
        Value = Format(Value, "###,###,##0.00")
      End If
  End Select
End Sub

Private Sub rpt()

  With FrmRPT
    .AddPageHeader UCase("Laporan Buku Angsuran"), tdbHalignCenter, , , , , 14, True
    .AddPageHeader "Nama Debitur", tdbHalignLeft, , 15, , , , , , True, , tdbPageHeaderSect
    .AddPageHeader ": " & cNama.Text
    .AddPageHeader "Alamat Debitur", tdbHalignLeft, , 15, True, , , , , True, , tdbPageHeaderSect
    .AddPageHeader ": " & cAlamat.Text
    .AddPageHeader "Nomor Rekening", tdbHalignLeft, , 15, True, , , , , True, , tdbPageHeaderSect
    .AddPageHeader ": " & cRekening
    .AddPageHeader "Jumlah Plafond", tdbHalignLeft, , 15, True, , , , , True, , tdbPageHeaderSect
    .AddPageHeader ": " & Format(nPlafond, "###,###,###,###,##0.00")
    .AddPageHeader "Lama Angsuran", tdbHalignLeft, , 15, True, , , , , True, , tdbPageHeaderSect
    .AddPageHeader ": " & nLama & " Bulan"
    
    .AddTableHeader "No. Transaksi", , , , 17, , , , , , True, tdbTableHeaderSect
    .AddTableHeader "Tanggal", , , , 8
    .AddTableHeader "No. SPK"
    .AddTableHeader "Plafond"
    .AddTableHeader "Tgl Cair"
    .AddTableHeader "Pokok"
    .AddTableHeader "Bunga"
    .AddTableHeader "Denda"
    .AddTableHeader "Total"
    .AddTableHeader "Baki Debet"
    
'        xArray(n, 0) = IIf((dbData!Pokok) < 0, (dbData!Faktur) & " Tarik", (dbData!Faktur))
'        xArray(n, 1) = (dbData!Tgl)
'        xArray(n, 2) = (dbData!NoSPK)
'        xArray(n, 3) = (dbData!Plafond)
'        xArray(n, 4) = (GetTglPencairan(objData, cRekening))
'        xArray(n, 5) = (dbData!Pokok)
'        xArray(n, 6) = (dbData!bunga)
'        xArray(n, 7) = (dbData!Denda)
'        xArray(n, 8) = (dbData!Total)
'        xArray(n, 9) = nTmpBakiDebet - xArray(n, 5)

    .AddTableBody
    .AddTableBody
    .AddTableBody , tdbHalignCenter, , , , , , , , , , tdbMergeOnText
    .AddTableBody Sis_Rpt_Number2, , , , , , , , , , , tdbMergeOnText
    .AddTableBody , , , , , , , , , , , tdbMergeOnText
    .AddTableBody Sis_Rpt_Number2, , , , , , , , , , , tdbMergeOnText
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableFooter "Total", , tdbHalignCenter, , , , , , , , , , , , 5
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    
    .Refresh
    .Preview xArray, True, , True
  End With
End Sub
Private Sub rpt2()

  With FrmRPT
'    .AddPageHeader UCase("Laporan Buku Angsuran"), tdbHalignCenter, , , , , 14, True
'    .AddPageHeader "Nama Debitur", tdbHalignLeft, , 10, , , , , , True, , tdbPageHeaderSect
'    .AddPageHeader ": " & cNama.Text
'    .AddPageHeader "Alamat Debitur", tdbHalignLeft, , 10, True, , , , , True, , tdbPageHeaderSect
'    .AddPageHeader ": " & cAlamat.Text
'    .AddPageHeader "Nomor Rekening", tdbHalignLeft, , 10, True, , , , , True, , tdbPageHeaderSect
'    .AddPageHeader ": " & cRekening
'    .AddPageHeader "Plafond", tdbHalignLeft, , 10, True, , , , , True, , tdbPageHeaderSect
'    .AddPageHeader ": " & Format(nPlafond, "###,###,###,###,##0.00")
'    .AddPageHeader "Jangka Waktu", tdbHalignLeft, , 10, True, , , , , True, , tdbPageHeaderSect
'    .AddPageHeader ": " & nLama & " Bulan"
    .AddTableHeader "Test", , , , , , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    
    .AddTableHeader "TGL", , , , 15
    .AddTableHeader "POKOK", , , , 19
    .AddTableHeader "BUNGA", , , , 18
    .AddTableHeader "JUMLAH"
    .AddTableHeader "PARAF", , , , 12
    
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .Refresh
    .Preview vaRPT, , False, , , , , , True, 200, 160
  End With
  
End Sub
