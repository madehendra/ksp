VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form MstJaminan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER JAMINAN"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11685
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4455
      Left            =   15
      Top             =   1380
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   7858
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
      Begin BiSANumberBoxProject.BiSANumberBox nLebar 
         Height          =   330
         Left            =   5010
         TabIndex        =   15
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
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
      Begin BiSATextBoxProject.BiSATextBox cJudul 
         Height          =   330
         Left            =   570
         TabIndex        =   14
         Top             =   75
         Width           =   4440
         _ExtentX        =   7832
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3945
         Left            =   75
         TabIndex        =   0
         Top             =   450
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   6959
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
         Columns(1).Caption=   "JUDUL KOLOM"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "LEBAR"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "###,###,##0.00"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   0
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=847"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=7858"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=7779"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2117"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2037"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
         HeadLines       =   1
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
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Named:id=33:Normal"
         _StyleDefs(50)  =   ":id=33,.parent=0"
         _StyleDefs(51)  =   "Named:id=34:Heading"
         _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   ":id=34,.wraptext=-1"
         _StyleDefs(54)  =   "Named:id=35:Footing"
         _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=36:Selected"
         _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=37:Caption"
         _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nNo 
         Height          =   330
         Left            =   90
         TabIndex        =   13
         Top             =   75
         Width           =   465
         _ExtentX        =   820
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   345
         Left            =   5850
         TabIndex        =   16
         Top             =   75
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
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
         Picture         =   "MstJaminan.frx":0000
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1380
      Left            =   0
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   2434
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
      Begin VB.OptionButton optJenis 
         Caption         =   "&1. Kendaraan"
         Height          =   330
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   915
         Width           =   1335
      End
      Begin VB.OptionButton optJenis 
         Caption         =   "&2. Lainnya"
         Height          =   330
         Index           =   1
         Left            =   3480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   915
         Width           =   1125
      End
      Begin BiSATextBoxProject.BiSATextBox cKode 
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   120
         Width           =   2490
         _ExtentX        =   4392
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
         Caption         =   "KODE"
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   495
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
      Begin VB.Label Label1 
         Caption         =   "JENIS JAMINAN"
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
         Left            =   180
         TabIndex        =   5
         Top             =   960
         Width           =   1680
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   4455
      Left            =   6390
      Top             =   1380
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   7858
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   4305
         Left            =   75
         TabIndex        =   6
         Top             =   75
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   7594
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
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1376"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1296"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=7117"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=7038"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "DAFTAR JENIS JAMINAN"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Named:id=33:Normal"
         _StyleDefs(46)  =   ":id=33,.parent=0"
         _StyleDefs(47)  =   "Named:id=34:Heading"
         _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   ":id=34,.wraptext=-1"
         _StyleDefs(50)  =   "Named:id=35:Footing"
         _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   "Named:id=36:Selected"
         _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=39:EvenRow"
         _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(60)  =   "Named:id=40:OddRow"
         _StyleDefs(61)  =   ":id=40,.parent=33"
         _StyleDefs(62)  =   "Named:id=41:RecordSelector"
         _StyleDefs(63)  =   ":id=41,.parent=34"
         _StyleDefs(64)  =   "Named:id=42:FilterBar"
         _StyleDefs(65)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5850
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
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2220
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
         Picture         =   "MstJaminan.frx":01AA
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3375
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
         Picture         =   "MstJaminan.frx":0434
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9435
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
         Picture         =   "MstJaminan.frx":05D3
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
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
         Picture         =   "MstJaminan.frx":09E9
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
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
         Picture         =   "MstJaminan.frx":0B15
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10515
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
         Picture         =   "MstJaminan.frx":0CC0
      End
   End
End
Attribute VB_Name = "MstJaminan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim cJ1 As String, cJ2 As String, cJ3 As String, cJ4 As String, cJ5 As String
Dim cJ6 As String, cJ7 As String, cJ8 As String, cJ9 As String, cJ10 As String
Dim cJ11 As String, cJ12 As String, cJ13 As String, cJ14 As String, cJ15 As String
Dim n As Single
Dim lEdit As Boolean
Dim lClick As Boolean
Dim nPos As SisPos

Private Sub cKode_Validate(Cancel As Boolean)
    If cKode.LastKey = 13 Then
        Set dbData = objData.Browse(GetDSN, "gagunan", "Kode", "Kode", sisAssign, cKode.Text)
        If Not dbData.eof Then
          If nPos = Add Then
            MsgBox "Kode Sudah Ada"
            Cancel = True
            cKode.Default
            cKode.SetFocus
          End If
          GetMemory
          If nPos = Delete Then DeleteData
        ElseIf dbData.RecordCount = 0 And nPos <> Add Then
          MsgBox "data Tidak ada", vbInformation
          Cancel = True
          cKode.SetFocus
          Exit Sub
        End If
    End If
End Sub

Private Sub DeleteData()
  If MsgBox("Data Dihapus ?", vbYesNo) = vbYes Then
    objData.Delete GetDSN, "gAgunan", "kode", sisAssign, cKode.Text
  End If
  initvalue
  GetEdit False
  GetSQL
End Sub

Private Sub GetMemory()
Dim cSQL As String
  cSQL = "select kode,keterangan,Jenis,j1,j2,j3,j4,j5,j6,j7,j8,j9,j10,j11,j12,j13,j14,j15 from gagunan"
  cSQL = cSQL & " where kode='" & cKode.Text & "'"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    cKeterangan.Text = GetNull(dbData!Keterangan)
    SetOpt optJenis, GetNull(dbData!Jenis)
    With dbData
      vaArray(0, 1) = left(GetNull(!j1), 47)
      vaArray(0, 2) = Mid(GetNull(!j1), 48)
      vaArray(1, 1) = left(GetNull(!j2), 47)
      vaArray(1, 2) = Mid(GetNull(!j2), 48)
      vaArray(2, 1) = left(GetNull(!j3), 47)
      vaArray(2, 2) = Mid(GetNull(!j3), 48)
      vaArray(3, 1) = left(GetNull(!j4), 47)
      vaArray(3, 2) = Mid(GetNull(!j4), 48)
      vaArray(4, 1) = left(GetNull(!j5), 47)
      vaArray(4, 2) = Mid(GetNull(!j5), 48)
      vaArray(5, 1) = left(GetNull(!j6), 47)
      vaArray(5, 2) = Mid(GetNull(!j6), 48)
      vaArray(6, 1) = left(GetNull(!j7), 47)
      vaArray(6, 2) = Mid(GetNull(!j7), 48)
      vaArray(7, 1) = left(GetNull(!j8), 47)
      vaArray(7, 2) = Mid(GetNull(!j8), 48)
      vaArray(8, 1) = left(GetNull(!j9), 47)
      vaArray(8, 2) = Mid(GetNull(!j9), 48)
      vaArray(9, 1) = left(GetNull(!j10), 47)
      vaArray(9, 2) = Mid(GetNull(!j10), 48)
      vaArray(10, 1) = left(GetNull(!j11), 47)
      vaArray(10, 2) = Mid(GetNull(!j11), 48)
      vaArray(11, 1) = left(GetNull(!j12), 47)
      vaArray(11, 2) = Mid(GetNull(!j12), 48)
      vaArray(12, 1) = left(GetNull(!j13), 47)
      vaArray(12, 2) = Mid(GetNull(!j13), 48)
      vaArray(13, 1) = left(GetNull(!j14), 47)
      vaArray(13, 2) = Mid(GetNull(!j14), 48)
      vaArray(14, 1) = left(GetNull(!j15), 47)
      vaArray(14, 2) = Mid(GetNull(!j15), 48)
    End With
  End If
  
  For n = 0 To vaArray.UpperBound(1)
    vaArray(n, 0) = n + 1
  Next
  
  TDBGrid1.DataSource = vaArray
  TDBGrid1.ReBind
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
  vaArray(n, 1) = cJudul.Text
  vaArray(n, 2) = nLebar.Value
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  
  Initdetail
  cJudul.SetFocus
  Exit Sub
End Sub

Private Sub cmdSimpan_Click()
  Dim vaValue As Variant, vaField As Variant
  
  TDBGrid1.MoveFirst
  cJ1 = Gabung(vaArray(0, 1), vaArray(0, 2))
  cJ2 = Gabung(vaArray(1, 1), vaArray(1, 2))
  cJ3 = Gabung(vaArray(2, 1), vaArray(2, 2))
  cJ4 = Gabung(vaArray(3, 1), vaArray(3, 2))
  cJ5 = Gabung(vaArray(4, 1), vaArray(4, 2))
  cJ6 = Gabung(vaArray(5, 1), vaArray(5, 2))
  cJ7 = Gabung(vaArray(6, 1), vaArray(6, 2))
  cJ8 = Gabung(vaArray(7, 1), vaArray(7, 2))
  cJ9 = Gabung(vaArray(8, 1), vaArray(8, 2))
  cJ10 = Gabung(vaArray(9, 1), vaArray(9, 2))
  cJ11 = Gabung(vaArray(10, 1), vaArray(10, 2))
  cJ12 = Gabung(vaArray(11, 1), vaArray(11, 2))
  cJ13 = Gabung(vaArray(12, 1), vaArray(12, 2))
  cJ14 = Gabung(vaArray(13, 1), vaArray(13, 2))
  cJ15 = Gabung(vaArray(14, 1), vaArray(14, 2))
  
  vaField = Array("kode", "keterangan", "Jenis", "J1", "J2", "J3", "J4", "J5", "J6", "J7", _
          "J8", "J9", "J10", "J11", "J12", "J13", "J14", "J15")
  vaValue = Array(cKode.Text, cKeterangan.Text, GetOpt(optJenis), cJ1, cJ2, cJ3, cJ4, cJ5, cJ6, cJ7, _
          cJ8, cJ9, cJ10, cJ11, cJ12, cJ13, cJ14, cJ15)
  
  objData.Update GetDSN, "gagunan", "kode='" & cKode.Text & "'", vaField, vaValue
  
  initvalue
  GetEdit False
  GetSQL
End Sub

Private Sub Initdetail()
  cJudul.Default
  nLebar.Value = 0
End Sub

Private Sub initvalue()
Dim n As Integer

  cKeterangan.Default
  cKode.Default
  optJenis(0).Value = True
  vaArray.Clear
  vaArray.ReDim 0, 14, 0, 2
  
  For n = 0 To vaArray.UpperBound(1)
    vaArray(n, 0) = n + 1
  Next
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  Initdetail
  nNo.Value = 1
End Sub

Private Sub Form_Load()

  CenterForm Me, True
  initvalue
  GetSQL
  GetEdit False
  
  TabIndex cKode, n
  TabIndex cKeterangan, n
  TabIndex optJenis(0), n
  TabIndex optJenis(1), n
  TabIndex nNo, n
  TabIndex cJudul, n
  TabIndex nLebar, n
  TabIndex cmdOK, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  cKode.SetFocus
End Sub

Function Gabung(Teks1 As String, Teks2 As String) As String
  Dim S1 As Byte
  If Teks1 = "" And Teks2 = "" Then
    Gabung = ""
  Else
    S1 = 47 - Len(RTrim(Teks1))
    Gabung = RTrim(Teks1) + Replicate(" ", S1) + Trim(Teks2)
  End If
End Function

Private Sub cmdAdd_Click()
  GetEdit True
  initvalue
  cKode.SetFocus
  nPos = Add
End Sub

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  GetEdit True
  nPos = Edit
  cKode.SetFocus
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    initvalue
    GetEdit False
  End If
End Sub

Private Sub nNo_Validate(Cancel As Boolean)
  If vaArray.UpperBound(1) >= 0 And nNo.Value - 1 <= vaArray.UpperBound(1) Then
    cJudul.Text = vaArray(nNo.Value - 1, 1)
    nLebar.Value = vaArray(nNo.Value - 1, 2)
  End If
End Sub

Private Sub optJenis_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub TDBGrid2_Click()
  lClick = True
End Sub

Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Then
    lClick = True
  End If
End Sub

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If lClick And Not lEdit Then
    If Not dbData.eof Then
      cKode.Text = TDBGrid2.Columns(0)
      cKeterangan.Text = TDBGrid2.Columns(1)
      GetMemory
    End If
  End If
  lClick = False
End Sub

Private Sub GetSQL()
  Set dbData = objData.Browse(GetDSN, "Gagunan", , , , , , "Kode")
  Set TDBGrid2.DataSource = dbData
End Sub
