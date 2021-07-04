VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptBakiDebet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN BAKI DEBET"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11430
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2130
      Left            =   0
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   3757
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
      Begin VB.OptionButton optTampil 
         Caption         =   "&3 Yg Lunas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   7740
         TabIndex        =   14
         Top             =   165
         Width           =   2070
      End
      Begin VB.OptionButton optTampil 
         Caption         =   "&2 Yg Belum Lunas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   13
         Top             =   165
         Width           =   2070
      End
      Begin VB.OptionButton optTampil 
         Caption         =   "&1 Semua"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4545
         TabIndex        =   12
         Top             =   165
         Width           =   1155
      End
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Left            =   315
         TabIndex        =   0
         Top             =   150
         Width           =   3270
         _ExtentX        =   5768
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
         Caption         =   "SAMPAI TANGGAL"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGolongan 
         Height          =   330
         Index           =   0
         Left            =   3060
         TabIndex        =   1
         Top             =   900
         Width           =   4095
         _ExtentX        =   7223
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   315
         TabIndex        =   2
         Top             =   900
         Width           =   2745
         _ExtentX        =   4842
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
         Caption         =   "GOLONGAN"
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
      Begin BiSATextBoxProject.BiSABrowse cAO 
         Height          =   330
         Index           =   0
         Left            =   315
         TabIndex        =   3
         Top             =   1635
         Width           =   3030
         _ExtentX        =   5345
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
         Caption         =   "ANTARA AO"
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
      Begin BiSATextBoxProject.BiSABrowse cAO 
         Height          =   330
         Index           =   1
         Left            =   3405
         TabIndex        =   4
         Top             =   1650
         Width           =   1800
         _ExtentX        =   3175
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10245
         TabIndex        =   6
         Top             =   1650
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
         Picture         =   "RptBakiDebet.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   7875
         TabIndex        =   7
         Top             =   1650
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
         Picture         =   "RptBakiDebet.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdPrint 
         Height          =   435
         Left            =   9030
         TabIndex        =   8
         Top             =   1650
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   767
         Caption         =   "    &Cetak"
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
         Picture         =   "RptBakiDebet.frx":032C
      End
      Begin BiSATextBoxProject.BiSATextBox cNamaGolongan 
         Height          =   330
         Index           =   1
         Left            =   3060
         TabIndex        =   9
         Top             =   1245
         Width           =   4095
         _ExtentX        =   7223
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   315
         TabIndex        =   10
         Top             =   1245
         Width           =   2745
         _ExtentX        =   4842
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
         Caption         =   "SD. GOLONGAN"
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
      Begin BiSADateProject.BiSADate dTglRealisasi 
         Height          =   330
         Left            =   315
         TabIndex        =   11
         Top             =   510
         Width           =   3270
         _ExtentX        =   5768
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
         Caption         =   "TGL VALUTA"
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
         Caption         =   "Lihat"
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
         Left            =   3885
         TabIndex        =   15
         Top             =   195
         Width           =   540
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   5235
      Left            =   0
      Top             =   2115
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   9234
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
      Begin TrueOleDBGrid70.TDBGrid DataGrid 
         Height          =   5130
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   11310
         _ExtentX        =   19950
         _ExtentY        =   9049
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "AO"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA AO"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "REKENING"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TGL REALISASI"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "NAMA"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "LAMA"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "JTH TEMPO"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "PLAFOND"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "BAKI DEBET"
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=953"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2937"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2858"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2672"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2593"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2514"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2434"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3784"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3704"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=512"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1667"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1588"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2302"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2223"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2990"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2910"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=1958"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1879"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
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
         DataView        =   2
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000014&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H80000001&"
         _StyleDefs(13)  =   ":id=2,.fgcolor=&H8000000E&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(18)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=70,.parent=13,.alignment=0"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
         _StyleDefs(75)  =   "Named:id=33:Normal"
         _StyleDefs(76)  =   ":id=33,.parent=0"
         _StyleDefs(77)  =   "Named:id=34:Heading"
         _StyleDefs(78)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   ":id=34,.wraptext=-1"
         _StyleDefs(80)  =   "Named:id=35:Footing"
         _StyleDefs(81)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(82)  =   "Named:id=36:Selected"
         _StyleDefs(83)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(84)  =   "Named:id=37:Caption"
         _StyleDefs(85)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(86)  =   "Named:id=38:HighlightRow"
         _StyleDefs(87)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(88)  =   "Named:id=39:EvenRow"
         _StyleDefs(89)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(90)  =   "Named:id=40:OddRow"
         _StyleDefs(91)  =   ":id=40,.parent=33"
         _StyleDefs(92)  =   "Named:id=41:RecordSelector"
         _StyleDefs(93)  =   ":id=41,.parent=34"
         _StyleDefs(94)  =   "Named:id=42:FilterBar"
         _StyleDefs(95)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "RptBakiDebet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim lClick As Boolean

Private Sub cGolongan_ButtonClick(Index As Integer)
  Set dbData = objData.Pick(GetDSN, "GolonganKredit", "Kode", cGolongan(Index), "Kode,Keterangan")
  If Not dbData.eof Then
    cNamaGolongan(Index).Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL True
End Sub

Private Sub GetRpt()
    With FrmRPT
    .AddPageHeader "LAPORAN BAKI DEBET", tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader cGolongan(0).Text & " sd. " & cGolongan(1).Text, tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "SAMPAI TANGGAL  : " & Format(dDate.Value, "dd MMMM yyyy"), tdbHalignCenter, , , True, , 12, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableGroupHeader True, "[]", , , , 10
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "NO. REKENING", , , , 12
    .AddTableHeader "TGL REALIASI", , , , 10
    .AddTableHeader "NAMA NASABAH"
    .AddTableHeader "LAMA", , , , 6
    .AddTableHeader "JTH TMP", , , , 10
    .AddTableHeader "PLAFOND", , , , 15
    .AddTableHeader "BAKI DEBET", , , , 15
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody , tdbHalignCenter
    .AddTableBody
    .AddTableBody , tdbHalignCenter
    .AddTableBody , tdbHalignCenter
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "SUB TOTAL", , tdbHalignRight, , , , , , , , , , , , 5
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&sum", Sis_Rpt_Number2
    .AddTableGroupFooter "&sum", Sis_Rpt_Number2
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "GRAND TOTAL", , tdbHalignRight, , , , , , , , , , , , 5
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub

Private Sub GetSQL(ByVal lGrid As Boolean)
Dim vaJoin
Dim cWhere As String
Dim cField As String
Dim n As Integer
Dim nTotal As Double
Dim nPlafond As Double

  vaArray.ReDim 0, -1, 0, 8
  cField = "d.AO,a.Nama as NamaAO, d.Rekening,d.Tgl,d.Lama,d.JatuhTempo,d.Plafond,r.Nama"
  cWhere = "And d.AO >='" & cAO(0).Text & "'"
  cWhere = cWhere & " and d.AO <='" & cAO(1).Text & "'"
  
'  If optTampil(1).Value = True Then
'    cWhere = cWhere & " and d.Status <> '1'"
'  End If
'
'  If optTampil(2).Value = True Then
'    cWhere = cWhere & " and d.Status = '1'"
'  End If
  
  cWhere = cWhere & " and tgl <= '" & Format(dTglRealisasi.Value, "yyyy-MM-dd") & "' and statuspencairan = 1"
'  cWhere = cWhere & " and tgl <= '" & Format(dTglRealisasi.Value, "yyyy-MM-dd") & "'"
  vaJoin = Array("Left Join RegisterNasabah r On d.Kode = r.Kode", _
                 "Left Join AO a on a.Kode = d.AO")
  Set dbData = objData.Browse(GetDSN, "Debitur d", cField, "d.GolonganKredit", sisGTEqual, cGolongan(0).Text, " and d.GolonganKredit<='" & cGolongan(1).Text & "'" & cWhere, "d.AO,d.tgl,d.rekening", vaJoin)
  If Not dbData.eof Then
    dbData.MoveFirst
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull(dbData!AO, "")
      vaArray(n, 1) = GetNull(dbData!namaao, "")
      vaArray(n, 2) = GetNull(dbData!Rekening, "")
      vaArray(n, 3) = Format(GetNull(dbData!Tgl), "dd-MM-yyyy")
      vaArray(n, 4) = GetNull(dbData!nama, "")
      vaArray(n, 5) = GetNull(dbData!Lama, "")
      vaArray(n, 6) = Format(GetNull(dbData!JatuhTempo), "dd-MM-yyyy")
      vaArray(n, 7) = GetNull(dbData!plafond)
      vaArray(n, 8) = GetBK(vaArray(n, 2), vaArray(n, 7))
      
      nPlafond = nPlafond + vaArray(n, 7)
      nTotal = nTotal + vaArray(n, 8)
            
      If optTampil(1).Value = True Then
        If vaArray(n, 8) <= 0 Then
          vaArray.DeleteRows n
        End If
      ElseIf optTampil(2).Value = True Then
        If vaArray(n, 8) > 0 Then
          vaArray.DeleteRows n
        End If
      End If
          
      
      dbData.MoveNext
    Loop
        
    nPlafond = 0
    nTotal = 0
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      nPlafond = nPlafond + vaArray(n, 7)
      nTotal = nTotal + vaArray(n, 8)
    Next n
    
    DataGrid.Columns(8).FooterText = Format(nTotal, "###,###,###,##0.00")
    DataGrid.Columns(7).FooterText = Format(nPlafond, "###,###,###,##0.00")
    FrmPB.EndPB
    If lGrid = False Then
      GetRpt
    Else
      Set DataGrid.Array = vaArray
      DataGrid.ReBind
      DataGrid.Refresh
    End If
  Else
    MsgBox "Data Tidak Ada,..", vbInformation, Me.Caption
    ClearGrid
  End If
End Sub
Private Sub ClearGrid()
  vaArray.ReDim 0, -1, 0, 8
  Set DataGrid.Array = vaArray
  DataGrid.ReBind
  DataGrid.Refresh
End Sub

Private Function GetBK(ByVal cRek As String, ByVal nPlafond As Double) As Double
Dim dbBK As New ADODB.Recordset

  GetBK = nPlafond
  Set dbBK = objData.Browse(GetDSN, "Angsuran", "Sum(Pokok) as Pokok", "Rekening", sisAssign, cRek, "And Tgl <='" & Format(dDate.Value, "yyyy-mm-dd") & "' Group By Rekening", "Rekening")
  If Not dbBK.eof Then
    GetBK = nPlafond - GetNull(dbBK!pokok)
  End If
End Function

Private Sub cAO_ButtonClick(Index As Integer)
  Set dbData = objData.Pick(GetDSN, "AO", "Kode", cAO(Index), "Kode,Nama")
End Sub

Private Sub cAO_Validate(Index As Integer, Cancel As Boolean)
  If cAO(Index).LastKey = 13 Then
    cAO_ButtonClick (Index)
  End If
End Sub

Private Sub cmdPrint_Click()
  GetSQL False
End Sub

Private Sub Form_Load()
Dim n As Single

  optTampil(0).Value = True
  
  lClick = True
  CenterForm Me
  dDate.Value = Date
  dTglRealisasi.Value = Date
  GetMinMax "AO", cAO, "Kode"
    
  TabIndex dDate, n
  TabIndex dTglRealisasi, n
  
  TabIndex optTampil(0), n
  TabIndex optTampil(1), n
  TabIndex optTampil(2), n
  
  TabIndex cGolongan(0), n
  TabIndex cGolongan(1), n
  TabIndex cAO(0), n
  TabIndex cAO(1), n
  TabIndex cmdPreview, n
  TabIndex cmdPrint, n
  TabIndex cmdKeluar, n
  
End Sub

Private Sub DataGrid_HeadClick(ByVal ColIndex As Integer)
  If lClick Then
    Select Case ColIndex
      Case 0
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 0, XORDER_ASCEND, XTYPE_STRING
        lClick = Not lClick
      Case 1
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
        lClick = Not lClick
      Case 2
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
        lClick = Not lClick
      Case 3
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 3, XORDER_ASCEND, XTYPE_DATE
        lClick = Not lClick
      Case 4
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 4, XORDER_ASCEND, XTYPE_STRING
        lClick = Not lClick
      Case 5
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 5, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 6
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 6, XORDER_ASCEND, XTYPE_DATE
        lClick = Not lClick
      Case 7
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 7, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 8
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 8, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
    End Select
  Else
    Select Case ColIndex
      Case 0
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 0, XORDER_DESCEND, XTYPE_STRING
        lClick = Not lClick
      Case 1
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 1, XORDER_DESCEND, XTYPE_STRING
        lClick = Not lClick
      Case 2
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 2, XORDER_DESCEND, XTYPE_STRING
        lClick = Not lClick
      Case 3
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE
        lClick = Not lClick
      Case 4
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 4, XORDER_DESCEND, XTYPE_STRING
        lClick = Not lClick
      Case 5
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 5, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 6
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 6, XORDER_DESCEND, XTYPE_DATE
        lClick = Not lClick
      Case 7
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 7, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 8
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 8, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
    End Select
  End If
  DataGrid.ReBind
End Sub

Private Sub optTampil_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub
