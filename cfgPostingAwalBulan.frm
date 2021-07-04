VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form cfgPostingAwalBulan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posting Bunga Harian (Proses Akhir Bulan)"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10635
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1335
      Left            =   0
      Top             =   15
      Width           =   10650
      _ExtentX        =   18785
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
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   315
         Width           =   1065
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   315
         Width           =   870
      End
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   705
         Width           =   2625
         _ExtentX        =   4630
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
         Caption         =   "Awal Bulan"
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
         Left            =   2850
         TabIndex        =   1
         Top             =   705
         Width           =   2685
         _ExtentX        =   4736
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
         Caption         =   "Akhir Bulan"
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
      Begin BiSATextBoxProject.BiSABrowse cKodeTransaksi 
         Height          =   330
         Left            =   6435
         TabIndex        =   9
         Top             =   315
         Width           =   1980
         _ExtentX        =   3493
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
         Caption         =   "K.Trans"
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
      Begin BiSATextBoxProject.BiSATextBox cRekTransaksi 
         Height          =   330
         Left            =   8445
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   315
         Width           =   1905
         _ExtentX        =   3360
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
      Begin BiSATextBoxProject.BiSATextBox cDK 
         Height          =   330
         Left            =   8445
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   675
         Width           =   945
         _ExtentX        =   1667
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
      Begin BiSATextBoxProject.BiSATextBox cKas 
         Height          =   330
         Left            =   9420
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   675
         Width           =   930
         _ExtentX        =   1640
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
      Begin VB.Label Label1 
         Caption         =   "PERIODE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   360
         Width           =   825
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   5610
      Left            =   0
      Top             =   1320
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   9895
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
         Height          =   5535
         Left            =   45
         TabIndex        =   3
         Top             =   30
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   9763
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Rekening"
         Columns(0).DataField=   "UserName"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Jumlah Bunga"
         Columns(2).DataField=   "Plafond"
         Columns(2).NumberFormat=   "###,###,###,###,##0.00"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3731"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3651"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=10636"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=10557"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3149"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3069"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483633
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
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HC0C0C0&"
         _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000012&,.appearance=0,.ellipsis=0,.borderColor=&HFF8000&"
         _StyleDefs(11)  =   ":id=4,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
         _StyleDefs(14)  =   ":id=2,.fgcolor=&H0&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(15)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(17)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(18)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(19)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(20)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(22)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(23)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(24)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H8000000F&"
         _StyleDefs(25)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(26)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(27)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(28)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(29)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(30)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   6900
      Width           =   10635
      _ExtentX        =   18759
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   435
         Left            =   6540
         TabIndex        =   14
         Top             =   90
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   767
         Caption         =   "Print"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   8415
         TabIndex        =   4
         Top             =   90
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
         Picture         =   "cfgPostingAwalBulan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9495
         TabIndex        =   5
         Top             =   90
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
         Picture         =   "cfgPostingAwalBulan.frx":0416
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   7260
         TabIndex        =   6
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Posting"
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
         Picture         =   "cfgPostingAwalBulan.frx":04BC
      End
      Begin VB.Label Label2 
         Caption         =   "Perhatian: Hanya saldo dengan perolehan diatas 100 rupiah yang akan di posting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   210
         TabIndex        =   13
         Top             =   135
         Width           =   6165
      End
   End
End
Attribute VB_Name = "cfgPostingAwalBulan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
  If vaArray.UpperBound(1) > 0 Then
    If MsgBox("Apakah akan mencetak daftar simpanan harian?", vbYesNo) = vbYes Then
      GetPreview
    End If
  End If
End Sub

Private Sub cKodeTransaksi_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "kodetransaksi")
  If Not db.eof Then
    cKodeTransaksi.Text = cKodeTransaksi.Browse(db)
    cKodeTransaksi.Text = GetNull(db!Kode)
    cRekTransaksi.Text = GetNull(db!Rekening)
    cDK.Text = GetNull(db!DK)
    cKas.Text = GetNull(db!Kas)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim n As Integer
Dim nTotal As Double
Dim db As New ADODB.Recordset

  nTotal = 0
  
  
  vaArray.ReDim 0, -1, 0, 6
  'ambil seluruh data tabungan yg menggunakan perhitungan bunga harian
'  Set dbData = objData.Browse(GetDSN, "tabungan t", "t.rekening,r.nama,g.rekening as rekeningakun,g.rekeningbunga,t.pdl,p.keterangan as namapdl", "g.jenisbunga", sisAssign, 3, , "t.pdl,t.rekening", Array("left join golongantabungan g on g.kode = t.golongantabungan", "left join registernasabah r on r.kode = t.kode", "left join pdl p on p.kode = t.pdl"))
  Set dbData = objData.Browse(GetDSN, "bungaharian b", "sum(b.bunga) as bunga,t.rekening,r.nama,g.rekening as rekeningakun,g.rekeningbunga,t.pdl,p.keterangan as namapdl", "b.tgl", sisGTEqual, Format(dTgl(0).Value, "yyyy-MM-dd"), " and b.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "' GROUP by b.rekeningsimpanan", "p.kode,b.rekeningsimpanan", Array("left join tabungan t on t.rekening = b.rekeningsimpanan", "left join registernasabah r on r.kode = t.kode", "left join pdl p on p.kode = t.pdl", "left join golongantabungan g on g.kode = t.golongantabungan"))
  FrmPB.InitPB dbData.RecordCount
  If Not dbData.eof Then
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Rekening)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!bunga) 'GetSaldoBungaBulanan(vaArray(n, 0), dTgl(0).Value, dTgl(1).Value)
      vaArray(n, 3) = GetNull(dbData!rekeningakun)
      vaArray(n, 4) = GetNull(dbData!Rekeningbunga)
      vaArray(n, 5) = GetNull(dbData!PDL)
      vaArray(n, 6) = GetNull(dbData!namapdl)
      nTotal = nTotal + vaArray(n, 2)
      If vaArray(n, 2) < 100 Then
        vaArray.DeleteRows n
      End If
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  
  
  Set DataGrid1.Array = vaArray
  DataGrid1.Columns(2).FooterText = Format(nTotal, "###,###,##0.00")
  
  DataGrid1.Update
  DataGrid1.Refresh
  DataGrid1.ReBind
'  If vaArray.UpperBound(1) > 0 Then
'    If MsgBox("Apakah akan mencetak daftar simpanan harian?", vbYesNo) = vbYes Then
'      GetPreview
'    End If
'  End If
End Sub

Private Sub GetPreview()
  With FrmRPT
    .AddPageHeader "HASIL PERHITUNGAN BUNGA SIMPANAN HARIAN", tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader aCfg(msNama), tdbHalignCenter, , , True
    .AddPageHeader "Periode Bulan/Tahun : " & Combo1.Text & "/" & Combo2.Text, tdbHalignCenter, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
   
    .AddTableHeader "REKENING", , , , 14
    .AddTableHeader "NAMA"
    .AddTableHeader "BUNGA", , , , 14
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "PDL", , , , 8
    .AddTableHeader "NAMA", , , , 29
    
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter
    .AddTableFooter
    
    .Preview vaArray, True
  End With
End Sub

Private Function GetSaldoBungaBulanan(ByVal Rekening As String, ByVal dAwalPeriode As Date, ByVal dAkhirPeriode As Date) As Double
GetSaldoBungaBulanan = 0
Dim db As New ADODB.Recordset
  
  GetSaldoBungaBulanan = 0
  Set db = objData.Browse(GetDSN, "bungaharian", "sum(bunga) as bunga", "tgl", sisGT, Format(dAwalPeriode, "yyyy-MM-dd"), " and tgl <= '" & Format(dAkhirPeriode, "yyyy-MM-dd") & "' and rekeningsimpanan = '" & Rekening & "'")
  If Not db.eof Then
    GetSaldoBungaBulanan = GetNull(db!bunga)
  End If
  
  'ambil seluruh tabungan yg perhitungan bunga simpanannya menggunakan harian
  'ambil saldo bunga harian pada tabel bunga harian
  'simpan ditabel simpananharianbulanan
  'posting ke bukubesar
  'posting keseluruh buku/kartu simpanan masing2 tabungan
End Function

Private Sub cmdSimpan_Click()
Dim cKodeCabang As String
Dim n As Integer
Dim nJumlah As Double
Dim cFaktur As String
  
  cFaktur = Combo1.Text & "-" & Combo2.Text
  nJumlah = 0
  cKodeCabang = aCfg(msKodeCabang)
  objData.Delete GetDSN, "postingakhirbulansimpananharian", "periode", sisAssign, Combo1.Text, " and tahun = '" & Combo2.Text & "'"
  DelKodeTr objData, msTabungan, cKodeCabang, cFaktur
  'simpan di table postingakhirbulansimpananharian
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    nJumlah = nJumlah + vaArray(n, 2)
    objData.Add GetDSN, "postingakhirbulansimpananharian", Array("faktur", "rekeningsimpanan", "periode", "tahun", "tglawal", "tglakhir", "jumlah", "username", "datetime"), Array(cFaktur, vaArray(n, 0), Combo1.Text, Combo2.Text, Format(dTgl(0).Value, "yyyy-MM-dd"), Format(dTgl(1).Value, "yyyy-MM-dd"), vaArray(n, 2), GetRegistry(reg_UserName), SNow)
    UpdMutasiTabungan objData, cKodeTransaksi.Text, cFaktur, Format(Date, "yyyy-MM-dd"), vaArray(n, 0), vaArray(n, 2), False, "Bunga Simpanan Harian an. " & vaArray(n, 1), False, cDK.Text, cRekTransaksi.Text
    
    'biaya
    ' hutang
    UpdKodeTr objData, msTabungan, cKodeCabang, cFaktur, Format(Date, "yyyy-MM-dd"), vaArray(n, 4), "Posting Bunga Harian Bulanan Periode " & cFaktur, vaArray(n, 2), , "N"
      UpdKodeTr objData, msTabungan, cKodeCabang, cFaktur, Format(Date, "yyyy-MM-dd"), vaArray(n, 3), "Posting Bunga Harian Bulanan Periode " & cFaktur, , vaArray(n, 2), "N"

  Next n
  vaArray.ReDim 0, -1, 0, 6
  Set DataGrid1.Array = vaArray
  DataGrid1.ReBind
  DataGrid1.Refresh
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
  GetCombo
End Sub

Private Sub Form_Load()
Dim n As Integer

  CenterForm Me
  For n = 1 To 12
    Combo1.AddItem n
  Next n
  Combo2.AddItem Year(Date)
  Combo1.Text = Month(Date)
  Combo2.Text = Year(Date)
  GetCombo
End Sub

Private Sub GetCombo()
  dTgl(0).Value = BOM(DateSerial(Combo2.Text, Combo1.Text, 1))
  dTgl(1).Value = EOM(DateSerial(Combo2.Text, Combo1.Text, 1))
End Sub
