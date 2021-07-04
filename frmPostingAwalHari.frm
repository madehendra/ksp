VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPostingAwalHari 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POSTING AWAL HARI"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   11505
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5610
      Left            =   0
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   4380
         Left            =   105
         TabIndex        =   3
         Top             =   1170
         Width           =   11310
         _ExtentX        =   19950
         _ExtentY        =   7726
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "BUNGA DEPOSITO"
         TabPicture(0)   =   "frmPostingAwalHari.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "BiSAFrame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "POKOK DEPOSITO"
         TabPicture(1)   =   "frmPostingAwalHari.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "BiSAFrame5"
         Tab(1).ControlCount=   1
         Begin BiSAFramProject.BiSAFrame BiSAFrame4 
            Height          =   3915
            Left            =   75
            Top             =   345
            Width           =   11160
            _ExtentX        =   19685
            _ExtentY        =   6906
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
               Height          =   3810
               Left            =   60
               TabIndex        =   6
               Top             =   60
               Width           =   11025
               _ExtentX        =   19447
               _ExtentY        =   6720
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "TGL VALUTA"
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "REKENING"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "NAMA"
               Columns(2).DataField=   ""
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "NOMINAL"
               Columns(3).DataField=   ""
               Columns(3).NumberFormat=   "###,###,###,###,##0.00"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "BUNGA"
               Columns(4).DataField=   ""
               Columns(4).NumberFormat=   "###,###,###,###,##0.00"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "PAJAK"
               Columns(5).DataField=   ""
               Columns(5).NumberFormat=   "###,###,###,###,##0.00"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).DataField=   ""
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).DataField=   ""
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(8)._VlistStyle=   0
               Columns(8)._MaxComboItems=   5
               Columns(8).DataField=   ""
               Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(9)._VlistStyle=   0
               Columns(9)._MaxComboItems=   5
               Columns(9).DataField=   ""
               Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(10)._VlistStyle=   0
               Columns(10)._MaxComboItems=   5
               Columns(10).Caption=   "REK SIMPANAN"
               Columns(10).DataField=   ""
               Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   11
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=11"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=2884"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2805"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=5768"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5689"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
               Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
               Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
               Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
               Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
               Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
               Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
               Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
               Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
               Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
               Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
               Splits(0)._ColumnProps(35)=   "Column(6).Visible=0"
               Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(37)=   "Column(7).Width=2725"
               Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=2646"
               Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=516"
               Splits(0)._ColumnProps(41)=   "Column(7).Visible=0"
               Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
               Splits(0)._ColumnProps(43)=   "Column(8).Width=2725"
               Splits(0)._ColumnProps(44)=   "Column(8).DividerColor=0"
               Splits(0)._ColumnProps(45)=   "Column(8)._WidthInPix=2646"
               Splits(0)._ColumnProps(46)=   "Column(8)._ColStyle=516"
               Splits(0)._ColumnProps(47)=   "Column(8).Visible=0"
               Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
               Splits(0)._ColumnProps(49)=   "Column(9).Width=2725"
               Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
               Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=2646"
               Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=516"
               Splits(0)._ColumnProps(53)=   "Column(9).Visible=0"
               Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
               Splits(0)._ColumnProps(55)=   "Column(10).Width=2725"
               Splits(0)._ColumnProps(56)=   "Column(10).DividerColor=0"
               Splits(0)._ColumnProps(57)=   "Column(10)._WidthInPix=2646"
               Splits(0)._ColumnProps(58)=   "Column(10)._ColStyle=516"
               Splits(0)._ColumnProps(59)=   "Column(10).Order=11"
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
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
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
               _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(26)  =   ":id=13,.strikethrough=0,.charset=0"
               _StyleDefs(27)  =   ":id=13,.fontname=Tahoma"
               _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(30)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(31)  =   ":id=14,.fontname=Tahoma"
               _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(35)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
               _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
               _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
               _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
               _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
               _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
               _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
               _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
               _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
               _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
               _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
               _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
               _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
               _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=78,.parent=13"
               _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
               _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
               _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
               _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
               _StyleDefs(70)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
               _StyleDefs(71)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
               _StyleDefs(72)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
               _StyleDefs(73)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
               _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
               _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
               _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
               _StyleDefs(77)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
               _StyleDefs(78)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
               _StyleDefs(79)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
               _StyleDefs(80)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
               _StyleDefs(81)  =   "Splits(0).Columns(10).Style:id=62,.parent=13"
               _StyleDefs(82)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=14"
               _StyleDefs(83)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=15"
               _StyleDefs(84)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=17"
               _StyleDefs(85)  =   "Named:id=33:Normal"
               _StyleDefs(86)  =   ":id=33,.parent=0"
               _StyleDefs(87)  =   "Named:id=34:Heading"
               _StyleDefs(88)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(89)  =   ":id=34,.wraptext=-1"
               _StyleDefs(90)  =   "Named:id=35:Footing"
               _StyleDefs(91)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(92)  =   "Named:id=36:Selected"
               _StyleDefs(93)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(94)  =   "Named:id=37:Caption"
               _StyleDefs(95)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(96)  =   "Named:id=38:HighlightRow"
               _StyleDefs(97)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(98)  =   "Named:id=39:EvenRow"
               _StyleDefs(99)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(100) =   "Named:id=40:OddRow"
               _StyleDefs(101) =   ":id=40,.parent=33"
               _StyleDefs(102) =   "Named:id=41:RecordSelector"
               _StyleDefs(103) =   ":id=41,.parent=34"
               _StyleDefs(104) =   "Named:id=42:FilterBar"
               _StyleDefs(105) =   ":id=42,.parent=33"
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame5 
            Height          =   3900
            Left            =   -74940
            Top             =   360
            Width           =   11160
            _ExtentX        =   19685
            _ExtentY        =   6879
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
            Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
               Height          =   3750
               Left            =   90
               TabIndex        =   7
               Top             =   60
               Width           =   11025
               _ExtentX        =   19447
               _ExtentY        =   6615
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "T. VALUTA"
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "REKENING"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "JTHTMP"
               Columns(2).DataField=   ""
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "NAMA"
               Columns(3).DataField=   ""
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "ALAMAT"
               Columns(4).DataField=   ""
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "NOMINAL"
               Columns(5).DataField=   ""
               Columns(5).NumberFormat=   "###,###,###,###,##0"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "REK SIMPANAN"
               Columns(6).DataField=   ""
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   7
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=7"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1826"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=2672"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2593"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=1905"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1826"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=4630"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=4551"
               Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
               Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(21)=   "Column(4).Width=5106"
               Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=5027"
               Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
               Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(26)=   "Column(5).Width=2858"
               Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2778"
               Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
               Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
               Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
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
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1.5
               FootLines       =   1
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
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
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
               _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(28)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(29)  =   ":id=14,.fontname=Tahoma"
               _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
               _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
               _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
               _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
               _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
               _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
               _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
               _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
               _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
               _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
               _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
               _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
               _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
               _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
               _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
               _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
               _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
               _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
               _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
               _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
               _StyleDefs(67)  =   "Named:id=33:Normal"
               _StyleDefs(68)  =   ":id=33,.parent=0"
               _StyleDefs(69)  =   "Named:id=34:Heading"
               _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(71)  =   ":id=34,.wraptext=-1"
               _StyleDefs(72)  =   "Named:id=35:Footing"
               _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(74)  =   "Named:id=36:Selected"
               _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(76)  =   "Named:id=37:Caption"
               _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(78)  =   "Named:id=38:HighlightRow"
               _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(80)  =   "Named:id=39:EvenRow"
               _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(82)  =   "Named:id=40:OddRow"
               _StyleDefs(83)  =   ":id=40,.parent=33"
               _StyleDefs(84)  =   "Named:id=41:RecordSelector"
               _StyleDefs(85)  =   ":id=41,.parent=34"
               _StyleDefs(86)  =   "Named:id=42:FilterBar"
               _StyleDefs(87)  =   ":id=42,.parent=33"
            End
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1020
         Left            =   105
         TabIndex        =   0
         Top             =   105
         Width           =   11310
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   675
      Left            =   0
      Top             =   5595
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   1191
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
         Left            =   2970
         TabIndex        =   8
         Top             =   195
         Width           =   2430
         _ExtentX        =   4286
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         Caption         =   "Tanggal"
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
         Height          =   465
         Left            =   10170
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   120
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   820
         Caption         =   " Keluar"
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
         Picture         =   "frmPostingAwalHari.frx":0038
      End
      Begin MSComctlLib.ProgressBar pr 
         Height          =   420
         Left            =   195
         TabIndex        =   2
         Top             =   165
         Visible         =   0   'False
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   741
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin BiSAButtonProject.BiSAButton cmdRefresh 
         Height          =   465
         Left            =   5505
         TabIndex        =   4
         ToolTipText     =   "Refresh"
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         Caption         =   "Posting"
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
         Picture         =   "frmPostingAwalHari.frx":00DE
      End
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   465
         Left            =   8835
         TabIndex        =   5
         ToolTipText     =   "Proses"
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
         Caption         =   "Simpan"
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
         Picture         =   "frmPostingAwalHari.frx":0288
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   465
         Left            =   6810
         TabIndex        =   9
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   820
         Caption         =   " Print"
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
         Picture         =   "frmPostingAwalHari.frx":06B3
      End
   End
End
Attribute VB_Name = "frmPostingAwalHari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaTabungan As New XArrayDB
Dim vaBungaDP As New XArrayDB
Dim vaPokokDP As New XArrayDB
Dim vaHasil As New XArrayDB

Private Sub InitPB(ByVal nMax)
  PR.Visible = True
  PR.Min = 0
  PR.Max = nMax + 1
  PR.Value = 0
End Sub

Private Sub RunPB()
  PR.Value = PR.Value + IIf(PR.Value < PR.Max, 1, 0)
End Sub

Private Sub EndPB()
  PR.Visible = False
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
Dim dDate As Date
  
  If vaBungaDP.UpperBound(1) < 0 And vaPokokDP.UpperBound(1) < 0 Then
    MsgBox "Tidak ada data yang disimpan...", vbOKOnly + vbInformation, "POSTING BUNGA / POKOK DEPOSITO"
    cmdKeluar.SetFocus
    Exit Sub
  End If
  
  If MsgBox("Data benar-benar sudah valid dan akan disimpan ?", vbOKCancel + vbInformation, "POSTING BUNGA / POKOK DEPOPSITO") = vbOK Then
    ProsesBungaDeposito
'    ProsesPokokDeposito
    MsgBox "Proses selesai...", vbOKOnly + vbInformation, "POSTING BUNGA / POKOK DEPOSITO"
    initvalue
    cmdKeluar.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cmdPreview_Click()
  With FrmRPT
    .AddPageHeader "POSTING BUNGA DEPOSITO", tdbHalignCenter, , , , dbArial, 12, True, True
    .AddPageHeader "TGL " & Format(dTgl.Value, "dd MMMM yy"), tdbHalignCenter, , , True, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "VALUTA", , , , 9
    .AddTableHeader "REKENING", , , , 10
    .AddTableHeader "NAMA"
    .AddTableHeader "NOMINAL", , , , 13
    .AddTableHeader "BUNGA", , , , 12
    .AddTableHeader "PAJAK", , , , 10
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    
    
    .AddTableFooter "TOTAL", , tdbHalignCenter, , , , , , , , , , , , 3
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    
    
    .Preview vaBungaDP
  End With
End Sub

Private Sub cmdRefresh_Click()
  PostingPokokDeposito
  PostingBungaDeposito
  SSTab1.Tab = 0
End Sub

Private Sub Form_Activate()
Dim dTanggal As Date
  dTanggal = Date
  If Not IsInPeriod(dTanggal) Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
Dim cText As String
    
  initvalue
  cText = "Posting Awal hari akan melakukan ha-hal sebagai berikut : " & vbCrLf
  cText = cText & "1. Pengeluaran Bunga Deposito ke rekening Titipan Bunga Deposito" & vbCrLf
  cText = cText & "2. Pengeluaran Pokok Deposito ke rekening Titipan Deposito Jatuh Tempo yang sudah jatuh tempo" & vbCrLf
  cText = cText & "3. Pengeluaran Pokok Deposito hanya yang Non ARO"
  Label1.Caption = cText
  CenterForm Me, True
  TabIndex cmdRefresh, n
  TabIndex cmdPreview, n
  TabIndex cmdOK, n
  TabIndex cmdKeluar, n
End Sub

Private Sub JatuhTempo(ByVal dValuta As Date, ByRef cDapatBunga As String, ByRef dLastBunga As Date)
Dim nDayBunga As Integer
Dim lStop As Boolean
Dim dTglBunga As Date
  
  lStop = True
  cDapatBunga = "0"
  dTglBunga = DateAdd("m", 1, dValuta)
  Do While lStop
    If Month(dTglBunga) = Month(dTgl.Value) Then
      If dTglBunga <= dTgl.Value Then
        dLastBunga = DateAdd("m", -1, dTglBunga)
        cDapatBunga = "1"
        lStop = False
      Else
        lStop = False
      End If
    End If
    dTglBunga = DateAdd("m", 1, dTglBunga)
  Loop
End Sub

Private Sub PostingPokokDeposito()
Dim cField As String
Dim cWhere As String
Dim vaJoin
Dim n As Integer
  
  vaPokokDP.ReDim 0, -1, 0, 7
  cField = "d.Rekening,d.Tgl,d.NominalDeposito,d.JthTmp,r.nama,r.Alamat,g.RekeningAkuntansi,g.Rekeningjatuhtempo"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode = d.Kode", _
                 "Left Join GolonganDeposito g on g.Kode=d.GolonganDeposito")
  cWhere = "And d.JthTmp >= '" & Format(dTgl.Value, "yyyy-mm-dd") & "'"
  cWhere = cWhere & "And d.Status <> '1'"
  cWhere = cWhere & "And d.SistemARO <> 'Y'"
  Set dbData = objData.Browse(GetDSN, "Deposito d", cField, "StatusPostingPokok", sisAssign, "0", cWhere, "d.rekening", vaJoin)
  If Not dbData.eof Then
    InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.eof
      RunPB
      vaPokokDP.InsertRows vaPokokDP.UpperBound(1) + 1
      n = vaPokokDP.UpperBound(1)
      vaPokokDP(n, 0) = GetNull(dbData!Tgl)
      vaPokokDP(n, 1) = GetNull(dbData!Rekening, "")
      vaPokokDP(n, 2) = GetNull(dbData!jthtmp)
      vaPokokDP(n, 3) = GetNull(dbData!nama, "")
      vaPokokDP(n, 4) = GetNull(dbData!alamat, "")
      vaPokokDP(n, 5) = GetNull(dbData!nominaldeposito)
      vaPokokDP(n, 6) = GetNull(dbData!RekeningAkuntansi, "")
      vaPokokDP(n, 7) = GetNull(dbData!RekeningJatuhtempo, "")
      dbData.MoveNext
    Loop
    EndPB
  End If
  Set TDBGrid3.Array = vaPokokDP
  TDBGrid3.ReBind
End Sub

Private Sub PostingBungaDeposito()
Dim n As Integer
Dim nJumlahHari As Integer
Dim dLastUpdate As Date
Dim nBunga As Double
Dim cField As String
Dim dTemp As Date
Dim trekening As String


  vaBungaDP.ReDim 0, -1, 0, 10
  cField = "d.Tgl,d.JthTmp,d.Rekening,d.NominalDeposito,d.SukuBunga,g.MinimumKenaPajak,g.PajakBunga, r.nama,d.rekeningsimpanan,"
  cField = cField & "g.RekeningBunga,g.CadanganBunga,g.RekeningPajakBunga"
  Set dbData = objData.Browse(GetDSN, "Deposito d", cField, "d.Status", sisDifference, "1", "And d.StatusPostingPokok <> '1'", "d.Tgl", _
                              Array("Left Join Registernasabah r on r.Kode = d.Kode", _
                                    "Left Join GolonganDeposito g on g.kode = d.GolonganDeposito"))
  If Not dbData.eof Then
    InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.eof
      RunPB
      trekening = GetNull(dbData!Rekening, "")
      If Format(DateAdd("m", 1, Format(GetNull(dbData!Tgl), "yyyy-MM-dd")), "yyyy-MM-dd") <= Format(dTgl.Value, "yyyy-MM-dd") And Format(dTgl.Value, "yyyy-MM-dd") <= Format(DateAdd("d", 7, dbData!jthtmp), "yyyy-MM-dd") Then
        vaBungaDP.InsertRows vaBungaDP.UpperBound(1) + 1
        n = vaBungaDP.UpperBound(1)
        vaBungaDP(n, 0) = GetNull(dbData!Tgl)
        vaBungaDP(n, 1) = GetNull(dbData!Rekening, "")
        vaBungaDP(n, 2) = GetNull(dbData!nama, "")
        vaBungaDP(n, 3) = GetNull(dbData!nominaldeposito)
        
        dLastUpdate = GetLastUpdateBungaDP(vaBungaDP(n, 1), vaBungaDP(n, 0))
        nJumlahHari = DateDiff("d", dLastUpdate, dTgl.Value)
        vaBungaDP(n, 4) = BungaDeposito(GetSukuBungaDeposito(vaBungaDP(n, 3), GetNull(dbData!SukuBunga)), vaBungaDP(n, 1), dTgl.Value) 'Round(vaBungaDP(n, 3) * nJumlahHari * GetNull(dbData!SukuBunga) / 100 / 360)
        
        If vaBungaDP(n, 3) >= GetNull(dbData!MinimumkenaPajak) Then
          vaBungaDP(n, 5) = Round(vaBungaDP(n, 4) * GetNull(dbData!pajakbunga) / 100)
        Else
          vaBungaDP(n, 5) = 0
        End If
        vaBungaDP(n, 6) = GetStatusBunga(dLastUpdate)
        vaBungaDP(n, 7) = GetNull(dbData!Rekeningbunga)
        vaBungaDP(n, 8) = GetNull(dbData!Cadanganbunga)
        vaBungaDP(n, 9) = GetNull(dbData!RekeningPajakbunga)
        vaBungaDP(n, 10) = GetNull(dbData!rekeningsimpanan)
      End If
      dbData.MoveNext
    Loop
    EndPB


    n = 0
    Do While n <= vaBungaDP.UpperBound(1)
      If Not GetValidPostingBungaDeposito(objData, vaBungaDP(n, 1)) Then
        vaBungaDP.DeleteRows n
        n = n - 1
      End If
      n = n + 1
    Loop
    
  End If
  Set TDBGrid2.Array = vaBungaDP
  TDBGrid2.ReBind
End Sub

Private Function BungaDeposito(nValue As Double, cRekening As String, dPosting As Date) As Double
Dim db As New ADODB.Recordset
Dim totalBunga As Double
Dim totalPencarian As Double
Dim n As Integer

  BungaDeposito = nValue
  
'  Set db = objData.Browse(GetDSN, "deposito", , "rekening", sisAssign, cRekening)
'  If Not db.eof Then
'    n = DateDiff("m", GetNull(db!Tgl), Format(dPosting, "yyyy-MM-dd"))
'    totalBunga = n * nValue
'  End If
'  totalPencarian = GetTotalPencairan(objData, cRekening) 'GetXPencairanBungaDeposito(objData, cRekening) * nValue
'  BungaDeposito = totalBunga - totalPencarian
End Function

Private Function GetTotalPencairan(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String) As Double
Dim db As New ADODB.Recordset
  
  GetTotalPencairan = 0
  Set db = objData.Browse(GetDSN, "mutasideposito", "sum(jumlah) as total, sum(pajak) as pajak", "rekening", sisAssign, cRekening, " and kodemutasi = '3'")
  If Not db.eof Then
    GetTotalPencairan = GetNull(db!Total) + GetNull(db!Pajak)
  End If
  'tambahkan juga dengan nilai yang belum dicairkan
  Set db = objData.Browse(GetDSN, "mutasibungadeposito", "sum(jumlah) as bunga, sum(pajak) as pajak", "rekening", sisAssign, cRekening)
  If Not db.eof Then
    GetTotalPencairan = GetTotalPencairan + GetNull(db!bunga) + GetNull(db!Pajak)
  End If
End Function

Private Function GetValidPerPanjangan(ByVal obj As CodeSuiteLibrary.data, ByVal Rekening As String) As Boolean
Dim db As New ADODB.Recordset

  GetValidPerPanjangan = True
  Set db = obj.Browse(GetDSN, "Deposito", "Tgl,JthTmp,LastPerpanjangan,Lama", "Rekening", sisAssign, Rekening)
  If Not db.eof Then
  End If
End Function

Private Function GetLastUpdateBungaDP(ByVal cRek As String, ByVal dValuta As Date) As Date
Dim dbTgl As New ADODB.Recordset

  Set dbTgl = objData.Browse(GetDSN, "BungaDeposito", , "Rekening", sisAssign, cRek, "group By ID", "ID")
  If Not dbTgl.eof Then
    dbTgl.MoveLast
    GetLastUpdateBungaDP = GetNull(dbTgl!Tgl)
  Else
    GetLastUpdateBungaDP = dValuta
  End If
End Function

Private Function GetValidPostingBungaDeposito(ByVal obj As CodeSuiteLibrary.data, ByVal Rekening As String) As Boolean
Dim db As New ADODB.Recordset
Dim nLama As Double
Dim cSQL As String
Dim dTanggal As Date
Dim cSistem As String
Dim dbPelunasan As New ADODB.Recordset

GetValidPostingBungaDeposito = True
cSQL = "select d.tgl, d.Lama, d.Rekening,d.SistemARO from deposito d"
cSQL = cSQL & " left join golongandeposito g on g.kode = d.Golongandeposito"
cSQL = cSQL & " where d.rekening = '" & Rekening & "'"
 
 nLama = 0
 dTanggal = Date
 Set db = obj.SQL(GetDSN, cSQL)
 If Not db.eof Then
  nLama = GetNull(db!Lama)
  dTanggal = GetNull(db!Tgl)
  cSistem = GetNull(db!SistemARO)
 End If
  
 'Cek dulu apakah deposito ini sudah lunas tutup/belum
 'Status = 0 = open
 'Status = 1 = close
 
 Set db = obj.Browse(GetDSN, "Deposito", "status", "Rekening", sisAssign, Rekening)
 If GetNull(db!status) = 1 Then
  GetValidPostingBungaDeposito = False
  Exit Function
 End If
 
 'Jika sudah lunas bunga jangan di posting
  Dim nJumlahPencairan As Integer
  nJumlahPencairan = 0
  nJumlahPencairan = GetXPencairanBungaDeposito(objData, Rekening)
  If nJumlahPencairan < GetLamaDepositoARO(Rekening) + nLama Then
    'cek dahulu apakah pemberian bunga ini jatuh pada bulan/thn valuta
    If Month(dTanggal) = Month(dTgl.Value) And Year(dTanggal) = Year(dTgl.Value) Then
      GetValidPostingBungaDeposito = False
      Exit Function
    End If
    'cek dalam pelunasan  apakah bulan ini sudah pernah
    'dilunasi? jika ya maka stop, jgn melakukan posting
    
'    Set db = obj.Browse(GetDSN, "MutasiBungaDeposito", "Rekening,Tgl", "Rekening", sisAssign, Rekening, " and month(tgl) ='" & Month(dTgl.Value) & "' and year(tgl)='" & Year(dTgl.Value) & "'")
'    If Not db.eof Then
'      GetValidPostingBungaDeposito = False
'      Exit Function
'    End If
    
    
    Set db = obj.Browse(GetDSN, "bungadeposito", , "rekening", sisAssign, Rekening, " and month(tgl) ='" & Month(dTgl.Value) & "' and year(tgl)='" & Year(dTgl.Value) & "'")
    If Not db.eof Then
      'cek apakah dalam bulan ini sudah dilunasi/belum
      GetValidPostingBungaDeposito = False
      Exit Function
    End If
    
    'cek dalam pelunasan  apakah bulan ini sudah pernah
    'dilunasi? jika ya maka stop, jgn melakukan posting
'    Set db = obj.Browse(GetDSN, "MutasiDeposito", "Rekening,Tgl", "Rekening", sisAssign, Rekening, " and month(tgl) ='" & Month(dTgl.Value) & "' and year(tgl)='" & Year(dTgl.Value) & "' and KodeMutasi = '3'")
'    If Not db.eof Then
'      GetValidPostingBungaDeposito = False
'      Exit Function
'    End If
    
    'cek apakah pada tgl ini memang jatuh tempo atau tidak?
'    If Day(dTanggal) <= Day(dTgl.Value) Then
'      GetValidPostingBungaDeposito = True
'    Else
'      GetValidPostingBungaDeposito = False
'    End If
    
    If Day(dTanggal) <= Day(dTgl.Value) Then
'      GetValidPostingBungaDeposito = True
    Else
      If Day(dTanggal) = 29 And Month(dTgl.Value) = 2 And Day(dTgl.Value) = 28 Then
        If Month(DateAdd("d", 1, dTgl.Value)) = 3 Then
          GetValidPostingBungaDeposito = True
        End If
      Else
        GetValidPostingBungaDeposito = False
      End If
    End If
    
  ElseIf nJumlahPencairan <= 0 Then
    ' jika belum sama sekali mendapat bunga deposito
    If Month(dTgl.Value) = Month(dTanggal) And Year(dTgl.Value) = Year(dTanggal) Then
      GetValidPostingBungaDeposito = False
      Exit Function
    End If
  ElseIf nJumlahPencairan >= GetLamaDepositoARO(Rekening) + nLama Then
    'STOP
    'cek apakah deposito ini aro atau tidak?
    If cSistem = "Y" Then
      Set db = obj.Browse(GetDSN, "Deposito d", "d.Rekening,d.Tgl,d.JthTmp,d.LastPerpanjangan,g.Lama", "d.Rekening", sisAssign, Rekening, , , Array("Left Join GolonganDeposito g on g.Kode = d.GolonganDeposito"))
      'cek dulu apakah bulan ini adalah bulan perpanjangan
      If Not db.eof Then
        If IsDalamPeriode(dTgl.Value, GetNull(db!Tgl), GetNull(db!jthtmp), GetNull(db!Lama)) Then
          'cek apakah dalam bulan ini sudah pernah pencairan
          Set dbPelunasan = obj.Browse(GetDSN, "MutasiBungaDeposito", "Rekening,Tgl", "Rekening", sisAssign, Rekening, " and Month(Tgl)=" & Month(dTgl.Value) & " and Year(Tgl)=" & Year(dTgl.Value) & "")
          If Not dbPelunasan.eof Then
            GetValidPostingBungaDeposito = False
            Exit Function
          End If
        End If
        If IsDalamPeriode(dTgl.Value, GetNull(db!LastPerpanjangan), GetNull(db!jthtmp), GetNull(db!Lama)) Then
          'cek apakah dalam bulan ini sudah pernah pencairan
          Set dbPelunasan = obj.Browse(GetDSN, "MutasiBungaDeposito", "Rekening,Tgl", "Rekening", sisAssign, Rekening, " and Month(Tgl)=" & Month(dTgl.Value) & " and Year(Tgl)=" & Year(dTgl.Value) & "")
          If Not dbPelunasan.eof Then
            GetValidPostingBungaDeposito = False
            Exit Function
          End If
        End If
      End If
      GetValidPostingBungaDeposito = False
    Else
      GetValidPostingBungaDeposito = False
    End If
    Exit Function
  Else
    GetValidPostingBungaDeposito = False
    Exit Function
  End If
End Function

Private Function GetLamaDepositoARO(ByVal Rekening As String) As Integer
Dim dbData As New ADODB.Recordset

  GetLamaDepositoARO = 0
  Set dbData = objData.Browse(GetDSN, "deposito", "jumlahperpanjangan,lama", "rekening", sisAssign, Rekening, " and sistemaro = 'Y'")
  If Not dbData.eof Then
    GetLamaDepositoARO = Val(GetNull(dbData!jumlahperpanjangan)) * Val(GetNull(dbData!Lama))
  End If
End Function

Private Function GetXPencairanBungaDeposito(ByVal obj As CodeSuiteLibrary.data, ByVal Rekening As String) As Integer
Dim db As New ADODB.Recordset
Dim nTmp As Integer

  nTmp = 0
'  Set db = obj.Browse(GetDSN, "MutasiBungaDeposito", "count(rekening) as count", "Rekening", sisAssign, Rekening)
'  If Not db.eof Then
'    nTmp = GetNull(db!Count)
'  End If
  Set db = obj.Browse(GetDSN, "MutasiDeposito", "count(rekening) as count", "Rekening", sisAssign, Rekening, " and KodeMutasi = '3'")
  If Not db.eof Then
    nTmp = nTmp + GetNull(db!Count)
  End If
  GetXPencairanBungaDeposito = nTmp
End Function

Private Sub ProsesBungaDeposito()
Dim n As Integer
Dim cFaktur As String
Dim nUrut As Double
Dim db As New ADODB.Recordset
Dim db1 As New ADODB.Recordset

  '==========================================================
  'No urut faktur diambil dari table BUNGA DEPOSITO
  'dengan satu nomor faktur untuk sekali posting ini, data kemudian disimpan di table:
  '
  'BUKUBESAR
  'BUNGADEPOSITO
  'MUTASIBUNGADEPOSITO
  '==========================================================
  'Penjelasan fungsi dari masing masing table
  '
  'Table BUNGADEPOSITO
  '===================
  'table ini digunakan untuk mencatat pemberian bunga deposito untuk
  'deposan. Record dalam table boleh dihapus jika dan hanya jika terjadi pembatalan
  'proses posting.
  '
  'Table MUTASIBUNGADEPOSITO
  '=========================
  'table ini digunakan untuk mencatat (mengcopy) data bunga deposito dari hasil posting.
  'Record dalam table ini terhapus jika deposan sudah mencairkannya (secara cash atau pemindahbukuan)
  '
  'Pembatalan posting awal hari (bunga Deposito) hanya berpengaruh pada
  '1. Buku Besar Account Titipan Bunga Deposito
  '2. Table Bunga Deposito dan MutasiBungaDeposito yang masih aktif.
  '
  
  nUrut = 100
  If GetFakturBunga >= nUrut Then
    nUrut = GetFakturBunga + 1
  End If
  
  InitPB vaBungaDP.UpperBound(1) + 1
  If vaBungaDP.UpperBound(1) >= 0 Then
    cFaktur = "BUNGADP" & Padl(Trim(Str(nUrut)), 13, "0")
    objData.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msDeposito, "And Faktur='" & cFaktur & "'"
    objData.Delete GetDSN, "BungaDeposito", "Faktur", sisAssign, cFaktur
    objData.Delete GetDSN, "MutasiBungaDeposito", "faktur", sisAssign, cFaktur
    objData.Delete GetDSN, "MutasiTabungan", "faktur", sisAssign, cFaktur
    
    For n = 0 To vaBungaDP.UpperBound(1)
      RunPB
      objData.Add GetDSN, "BungaDeposito", Array("Faktur", "tgl", "Rekening", "Bunga", "Pajak"), _
                                           Array(cFaktur, dTgl.Value, vaBungaDP(n, 1), vaBungaDP(n, 4), vaBungaDP(n, 5))
      
      'Jika ybs punya rekening simpanan. Bunga supaya diposting langsung ke simpanan.
      'cek apakah rekening ini valid
      'jika valid maka simpan langsung ke rekening
      Set db = objData.Browse(GetDSN, "tabungan", "rekening,golongantabungan", "rekening", sisAssign, vaBungaDP(n, 10))
      If Not db.eof Then
        
        Set db1 = objData.Browse(GetDSN, "golongantabungan", , "kode", sisAssign, GetNull(db!GolonganTabungan))
        If Not db1.eof Then
          
          'posting bunga ke neraca
          UpdKodeTr objData, msDeposito, aCfg(msKodeCabang), cFaktur, dTgl.Value, vaBungaDP(n, 7), "Biaya Bunga Deposito", vaBungaDP(n, 4), 0, "N", SNow
              UpdKodeTr objData, msDeposito, aCfg(msKodeCabang), cFaktur, dTgl.Value, GetNull(db1!Rekening), "Bunga Deposito", 0, vaBungaDP(n, 4) - vaBungaDP(n, 5), "N", SNow
              UpdKodeTr objData, msDeposito, aCfg(msKodeCabang), cFaktur, dTgl.Value, vaBungaDP(n, 9), "Hutang Pajak Bunga Deposito", 0, vaBungaDP(n, 5), "N", SNow
            
          'masukkan ke dalam mutasi tabungan
'          UpdMutasiTabungan objData, aCfg(msKodeTransaksiPB), cFaktur, dTgl.Value, vaBungaDP(n, 10), vaBungaDP(n, 4) - vaBungaDP(n, 5), True, "Pencairan Bunga Deposito ke Tabungan", False
           UpdMutasiTabungan objData, aCfg(msKodeTransaksiPB), cFaktur, dTgl.Value, vaBungaDP(n, 10), vaBungaDP(n, 4) - vaBungaDP(n, 5), False, "Pencairan Bunga Deposito ke Tabungan", False
         End If
        Else
          'bunga(5)
          '   hutang titipan bunga(2)
          '   hutang pajak bunga deposito(2)
          
          UpdKodeTr objData, msDeposito, aCfg(msKodeCabang), cFaktur, dTgl.Value, vaBungaDP(n, 7), "Biaya Bunga Deposito", vaBungaDP(n, 4), 0, "N", SNow
              UpdKodeTr objData, msDeposito, aCfg(msKodeCabang), cFaktur, dTgl.Value, vaBungaDP(n, 8), "Titipan Bunga Deposito", 0, vaBungaDP(n, 4) - vaBungaDP(n, 5), "N", SNow
              UpdKodeTr objData, msDeposito, aCfg(msKodeCabang), cFaktur, dTgl.Value, vaBungaDP(n, 9), "Hutang Pajak Bunga Deposito", 0, vaBungaDP(n, 5), "N", SNow
              
          objData.Add GetDSN, "MutasiBungaDeposito", _
            Array("Faktur", "Rekening", "Tgl", "Jumlah", "Pajak", "UserName", "DateTime"), _
            Array(cFaktur, vaBungaDP(n, 1), Format(dTgl.Value, "yyyy-MM-dd"), vaBungaDP(n, 4), vaBungaDP(n, 5), GetRegistry(reg_UserName), SNow)
      End If
      
    Next
    EndPB
  End If
End Sub

Private Sub ProsesPokokDeposito()
Dim n As Integer
Dim nUrut As Double
Dim cFaktur As String
    
  nUrut = GetFakturPokok + 1
  InitPB vaPokokDP.UpperBound(1) + 1
  For n = 0 To vaPokokDP.UpperBound(1)
    RunPB
    cFaktur = "POKOKDP" & Padl(Trim(Str(nUrut)), 13, "0")
    objData.Edit GetDSN, "Deposito", "Rekening='" & vaPokokDP(n, 1) & "'", Array("StatusPostingPokok"), Array("1")
    objData.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msDeposito, "And Faktur='" & cFaktur & "'"
    UpdKodeTr objData, msDeposito, aCfg(msKodeCabang), cFaktur, dTgl.Value, vaPokokDP(n, 6), "Titipan Pokok Deposito", vaPokokDP(n, 5), 0, "N", SNow
        UpdKodeTr objData, msDeposito, aCfg(msKodeCabang), cFaktur, dTgl.Value, vaPokokDP(n, 7), "Titipan Pokok Deposito", 0, vaPokokDP(n, 5), "N", SNow
    nUrut = nUrut + 1
  Next
  EndPB
End Sub

Private Function GetFakturPokok() As Double
  GetFakturPokok = 0
  Set dbData = objData.Browse(GetDSN, "BukuBesar", "Max(Faktur) as faktur", "Faktur", sisPrefix, "POKOKDP")
  If Not dbData.eof Then
    GetFakturPokok = Val(Mid(GetNull(dbData!Faktur), 8))
  End If
End Function

Private Function GetFakturBunga() As Double
  GetFakturBunga = 0
  Set dbData = objData.Browse(GetDSN, "BungaDeposito", "Max(Faktur) as faktur", "Faktur", sisPrefix, "BUNGADP")
  If Not dbData.eof Then
    GetFakturBunga = Val(Mid(GetNull(dbData!Faktur), 8))
  End If
End Function

Private Function GetStatusBunga(ByVal dLastUpdate) As String
Dim dTglBerikut As Date

  GetStatusBunga = "0"
  dTglBerikut = DateAdd("m", 1, dLastUpdate)
  If dTgl.Value >= dTglBerikut Then
    GetStatusBunga = "1"
  End If
End Function

Private Sub initvalue()
  vaBungaDP.Clear
  vaBungaDP.ReDim 0, -1, 0, 10
  Set TDBGrid2.Array = vaBungaDP
  TDBGrid2.ReBind
  
  vaPokokDP.Clear
  vaPokokDP.ReDim 0, -1, 0, 7
  Set TDBGrid3.Array = vaPokokDP
  TDBGrid3.ReBind
End Sub
