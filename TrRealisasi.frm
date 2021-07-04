VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form TrRealisasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI REALISASI PINJAMAN"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11805
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   5940
      Left            =   0
      Top             =   15
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10478
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
         Height          =   5820
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   10266
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabMaxWidth     =   2646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "DATA NASABAH"
         TabPicture(0)   =   "TrRealisasi.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "BiSAFrame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "JAMINAN"
         TabPicture(1)   =   "TrRealisasi.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "BiSAFrame6"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "REALISASI"
         TabPicture(2)   =   "TrRealisasi.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "BiSAFrame7"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "JADWAL ANGSURAN"
         TabPicture(3)   =   "TrRealisasi.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "BiSAFrame4"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "DATA ANALISA KEUANGAN"
         TabPicture(4)   =   "TrRealisasi.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "BiSAFrame5"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin BiSAFramProject.BiSAFrame BiSAFrame4 
            Height          =   4770
            Left            =   -74925
            Top             =   375
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   8414
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
               Height          =   4635
               Left            =   75
               TabIndex        =   83
               Top             =   75
               Width           =   11310
               _ExtentX        =   19950
               _ExtentY        =   8176
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "KE"
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "JATUH TEMPO"
               Columns(1).FooterText=   "Jumlah"
               Columns(1).DataField=   ""
               Columns(1).NumberFormat=   "dd-MM-yyyy"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "BUNGA"
               Columns(2).DataField=   ""
               Columns(2).NumberFormat=   "FormatText Event"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "POKOK"
               Columns(3).DataField=   ""
               Columns(3).NumberFormat=   "FormatText Event"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "ANGSURAN"
               Columns(4).DataField=   ""
               Columns(4).NumberFormat=   "FormatText Event"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "SISA BUNGA"
               Columns(5).DataField=   ""
               Columns(5).NumberFormat=   "FormatText Event"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "SISA POKOK"
               Columns(6).DataField=   ""
               Columns(6).NumberFormat=   "FormatText Event"
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1217"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1138"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=2540"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2461"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=3254"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
               Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
               Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(21)=   "Column(4).Width=3043"
               Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2963"
               Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
               Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(26)=   "Column(5).Width=2963"
               Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2884"
               Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
               Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(31)=   "Column(6).Width=3731"
               Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=3651"
               Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
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
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=104,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFCFCED&,.fgcolor=&H80000008&"
               _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
               _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
               _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HEBDACB&"
               _StyleDefs(15)  =   ":id=3,.fgcolor=&H80000008&,.bold=0,.fontsize=825,.italic=0,.underline=0"
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
               _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(27)  =   ":id=13,.strikethrough=0,.charset=0"
               _StyleDefs(28)  =   ":id=13,.fontname=Tahoma"
               _StyleDefs(29)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(30)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(31)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(32)  =   ":id=14,.fontname=Tahoma"
               _StyleDefs(33)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(34)  =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(35)  =   ":id=15,.fontname=Tahoma"
               _StyleDefs(36)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(37)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(38)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(39)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(40)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(41)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(42)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(43)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
               _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
               _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
               _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
               _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
               _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
               _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
               _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
               _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
               _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
               _StyleDefs(65)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
               _StyleDefs(66)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
               _StyleDefs(67)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
               _StyleDefs(68)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
               _StyleDefs(69)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
               _StyleDefs(70)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
               _StyleDefs(71)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
               _StyleDefs(72)  =   "Named:id=33:Normal"
               _StyleDefs(73)  =   ":id=33,.parent=0"
               _StyleDefs(74)  =   "Named:id=34:Heading"
               _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(76)  =   ":id=34,.wraptext=-1"
               _StyleDefs(77)  =   "Named:id=35:Footing"
               _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(79)  =   "Named:id=36:Selected"
               _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(81)  =   "Named:id=37:Caption"
               _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(83)  =   "Named:id=38:HighlightRow"
               _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(85)  =   "Named:id=39:EvenRow"
               _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(87)  =   "Named:id=40:OddRow"
               _StyleDefs(88)  =   ":id=40,.parent=33"
               _StyleDefs(89)  =   "Named:id=41:RecordSelector"
               _StyleDefs(90)  =   ":id=41,.parent=34"
               _StyleDefs(91)  =   "Named:id=42:FilterBar"
               _StyleDefs(92)  =   ":id=42,.parent=33"
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame7 
            Height          =   4815
            Left            =   -74970
            Top             =   345
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   8493
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
            Begin BiSANumberBoxProject.BiSANumberBox nPersBunga 
               Height          =   345
               Left            =   2685
               TabIndex        =   78
               Top             =   4080
               Visible         =   0   'False
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   609
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
            Begin BiSANumberBoxProject.BiSANumberBox nJumlahADM 
               Height          =   330
               Left            =   8310
               TabIndex        =   72
               Top             =   1305
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   582
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
               Caption         =   "% x Plafond ="
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
            Begin BiSANumberBoxProject.BiSANumberBox nPeriode 
               Height          =   330
               Left            =   420
               TabIndex        =   66
               Top             =   2460
               Width           =   3315
               _ExtentX        =   5847
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
               Caption         =   "Max Bayar (Dlm 1 Bulan)"
               CaptionWidth    =   2500
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
            Begin VB.OptionButton optCara 
               Caption         =   "&Bulanan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   3000
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   1680
               Width           =   990
            End
            Begin VB.OptionButton optCara 
               Caption         =   "&Harian"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   3000
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   1380
               Visible         =   0   'False
               Width           =   930
            End
            Begin BiSANumberBoxProject.BiSANumberBox nPersBunga1 
               Height          =   330
               Left            =   420
               TabIndex        =   1
               Top             =   630
               Width           =   3315
               _ExtentX        =   5847
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
               Caption         =   "Suku Bunga"
               CaptionWidth    =   2500
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
               Height          =   330
               Left            =   420
               TabIndex        =   2
               Top             =   1020
               Width           =   3315
               _ExtentX        =   5847
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
               Caption         =   "Lama"
               CaptionWidth    =   2500
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
            Begin BiSANumberBoxProject.BiSANumberBox nPlafond 
               Height          =   330
               Left            =   5520
               TabIndex        =   3
               Top             =   555
               Width           =   3780
               _ExtentX        =   6668
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
               BackColor       =   16777215
               Caption         =   "Plafond"
               CaptionWidth    =   1800
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
               Left            =   5520
               TabIndex        =   4
               Top             =   930
               Width           =   3780
               _ExtentX        =   6668
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
               BackColor       =   16777215
               Caption         =   "Total Bunga"
               CaptionWidth    =   1800
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
            Begin BiSANumberBoxProject.BiSANumberBox nAdministrasi 
               Height          =   330
               Left            =   5520
               TabIndex        =   5
               Top             =   1305
               Width           =   2700
               _ExtentX        =   4763
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
               Caption         =   "Administrasi"
               CaptionWidth    =   1800
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
            Begin BiSANumberBoxProject.BiSANumberBox nMaterai 
               Height          =   330
               Left            =   5520
               TabIndex        =   6
               Top             =   2070
               Width           =   3780
               _ExtentX        =   6668
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
               Caption         =   "Materai"
               CaptionWidth    =   1800
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
            Begin BiSANumberBoxProject.BiSANumberBox nTotal 
               Height          =   330
               Left            =   5520
               TabIndex        =   7
               Top             =   4260
               Width           =   3780
               _ExtentX        =   6668
               _ExtentY        =   582
               Appearance      =   0
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   15456971
               ForeColor       =   -2147483635
               Caption         =   "Total Realisasi"
               CaptionWidth    =   1800
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
               Left            =   420
               TabIndex        =   68
               Top             =   2055
               Width           =   3315
               _ExtentX        =   5847
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
               Caption         =   "Min Bayar (Dlm 1 Bulan)"
               CaptionWidth    =   2500
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
            Begin BiSANumberBoxProject.BiSANumberBox nNotaris 
               Height          =   330
               Left            =   5520
               TabIndex        =   70
               Top             =   2445
               Width           =   3780
               _ExtentX        =   6668
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
               Caption         =   "Notaris"
               CaptionWidth    =   1800
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
            Begin BiSANumberBoxProject.BiSANumberBox nProvisi 
               Height          =   330
               Left            =   5520
               TabIndex        =   71
               Top             =   1695
               Width           =   2700
               _ExtentX        =   4763
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
               Caption         =   "Provisi"
               CaptionWidth    =   1800
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
            Begin BiSANumberBoxProject.BiSANumberBox nJumlahProvisi 
               Height          =   330
               Left            =   8310
               TabIndex        =   73
               Top             =   1695
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   582
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
               Caption         =   "% x Plafond ="
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
            Begin BiSANumberBoxProject.BiSANumberBox nKonp 
               Height          =   330
               Left            =   420
               TabIndex        =   74
               Top             =   2880
               Width           =   3315
               _ExtentX        =   5847
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
               Caption         =   "Konpensasi Telat Bayar"
               CaptionWidth    =   2500
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
            Begin BiSANumberBoxProject.BiSANumberBox nLainLain 
               Height          =   330
               Left            =   5520
               TabIndex        =   76
               Top             =   2820
               Width           =   3780
               _ExtentX        =   6668
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
               Caption         =   "Lain-Lain"
               CaptionWidth    =   1800
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
            Begin BiSANumberBoxProject.BiSANumberBox nDendaTelat 
               Height          =   330
               Left            =   420
               TabIndex        =   77
               Top             =   3270
               Width           =   3315
               _ExtentX        =   5847
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
               Caption         =   "Denda Angs (%)"
               CaptionWidth    =   2500
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
            Begin BiSANumberBoxProject.BiSANumberBox nWajibPokok 
               Height          =   330
               Left            =   5520
               TabIndex        =   94
               Top             =   3180
               Width           =   3780
               _ExtentX        =   6668
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
               Caption         =   "Wajib Pokok/Bulan"
               CaptionWidth    =   1800
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
            Begin BiSANumberBoxProject.BiSANumberBox nSimpananWajib 
               Height          =   330
               Left            =   5535
               TabIndex        =   99
               Top             =   3540
               Width           =   3780
               _ExtentX        =   6668
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
               Caption         =   "Simp. Wajib"
               CaptionWidth    =   1800
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
            Begin VB.Line Line1 
               X1              =   5550
               X2              =   9330
               Y1              =   4185
               Y2              =   4185
            End
            Begin VB.Label Label6 
               Caption         =   "Hari"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3840
               TabIndex        =   75
               Top             =   2955
               Width           =   375
            End
            Begin VB.Label Label3 
               Caption         =   "Kali Angsuran"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3840
               TabIndex        =   69
               Top             =   2115
               Width           =   1185
            End
            Begin VB.Label Label2 
               Caption         =   "Kali Angsuran"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3840
               TabIndex        =   67
               Top             =   2535
               Width           =   1185
            End
            Begin VB.Label Label1 
               Caption         =   "Cara Angsuran"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   450
               TabIndex        =   63
               Top             =   1440
               Width           =   1485
            End
            Begin VB.Label Label4 
               Caption         =   "% / Bulan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3840
               TabIndex        =   9
               Top             =   705
               Width           =   900
            End
            Begin VB.Label Label5 
               Caption         =   "Bulan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3840
               TabIndex        =   8
               Top             =   1095
               Width           =   600
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame6 
            Height          =   4905
            Left            =   -74955
            Top             =   360
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   8652
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   1
               Left            =   6615
               TabIndex        =   39
               Top             =   105
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
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
            Begin BiSANumberBoxProject.BiSANumberBox nNilaiJaminan 
               Height          =   330
               Left            =   135
               TabIndex        =   32
               Top             =   825
               Width           =   3045
               _ExtentX        =   5371
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
               Caption         =   "Nilai"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaJaminan 
               Height          =   330
               Left            =   135
               TabIndex        =   31
               Top             =   465
               Width           =   4050
               _ExtentX        =   7144
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
               Caption         =   "Keterangan"
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
            Begin BiSATextBoxProject.BiSABrowse cJaminan 
               Height          =   330
               Left            =   135
               TabIndex        =   30
               Top             =   105
               Width           =   2055
               _ExtentX        =   3625
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
               Caption         =   "Jaminan"
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   2
               Left            =   6615
               TabIndex        =   40
               Top             =   405
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   3
               Left            =   6615
               TabIndex        =   41
               Top             =   720
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   4
               Left            =   6615
               TabIndex        =   42
               Top             =   1035
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   5
               Left            =   6615
               TabIndex        =   43
               Top             =   1350
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   6
               Left            =   6615
               TabIndex        =   44
               Top             =   1665
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   7
               Left            =   6615
               TabIndex        =   51
               Top             =   1980
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   8
               Left            =   6615
               TabIndex        =   52
               Top             =   2295
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   9
               Left            =   6615
               TabIndex        =   53
               Top             =   2610
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   10
               Left            =   6615
               TabIndex        =   54
               Top             =   2925
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   11
               Left            =   6615
               TabIndex        =   55
               Top             =   3240
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   12
               Left            =   6615
               TabIndex        =   56
               Top             =   3555
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   13
               Left            =   6615
               TabIndex        =   57
               Top             =   3870
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   14
               Left            =   6615
               TabIndex        =   58
               Top             =   4185
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cText 
               Height          =   300
               Index           =   15
               Left            =   6615
               TabIndex        =   61
               Top             =   4500
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
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
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   15
               Left            =   4440
               TabIndex        =   62
               Top             =   4500
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   13
               Left            =   4440
               TabIndex        =   60
               Top             =   3870
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   14
               Left            =   4440
               TabIndex        =   59
               Top             =   4185
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   12
               Left            =   4440
               TabIndex        =   50
               Top             =   3555
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   11
               Left            =   4440
               TabIndex        =   49
               Top             =   3240
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   10
               Left            =   4440
               TabIndex        =   48
               Top             =   2925
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   9
               Left            =   4440
               TabIndex        =   47
               Top             =   2610
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   8
               Left            =   4440
               TabIndex        =   46
               Top             =   2295
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   7
               Left            =   4440
               TabIndex        =   45
               Top             =   1980
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   4440
               TabIndex        =   38
               Top             =   1665
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   4440
               TabIndex        =   37
               Top             =   1350
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   4440
               TabIndex        =   36
               Top             =   1035
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   4440
               TabIndex        =   35
               Top             =   720
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   4440
               TabIndex        =   34
               Top             =   405
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label lLabel 
               Caption         =   "CAPTION                               :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   4440
               TabIndex        =   33
               Top             =   105
               Visible         =   0   'False
               Width           =   2145
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame3 
            Height          =   5385
            Left            =   75
            Top             =   375
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   9499
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
            Begin VB.Frame Frame1 
               Height          =   555
               Left            =   2085
               TabIndex        =   96
               Top             =   4650
               Width           =   4110
               Begin VB.OptionButton optCaraAngsuran 
                  Caption         =   "&2 Flat"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   1
                  Left            =   1350
                  TabIndex        =   98
                  Top             =   195
                  Width           =   915
               End
               Begin VB.OptionButton optCaraAngsuran 
                  Caption         =   "&1 Menurun"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   0
                  Left            =   150
                  TabIndex        =   97
                  Top             =   195
                  Width           =   1200
               End
            End
            Begin BiSAFramProject.BiSAFrame BiSAFrame1 
               Height          =   630
               Left            =   465
               Top             =   1185
               Width           =   5730
               _ExtentX        =   10107
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
               BackColor       =   12632256
               Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
                  Height          =   360
                  Left            =   4290
                  TabIndex        =   79
                  Top             =   150
                  Width           =   420
                  _ExtentX        =   741
                  _ExtentY        =   635
                  Text            =   "12"
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Verdana"
                  FontSize        =   9.75
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
                  Height          =   360
                  Left            =   2580
                  TabIndex        =   80
                  Top             =   150
                  Width           =   750
                  _ExtentX        =   1323
                  _ExtentY        =   635
                  Text            =   "12"
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Verdana"
                  FontSize        =   9.75
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
                  Height          =   360
                  Left            =   210
                  TabIndex        =   81
                  Top             =   150
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   635
                  Text            =   "12"
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Verdana"
                  FontSize        =   9.75
                  MaxLength       =   2
                  Caption         =   "NO. REKENING"
                  CaptionWidth    =   1800
                  CaptionBackColor=   12632256
                  CaptionForeColor=   -2147483635
                  BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin BiSATextBoxProject.BiSATextBox cUrut 
                  Height          =   360
                  Left            =   3330
                  TabIndex        =   82
                  Top             =   150
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   635
                  Text            =   "123456"
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Verdana"
                  FontSize        =   9.75
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
            End
            Begin BiSATextBoxProject.BiSATextBox cNoSPK 
               Height          =   330
               Left            =   465
               TabIndex        =   10
               Top             =   3615
               Width           =   4320
               _ExtentX        =   7620
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
               Caption         =   "No SPK"
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
            Begin BiSATextBoxProject.BiSABrowse cWilayah 
               Height          =   330
               Left            =   465
               TabIndex        =   11
               Top             =   3960
               Width           =   2775
               _ExtentX        =   4895
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
               Caption         =   "Wilayah"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaWilayah 
               Height          =   330
               Left            =   3255
               TabIndex        =   12
               Top             =   3960
               Width           =   2925
               _ExtentX        =   5159
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
            Begin BiSATextBoxProject.BiSABrowse cAO 
               Height          =   330
               Left            =   465
               TabIndex        =   13
               Top             =   4320
               Width           =   2775
               _ExtentX        =   4895
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
               Caption         =   "AO"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaAO 
               Height          =   330
               Left            =   3255
               TabIndex        =   14
               Top             =   4305
               Width           =   2925
               _ExtentX        =   5159
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
            Begin BiSADateProject.BiSADate dTgl 
               Height          =   345
               Left            =   465
               TabIndex        =   21
               Top             =   1845
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   609
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
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
            Begin BiSATextBoxProject.BiSABrowse cNama 
               Height          =   330
               Left            =   465
               TabIndex        =   22
               Top             =   480
               Width           =   5040
               _ExtentX        =   8890
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
            Begin BiSATextBoxProject.BiSATextBox cKode 
               Height          =   330
               Left            =   2550
               TabIndex        =   23
               Top             =   105
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
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
            Begin BiSATextBoxProject.BiSATextBox cCabang1 
               Height          =   330
               Left            =   465
               TabIndex        =   24
               Top             =   105
               Width           =   2055
               _ExtentX        =   3625
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
               Caption         =   "No Register"
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
               Height          =   330
               Left            =   465
               TabIndex        =   25
               Top             =   840
               Width           =   5745
               _ExtentX        =   10134
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
            Begin BiSANumberBoxProject.BiSANumberBox nPlafondPengajuan 
               Height          =   330
               Left            =   465
               TabIndex        =   26
               Top             =   3270
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   582
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
            Begin BiSATextBoxProject.BiSABrowse cPengajuan 
               Height          =   330
               Left            =   465
               TabIndex        =   27
               Top             =   2205
               Width           =   3510
               _ExtentX        =   6191
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
               Caption         =   "No Pengajuan"
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
            Begin BiSATextBoxProject.BiSABrowse cNamaPengajuan 
               Height          =   330
               Left            =   465
               TabIndex        =   28
               Top             =   2550
               Width           =   5745
               _ExtentX        =   10134
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
            Begin BiSATextBoxProject.BiSATextBox cJaminanPengajuan 
               Height          =   330
               Left            =   465
               TabIndex        =   29
               Top             =   2910
               Width           =   5535
               _ExtentX        =   9763
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
               Caption         =   "Jaminan"
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
            Begin BiSATextBoxProject.BiSATextBox cFaktur 
               Height          =   330
               Left            =   3585
               TabIndex        =   84
               Top             =   120
               Visible         =   0   'False
               Width           =   3960
               _ExtentX        =   6985
               _ExtentY        =   582
               Text            =   "12345678901234567890"
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
               MaxLength       =   20
               Caption         =   "FAKTUR"
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
            Begin VB.Label Label7 
               Caption         =   "Cara perhitungan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   510
               TabIndex        =   95
               Top             =   4815
               Width           =   1545
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame5 
            Height          =   4245
            Left            =   -74955
            Top             =   480
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   7488
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
            Begin BiSANumberBoxProject.BiSANumberBox nBiayaRT 
               Height          =   330
               Left            =   4035
               TabIndex        =   85
               Top             =   855
               Width           =   3810
               _ExtentX        =   6720
               _ExtentY        =   582
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
               Caption         =   "Biaya Rumah Tangga"
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
            Begin BiSANumberBoxProject.BiSANumberBox nBiayaTK 
               Height          =   330
               Left            =   4035
               TabIndex        =   86
               Top             =   1215
               Width           =   3810
               _ExtentX        =   6720
               _ExtentY        =   582
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
               Caption         =   "Biaya Telepon"
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
            Begin BiSANumberBoxProject.BiSANumberBox nBiayaListrik 
               Height          =   330
               Left            =   4035
               TabIndex        =   87
               Top             =   1575
               Width           =   3810
               _ExtentX        =   6720
               _ExtentY        =   582
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
               Caption         =   "Biaya Listrik/Air"
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
            Begin BiSANumberBoxProject.BiSANumberBox nBiayaPemeliharaan 
               Height          =   330
               Left            =   4035
               TabIndex        =   88
               Top             =   1920
               Width           =   3810
               _ExtentX        =   6720
               _ExtentY        =   582
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
               Caption         =   "Biaya Pemeliharaan"
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
            Begin BiSANumberBoxProject.BiSANumberBox nBiayaLain 
               Height          =   330
               Left            =   4035
               TabIndex        =   89
               Top             =   2280
               Width           =   3810
               _ExtentX        =   6720
               _ExtentY        =   582
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
               Caption         =   "Biaya Lain - Lain"
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
            Begin BiSANumberBoxProject.BiSANumberBox nPendapatanUtama 
               Height          =   345
               Left            =   75
               TabIndex        =   90
               Top             =   120
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   609
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
               Caption         =   "Pendapatan Utama"
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
            Begin BiSANumberBoxProject.BiSANumberBox nPendapatanLain 
               Height          =   345
               Left            =   75
               TabIndex        =   91
               Top             =   495
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   609
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
               Caption         =   "Pendapatan Lain Lain"
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
            Begin BiSANumberBoxProject.BiSANumberBox nJumlahPendapatan 
               Height          =   345
               Left            =   45
               TabIndex        =   92
               Top             =   2775
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   609
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
               BackColor       =   -2147483633
               Caption         =   "Jumlah"
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
            Begin BiSANumberBoxProject.BiSANumberBox nJumlahBiaya 
               Height          =   345
               Left            =   4035
               TabIndex        =   93
               Top             =   2775
               Width           =   3825
               _ExtentX        =   6747
               _ExtentY        =   609
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
               BackColor       =   -2147483633
               Caption         =   "Jumlah"
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
            Begin VB.Line Line3 
               X1              =   2145
               X2              =   4020
               Y1              =   2670
               Y2              =   2670
            End
            Begin VB.Line Line2 
               X1              =   5895
               X2              =   7785
               Y1              =   2685
               Y2              =   2685
            End
         End
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame8 
      Height          =   630
      Left            =   0
      Top             =   5940
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
         Left            =   2235
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
         Picture         =   "TrRealisasi.frx":008C
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3405
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
         Picture         =   "TrRealisasi.frx":0316
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9495
         TabIndex        =   17
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
         Picture         =   "TrRealisasi.frx":04B5
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1185
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
         Picture         =   "TrRealisasi.frx":08CB
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   120
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
         Picture         =   "TrRealisasi.frx":09F7
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10575
         TabIndex        =   20
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
         Picture         =   "TrRealisasi.frx":0BA2
      End
   End
End
Attribute VB_Name = "TrRealisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim dbData1 As New ADODB.Recordset
Dim xArray As New XArrayDB
Dim vaArray As New XArrayDB
Dim vaRPT As New XArrayDB
Dim nPos  As SisPos
Dim lEdit As Boolean
Dim cSQL As String
Dim NoRek As String
Dim dJatuhTempo As Date
Dim nTotalPokok As Double, nTotalBunga As Double, nTotalAngsuran As Double, nTotalTabungan As Double

Private Sub cAO_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "AO", "Kode,Nama,Alamat", "kode", sisContent, cAO.Text)
  cAO.Text = cAO.Browse(dbData)
  If Not dbData.eof Then
    cNamaAO.Text = GetNull(dbData!nama, "")
  End If
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
Dim cRekening As String

    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    cNoSPK.Text = cRekening
    Set dbData = objData.Browse(GetDSN, "Debitur", "Rekening,StatusPencairan", "Rekening", sisAssign, cRekening)
    If Not dbData.eof Then
      If GetNull(dbData!statuspencairan, "") = "1" Then
        MsgBox "Kredit Sudah Dicairkan, Tidak Bisa Di Koreksi" + vbCrLf + "Jika Ingin Melakukan Koreksi, Hapus Transaksi Pencairannya terlebih dahulu", vbExclamation + vbOKOnly
        Cancel = True
        initvalue
        GetEdit False
        Exit Sub
      End If
      If nPos = Delete Then   ' Hapus
        If GetNull(dbData!statuspencairan, "") = "1" Then
          MsgBox "Kredit Sudah dicairkan.., Data tidak bisa dihapus..", vbExclamation
          initvalue
          GetEdit False
        Else
          DeleteRekening
          initvalue
          GetEdit False
        End If
      End If
      GetMemory
'    Else
'      MsgBox "No. Rekening Tidak Ditemukan, Ulangi Pengisian", vbExclamation + vbOKOnly
'      Cancel = True
'      cFrekuensi.SetFocus
    End If
End Sub

Private Sub cGolongan_ButtonClick()
Dim cRekening As String

  Set dbData = objData.Pick(GetDSN, "GolonganKredit", "Kode", cGolongan, "Kode,Keterangan")
  If nPos = Add Then cFrekuensi.Text = GetFrekuensi("Debitur", cCabang.Text, Kredit, cGolongan.Text, cUrut.Text)
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cNoSPK.Text = cRekening

End Sub

Private Sub cJaminan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gagunan", "kode,keterangan", "kode", sisContent, cJaminan.Text)
  cJaminan.Text = cJaminan.Browse(dbData)
  If Not dbData.eof Then
    cNamaJaminan.Text = GetNull(dbData!Keterangan, "")
    GetDetailJaminan
  End If
End Sub

Private Sub GetDetailJaminan()
  Set dbData = objData.Browse(GetDSN, "GAgunan", , "Kode", sisAssign, cJaminan.Text)
  If Not dbData.eof Then
    With dbData
      DJamin 1, !j1
      DJamin 2, !j2
      DJamin 3, !j3
      DJamin 4, !j4
      DJamin 5, !j5
      DJamin 6, !j6
      DJamin 7, !j7
      DJamin 8, !j8
      DJamin 9, !j9
      DJamin 10, !j10
      DJamin 11, !j11
      DJamin 12, !j12
      DJamin 13, !j13
      DJamin 14, !j14
      DJamin 15, !j15
    End With
  End If
End Sub

Private Sub DJamin(Index As Single, cData As String)
Dim cWid As Double
Dim cMaxLength As Double

  If cData <> "" Then
    lLabel(Index).Caption = left(cData, 47)
    lLabel(Index).Visible = True
    cWid = 200 + TextWidth(Replicate("A", Val(Mid(cData, 47))))
    cMaxLength = Val(Mid(cData, 47))
    If cWid > 2865 Then
      cWid = 2865
    End If
    cText(Index).Width = cWid
    cText(Index).MaxLength = cMaxLength
    If Val(Mid(cData, 47)) > 0 Then
      cText(Index).Visible = True
    End If
  Else
    cText(Index).Visible = False
    lLabel(Index).Visible = False
  End If
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
Dim cNoRegister As String

  cKode.Text = Padl(Trim(cKode.Text), cKode.MaxLength, "0")
  cNoRegister = cCabang1.Text & "." & cKode.Text
  Set dbData = objData.Browse(GetDSN, "RegisterNasabah", "Nama,ALamat", "Kode", sisAssign, cNoRegister)
  If dbData.eof Then
     MsgBox "Maaf, Nomor Register Nasabah : " & cCabang1.Text & "." & cKode.Text & " Tidak Ada. Silahkan Mengulangi Pengisian !", vbInformation + vbOKOnly
     Cancel = True
     cKode.Default
     cNama.Default
     cAlamat.Default
     cKode.SetFocus
     Exit Sub
  End If
  cNama.Text = GetNull(dbData!nama, "")
  cAlamat.Text = GetNull(dbData!alamat, "")
  cUrut.Text = cKode.Text
  Exit Sub
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  lPilih
  cKode.Enabled = True
  cNama.Button = True
  cNama.Enabled = True
  cCabang1.SetFocus
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  initvalue
  GetEdit True
  lPilih
  cCabang.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  initvalue
  lPilih
  cCabang.SetFocus
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    
    xArray.Clear
    Set TDBGrid1.Array = xArray
    TDBGrid1.ReBind
    
    GetEdit False
    initvalue
  End If
End Sub

Private Function DeleteRekening() As Boolean
Dim cRek As String

  DeleteRekening = False
  If ValidDeleted() Then
    If MsgBox("Data Benar-benar Dihapus ?", vbQuestion + vbYesNo) = vbYes Then
      cRek = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
      objData.Delete GetDSN, "Debitur", "Rekening", sisAssign, cRek
      objData.Delete GetDSN, "Agunan", "Kode", sisAssign, cJaminan.Text, "And Rekening='" & cRek & "'"
      DeleteRealisasi objData, cFaktur.Text
      MsgBox "Data Sudah Dihapus", vbExclamation + vbOKOnly
      GetEdit False
      DeleteRekening = True
    End If
  End If
End Function

Private Function ValidDeleted() As Boolean
  ValidDeleted = True
  
'  Cek apakah Pernah ada angsuran
'  If dbData!PelunasanPokok > 0 Or dbData!PelunasanBunga > 0 Then
'    MsgBox "Nomor Rekening Tersebut Pernah Angsur" & vbCrLf & "Data Tidak Bisa Dihapus", vbExclamation
'    ValidDeleted = False
'    Exit Function
'  End If
End Function

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim cRek As String
Dim vaField1, vaValue1

  dJatuhTempo = DateAdd("M", nLama.Value, dTgl.Value)
  cRek = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  If ValidSaving() Then
    If MsgBox("Data benar-benar sudah Valid ?", vbYesNo + vbInformation) = vbYes Then
        SimpanRealisasi objData, cRek, cWilayah.Text, cCabang1.Text & "." & cKode.Text, cGolongan.Text, _
                     cNoSPK.Text, nPersBunga.Value, dTgl.Value, _
                     nPlafond.Value, nLama.Value, cAO.Text, nAdministrasi.Value, _
                     nMaterai.Value, dJatuhTempo, cPengajuan.Text, nBunga.Value, GetOpt(optCara), _
                     nPeriode.Value, nMinimum.Value, nProvisi.Value, nNotaris.Value, nKonp.Value, _
                     nLainLain.Value, nDendaTelat.Value, nWajibPokok.Value, GetOpt(optCaraAngsuran), nSimpananWajib.Value
                     
        'Simpan Jaminan
        SimpanJaminan cRek
        initvalue
        GetEdit False
    End If
  End If
End Sub

Private Sub SimpanRealisasi(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String, ByVal cWilayah As String, _
                 ByVal cKode As String, ByVal cGolonganKredit As String, _
                 ByVal cNoSPK As String, ByVal nPersBunga As Double, _
                 ByVal dTgl As Date, ByVal nPlafond As Double, ByVal nLama As Integer, _
                 ByVal cAO As String, ByVal nAdministrasi As Double, ByVal nMaterai As Double, _
                 ByVal dJatuhTempo As Date, ByVal cNoPengaJuan As String, _
                 ByVal nTotalBunga As Double, ByVal cCaraPembayaran As String, ByVal nPeriod As Integer, _
                 ByVal nMinPeriode As Integer, ByVal nProv As Double, ByVal nNot As Double, _
                 ByVal nKonpensasi As Integer, ByVal nBiayalain As Double, ByVal nDendaKeterlamabatan As Double, ByVal wajibpokok As Double, ByVal cCaraPerhitungan As String, ByVal SimpananWajib As Double)
Dim vaField
Dim vaValue
Dim n As Single

  vaField = Array("Rekening", "Wilayah", "Kode", "GolonganKredit", _
                  "NoSPK", "SukuBunga", "Tgl", _
                  "Plafond", "Lama", "AO", "Administrasi", _
                  "Materai", "JatuhTempo", _
                  "TotalBunga", "NoPengajuan", _
                  "CaraAngsuran", "PeriodeBungaMenurun", _
                  "MinimalPeriode", "Provisi", "Notaris", "BiayaLainLain", _
                  "wajibpokok", "KonpensasiTelat", "DendaTelatBayar", "UserName", "DateTime", "caraperhitungan", "simpananwajib")
                  
  vaValue = Array(cRekening, cWilayah, cKode, cGolonganKredit, _
                  cNoSPK, nPersBunga, dTgl, _
                  nPlafond, nLama, cAO, nAdministrasi, _
                  nMaterai, dJatuhTempo, _
                  nTotalBunga, cNoPengaJuan, _
                  cCaraPembayaran, nPeriod, _
                  nMinPeriode, nProv, nNot, nBiayalain, _
                  wajibpokok, nKonpensasi, nDendaKeterlamabatan, cusername, SNow, cCaraPerhitungan, nSimpananWajib.Value)
                  
  obj.Update GetDSN, "Debitur", "Rekening = '" & cRekening & "'", vaField, vaValue
  obj.Edit GetDSN, "PengajuanKredit", "Kode='" & cNoPengaJuan & "'", Array("StatusPengajuan"), Array("1")
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cUrut.Text, "Nomor Urut Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    Exit Function
  End If
  
  If Not CheckData(cFrekuensi.Text, "Nomor Frekuensi Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    Exit Function
  End If
  
  If Not CheckData(cKode.Text, "Kode Debitur Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cNoSPK.Text, "Nomor SPK Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cNoSPK.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cWilayah.Text, "Wilayah Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cWilayah.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cAO.Text, "A/O Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    cAO.SetFocus
    Exit Function
  End If
  
  If Not CheckData(nPersBunga.Value, "Suku Bunga Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    'nPersBunga.SetFocus
    Exit Function
  End If
End Function

Private Sub cNama_ButtonClick()
  If nPos = Add Then
    Set dbData = objData.Browse(GetDSN, "RegisterNasabah", "Nama,Alamat,Kode", "Nama", sisContent, cNama.Text, , "Nama")
    cNama.Text = cNama.Browse(dbData)
    If Not dbData.eof Then
      cKode.Text = Right(dbData!Kode, 6)
      cNama.Text = GetNull(dbData!nama, "")
      cAlamat.Text = GetNull(dbData!alamat, "")
      cUrut.Text = cKode.Text
    End If
  End If
End Sub

Private Sub cNamaPengajuan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "PengajuanKredit", "Nama,Kode,Jaminan,Plafond", "Nama", sisContent, cNamaPengajuan.Text, "And StatusPengajuan <> '1'", "Nama")
  cNamaPengajuan.Text = cNamaPengajuan.Browse(dbData)
  If Not dbData.eof Then
    cPengajuan.Text = GetNull(dbData!Kode, "")
    cJaminanPengajuan.Text = GetNull(dbData!Jaminan, "")
    nPlafondPengajuan.Value = GetNull(dbData!plafond)
  End If
End Sub

Private Sub GetMemory()
Dim n As Integer
Dim vaJoin
Dim cField As String
Dim cRekening As String

  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cField = "d.Tgl,d.faktur,d.NoSPk,d.SukuBunga,d.Plafond,d.Lama,d.AO,d.NoPengajuan,d.CaraAngsuran,d.PeriodeBungaMenurun,d.MinimalPeriode,d.KonpensasiTelat,d.BiayalainLain,d.DendaTelatBayar,d.wajibpokok,d.caraperhitungan,d.caraperhitungan,"
  cField = cField & " d.Wilayah,d.Administrasi,d.Materai,d.Provisi,d.Notaris,d.simpananwajib,"
  cField = cField & " d.Kode as KodeDebitur,r.Nama, r.Alamat,w.Keterangan as NamaWilayah,"
  cField = cField & " h.Nama as NamaAO,"
  cField = cField & " p.Nama as NamaPengajuan,p.Jaminan as JaminanPengajuan,p.Plafond as PlafondPengajuan"
  
  vaJoin = Array("Left Join Wilayah w on w.Kode = d.Wilayah", _
                 "Left Join RegisterNasabah r on r.Kode = d.Kode", _
                 "Left Join AO h on h.Kode = d.AO", _
                 "Left Join PengajuanKredit p on p.Kode = d.NoPengajuan")
                 
  Set dbData = objData.Browse(GetDSN, "Debitur d", cField, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
    cCabang1.Text = left(GetNull(dbData!KodeDebitur, ""), 2)
    cKode.Text = Right(GetNull(dbData!KodeDebitur, ""), 6)
    dTgl.Value = GetNull(dbData!Tgl, "")
    cFaktur.Text = GetNull(dbData!Faktur, "")
    cNama.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cWilayah.Text = GetNull(dbData!Wilayah, "")
    cNamaWilayah.Text = GetNull(dbData!Namawilayah, "")
    cNoSPK.Text = GetNull(dbData!NoSPK, "")
    nPersBunga.Value = GetNull(dbData!SukuBunga)
    nPersBunga1.Value = nPersBunga.Value / 12
    nPlafond.Value = GetNull(dbData!plafond)
    nLama.Value = GetNull(dbData!Lama)
    cAO.Text = GetNull(dbData!AO, "")
    cNamaAO.Text = GetNull(dbData!namaao, "")
    nWajibPokok.Value = GetNull(dbData!wajibpokok, 0)
    
    nAdministrasi.Value = GetNull(dbData!Administrasi)
    nMaterai.Value = GetNull(dbData!Materai)
    nProvisi.Value = GetNull(dbData!Provisi)
    nNotaris.Value = GetNull(dbData!Notaris)
    nLainLain.Value = GetNull(dbData!BiayaLainLain)
    
    cPengajuan.Text = GetNull(dbData!NoPengajuan)
    cNamaPengajuan.Text = GetNull(dbData!NamaPengajuan)
    cJaminanPengajuan.Text = GetNull(dbData!JaminanPengajuan)
    nPlafondPengajuan.Value = GetNull(dbData!PlafondPengajuan, 0)
    SetOpt optCara, GetNull(dbData!CaraAngsuran)
    nPeriode.Value = GetNull(dbData!PeriodeBungaMenurun)
    nMinimum.Value = GetNull(dbData!MinimalPeriode)
    nKonp.Value = GetNull(dbData!KonpensasiTelat)
    nDendaTelat.Value = GetNull(dbData!DendaTelatBayar)
    nSimpananWajib.Value = GetNull(dbData!SimpananWajib)
    Select Case GetNull(dbData!caraperhitungan)
      Case "1"
        GetJadwalMenurunNonPeriodik
      Case "2"
        GetJadwalFlat
    End Select
    SetOpt optCaraAngsuran, GetNull(dbData!caraperhitungan)
    GetDataJaminan cRekening
    SumTotal
  End If
End Sub

Private Sub cPengajuan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "PengajuanKredit p", "p.*", "p.Kode", sisContent, cPengajuan.Text, "And p.StatusPengajuan <> '1'")
  cPengajuan.Text = cPengajuan.Browse(dbData)
  If Not dbData.eof Then
    cNamaPengajuan.Text = GetNull(dbData!nama, "")
    cJaminanPengajuan.Text = GetNull(dbData!Jaminan, "")
    nPlafondPengajuan.Value = GetNull(dbData!plafond)
    GetDataAnalisaKeuangan
  End If
End Sub

Private Sub GetDataAnalisaKeuangan()
  'data analisa keuangan
  nBiayaRT.Value = GetNull(dbData!nBiayaRT)
  nBiayaTK.Value = GetNull(dbData!nBiayaTK)
  nBiayaListrik.Value = GetNull(dbData!nBiayaListrik)
  nBiayaPemeliharaan.Value = GetNull(dbData!nBiayaPemeliharaan)
  nBiayalain.Value = GetNull(dbData!nBiayalain)
  nPendapatanLain.Value = GetNull(dbData!nPendapatanLain)
  nPendapatanUtama.Value = GetNull(dbData!nPendapatanUtama)
  nPlafond.Value = GetNull(dbData!plafond)
  SUMJUMLAH
End Sub

Private Sub SUMJUMLAH()
  nJumlahPendapatan.Value = nPendapatanLain.Value + nPendapatanUtama.Value
  nJumlahBiaya.Value = nBiayaRT.Value + _
                       nBiayaTK.Value + _
                       nBiayaListrik.Value + _
                       nBiayaPemeliharaan.Value + _
                       nBiayalain.Value
End Sub

Private Sub cText_Validate(Index As Integer, Cancel As Boolean)
  If cText(Index).LastKey = 13 Or cText(Index).LastKey = 40 Then
    If cText(Index + 1).Visible = False Then
      nPersBunga1.SetFocus
      SSTab1.Tab = 2
      Exit Sub
    End If
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub cWilayah_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "wilayah", "kode,keterangan", "kode", sisContent, cWilayah.Text)
  cWilayah.Text = cWilayah.Browse(dbData)
  If Not dbData.eof Then
    cNamaWilayah.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Or (dTgl.Value > Date) Then
    Cancel = True
    dTgl.SetFocus
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
Dim i As Integer

  CenterForm Me, True
  SSTab1.Tab = 0
  GetEdit False
  initvalue
  cCabang.Text = aCfg(msKodeCabang)
  cCabang1.Text = cCabang.Text
    
  TabIndex cCabang1, n
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cCabang, n
  
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex dTgl, n
  TabIndex cPengajuan, n
  TabIndex cNamaPengajuan, n
  
  TabIndex cNoSPK, n
  TabIndex cWilayah, n
  TabIndex cAO, n
  TabIndex optCaraAngsuran(0), n
  TabIndex optCaraAngsuran(1), n
  
  TabIndex cJaminan, n
  TabIndex nNilaiJaminan, n
  For i = 1 To 15
    TabIndex cText(i), n
  Next
  
  TabIndex nPersBunga1, n
  TabIndex nLama, n
  TabIndex optCara(0), n
  TabIndex optCara(1), n
  TabIndex nMinimum, n
  TabIndex nPeriode, n
  TabIndex nKonp, n
  TabIndex nDendaTelat, n
  TabIndex nPlafond, n
  TabIndex nBunga, n
  TabIndex nAdministrasi, n
  TabIndex nProvisi, n
  TabIndex nMaterai, n
  TabIndex nNotaris, n
  TabIndex nLainLain, n
  TabIndex nWajibPokok, n
  TabIndex nSimpananWajib, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub initvalue()
Dim i As Integer

  dTgl.Value = Date
  cFaktur.Default
    
  cKode.Default
  cNama.Default
  cAlamat.Default
  
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  
  cPengajuan.Default
  cNamaPengajuan.Default
  cJaminanPengajuan.Default
  nPlafondPengajuan.Value = 0
  cNoSPK.Default
  cWilayah.Default
  cNamaWilayah.Default
  nPersBunga.Value = 0
  nPersBunga1.Value = 0
  nLama.Value = 1
  nPlafond.Value = 0
  nBunga.Value = 0
  nDendaTelat.Value = 3
  nAdministrasi.Value = 2
  nMaterai.Value = 0
  nProvisi.Value = 1
  nJumlahProvisi.Value = 0
  nJumlahADM.Value = 0
  nNotaris.Value = 0
  nTotal.Value = 0
  cAO.Default
  cNamaAO.Default
  optCara(1).Value = True
  nPeriode.Value = 0
  nMinimum.Value = 0
  nKonp.Value = 7
  nLainLain.Value = 0
  nWajibPokok.Default
  nSimpananWajib.Default
  optCaraAngsuran(0).Value = True
  
  cJaminan.Default
  cNamaJaminan.Default
  nNilaiJaminan.Value = 0
  For i = 1 To 15
    cText(i).Default
    cText(i).Visible = False
    lLabel(i).Visible = False
  Next
  
  xArray.Clear
  xArray.ReDim 0, -1, 0, 6
  Set TDBGrid1.Array = xArray
  TDBGrid1.ReBind
  
  SSTab1.Tab = 0
End Sub

Private Sub nAdministrasi_Change()
  nJumlahADM.Value = Round(nAdministrasi.Value / 100 * nPlafond.Value)
  SumTotal
End Sub

Private Sub nAdministrasi_Validate(Cancel As Boolean)
  nJumlahADM.Value = Round(nAdministrasi.Value / 100 * nPlafond.Value)
  SumTotal
End Sub

Private Sub nBunga_Validate(Cancel As Boolean)
  If optCaraAngsuran(0).Value = True Then
    GetJadwalMenurunNonPeriodik
  ElseIf optCaraAngsuran(1).Value = True Then
    GetJadwalFlat
  End If
End Sub

Private Sub nLainLain_Change()
  SumTotal
End Sub

Private Sub nMaterai_Change()
  SumTotal
End Sub

Private Sub nNotaris_Change()
  SumTotal
End Sub

Private Sub nPersBunga_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Then
    SSTab1.Tab = 1
    nNilaiJaminan.SetFocus
    Exit Sub
  End If
End Sub

Private Sub nPersBunga1_Change()
  nPersBunga.Value = nPersBunga1.Value * 12
End Sub

Private Sub nPlafond_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    If nPlafond.Value <= 0 Then
      MsgBox "Nilai Plafond tidak NOL atau lebih kecil NOL. Silahkan ulangi pengisian.", vbOKOnly + vbInformation, "Realisasi Kredit"
      nPlafond.SetFocus
      Exit Sub
    End If
    SumTotal
  End If
End Sub

Private Sub SumTotal()
  nBunga.Value = Round(nPlafond.Value * nPersBunga.Value / 100 / 12 * nLama.Value, 0)
  nTotal.Value = nPlafond.Value - nJumlahADM.Value - nMaterai.Value - nJumlahProvisi.Value - nNotaris.Value - nLainLain.Value - nSimpananWajib.Value
End Sub

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame2.Enabled = lPar
  
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  initvalue
End Sub

Private Sub GetJadwalMenurunNonPeriodik()
Dim n As Single
Dim dTanggal As Date
Dim nSukuBungaPerBulan As Double
Dim nKe As Integer

  nTotalPokok = 0
  nTotalBunga = 0
  xArray.ReDim 0, nLama.Value, 0, 6
  dTanggal = (DateAdd("m", 1, dTgl.Value))
  nSukuBungaPerBulan = Round(nPersBunga.Value / 12, 2)
  xArray(0, 5) = nBunga.Value
  xArray(0, 6) = nPlafond.Value
  nKe = 1
  For n = 1 To nLama.Value
    xArray(n, 0) = n
    xArray(n, 1) = dTanggal
    xArray(n, 2) = GetBungaReguler(xArray(n - 1, 6), nSukuBungaPerBulan)
    xArray(n, 3) = nPlafond.Value / (nLama.Value)
    xArray(n, 4) = xArray(n, 2) + xArray(n, 3)
    xArray(n, 5) = xArray(n - 1, 5) - xArray(n, 2)
    xArray(n, 6) = xArray(n - 1, 6) - xArray(n, 3)
    dTanggal = (DateAdd("m", 1, xArray(n, 1)))
  Next

  For n = 1 To xArray.UpperBound(1)
    nTotalBunga = nTotalBunga + xArray(n, 2)
    nTotalPokok = nTotalPokok + xArray(n, 3)
  Next

  TDBGrid1.Columns(2).FooterText = Format(nTotalBunga, "##,###,###,##0")
  TDBGrid1.Columns(3).FooterText = Format(nTotalPokok, "##,###,###,##0")

  TDBGrid1.Array = xArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub GetJadwalFlat()
Dim n As Single
Dim dTanggal As Date
Dim nSukuBungaPerBulan As Double
Dim nKe As Integer

  nTotalPokok = 0
  nTotalBunga = 0
  xArray.ReDim 0, nLama.Value, 0, 6
  dTanggal = (DateAdd("m", 1, dTgl.Value))
  nSukuBungaPerBulan = Round(nPersBunga.Value / 12, 2)
  xArray(0, 5) = nBunga.Value
  xArray(0, 6) = nPlafond.Value
  nKe = 1
  For n = 1 To nLama.Value
    xArray(n, 0) = n
    xArray(n, 1) = dTanggal
    xArray(n, 2) = Devide(nBunga.Value, nLama.Value) 'GetBungaReguler(xArray(n - 1, 6), nSukuBungaPerBulan)
    xArray(n, 3) = nPlafond.Value / (nLama.Value)
    xArray(n, 4) = xArray(n, 2) + xArray(n, 3)
    xArray(n, 5) = xArray(n - 1, 5) - xArray(n, 2)
    xArray(n, 6) = xArray(n - 1, 6) - xArray(n, 3)
    dTanggal = (DateAdd("m", 1, xArray(n, 1)))
  Next

  For n = 1 To xArray.UpperBound(1)
    nTotalBunga = nTotalBunga + xArray(n, 2)
    nTotalPokok = nTotalPokok + xArray(n, 3)
  Next

  TDBGrid1.Columns(2).FooterText = Format(nTotalBunga, "##,###,###,##0")
  TDBGrid1.Columns(3).FooterText = Format(nTotalPokok, "##,###,###,##0")

  TDBGrid1.Array = xArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Function GetBungaReguler(ByVal nSisaPokok As Double, ByVal nBunga As Double) As Double
  GetBungaReguler = nSisaPokok * (nBunga / 100)
  GetBungaReguler = Mod50(GetBungaReguler)
End Function

Private Sub nPlafond_Validate(Cancel As Boolean)
  nBunga.Value = Mod50(Round(nPersBunga.Value / 100 * nLama.Value * nPlafond.Value / 12, 0))
End Sub

Private Sub nProvisi_Change()
  nJumlahProvisi.Value = Round(nProvisi.Value / 100 * nPlafond.Value)
  SumTotal
End Sub

Private Sub nProvisi_Validate(Cancel As Boolean)
  nJumlahProvisi.Value = Round(nProvisi.Value / 100 * nPlafond.Value)
  SumTotal
End Sub

Private Sub optCara_Click(Index As Integer)
  If Index = 0 Then
    nPeriode.Enabled = True
    nPeriode.BackColor = &H80000005
    nMinimum.Enabled = True
    nMinimum.BackColor = &H80000005
    nKonp.Enabled = False
    nKonp.BackColor = &H8000000F
  Else
    nPeriode.Enabled = False
    nPeriode.BackColor = &H8000000F
    nMinimum.Enabled = False
    nMinimum.BackColor = &H8000000F
    nPeriode.Value = 0
    nMinimum.Value = 0
    nKonp.Enabled = True
    nKonp.BackColor = &H80000005
  End If
End Sub

Private Sub optCara_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optCaraAngsuran_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    SSTab1.Tab = 1
    cJaminan.SetFocus
  End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  Value = Format(Value, "###,###,###,###,###,##0")
End Sub

Private Sub lPilih()
  If nPos = Add Then
    cGolongan.Enabled = True
    cUrut.Enabled = False
    cFrekuensi.Enabled = False
  Else
    cGolongan.Enabled = True
    cUrut.Enabled = True
    cFrekuensi.Enabled = True
  End If
End Sub

Private Sub SimpanJaminan(ByVal cRekening As String)
Dim vaField
Dim vaValue
    
    objData.Delete GetDSN, "Agunan", "rekening", sisAssign, cRekening
    vaField = Array("Kode", "Rekening", "NilaiJaminan", "j1", "j2", "j3", "j4", "j5", "j6", "j7", "j8", "j9", "j10", "j11", "j12", "j13", "j14", "j15")
    vaValue = Array(cJaminan.Text, cRekening, nNilaiJaminan.Value, cText(1).Text, cText(2).Text, cText(3).Text, cText(4).Text, cText(5).Text, _
                  cText(6).Text, cText(7).Text, cText(8).Text, cText(9).Text, cText(10).Text, cText(11).Text, cText(12).Text, _
                  cText(13).Text, cText(14).Text, cText(15).Text)
    objData.Update GetDSN, "Agunan", "Kode='" & cJaminan.Text & "' And Rekening='" & cRekening & "'", vaField, vaValue
End Sub

Private Sub GetDataJaminan(ByVal cRekening As String)
Dim dbJaminan As New ADODB.Recordset
    
    Set dbData = objData.Browse(GetDSN, "Agunan a", "a.*,g.Keterangan", "a.Rekening", sisAssign, cRekening, , , _
                              Array("Left Join GAgunan g on g.Kode=a.Kode"))
    If Not dbData.eof Then
        dbData.MoveFirst
        cJaminan.Text = GetNull(dbData!Kode, "")
        cNamaJaminan.Text = GetNull(dbData!Keterangan, "")
        nNilaiJaminan.Value = GetNull(dbData!nilaijaminan)
        Set dbJaminan = objData.Browse(GetDSN, "Gagunan", , "Kode", sisAssign, cJaminan.Text)
          If Not dbJaminan.eof Then
            DJamin 1, GetNull(dbJaminan!j1, "")
            cText(1).Text = GetNull(dbData!j1, "")
            DJamin 2, GetNull(dbJaminan!j2, "")
            cText(2).Text = GetNull(dbData!j2, "")
            DJamin 3, GetNull(dbJaminan!j3, "")
            cText(3).Text = GetNull(dbData!j3, "")
            DJamin 4, GetNull(dbJaminan!j4, "")
            cText(4).Text = GetNull(dbData!j4, "")
            DJamin 5, GetNull(dbJaminan!j5, "")
            cText(5).Text = GetNull(dbData!j5, "")
            DJamin 6, GetNull(dbJaminan!j6, "")
            cText(6).Text = GetNull(dbData!j6, "")
            DJamin 7, GetNull(dbJaminan!j7, "")
            cText(7).Text = GetNull(dbData!j7, "")
            DJamin 8, GetNull(dbJaminan!j8, "")
            cText(8).Text = GetNull(dbData!j8, "")
            DJamin 9, GetNull(dbJaminan!j9, "")
            cText(9).Text = GetNull(dbData!j9, "")
            DJamin 10, GetNull(dbJaminan!j10, "")
            cText(10).Text = GetNull(dbData!j10, "")
            DJamin 11, GetNull(dbJaminan!j11, "")
            cText(11).Text = GetNull(dbData!j11, "")
            DJamin 12, GetNull(dbJaminan!j12, "")
            cText(12).Text = GetNull(dbData!j12, "")
            DJamin 13, GetNull(dbJaminan!j13, "")
            cText(13).Text = GetNull(dbData!j13, "")
            DJamin 14, GetNull(dbJaminan!j14, "")
            cText(14).Text = GetNull(dbData!j14, "")
            DJamin 15, GetNull(dbJaminan!j15, "")
            cText(15).Text = GetNull(dbData!j15, "")
          End If
    End If
End Sub
