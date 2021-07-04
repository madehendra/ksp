VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trKoreksiMutasiDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HAPUS MUTASI DEPOSITO"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11550
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5760
      Left            =   -30
      Top             =   15
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   10160
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
      Begin BiSAFramProject.BiSAFrame frmPesan 
         Height          =   510
         Left            =   6465
         Top             =   60
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   900
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "REKENING INI SUDAH DICARIKAN (TUTUP)"
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   120
            Width           =   4725
         End
      End
      Begin BiSANumberBoxProject.BiSANumberBox cJangkaWaktu 
         Height          =   330
         Left            =   150
         TabIndex        =   0
         Top             =   1560
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Decimals        =   0
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
         Caption         =   "JANGKA WAKTU"
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
      Begin BiSADateProject.BiSADate dValuta 
         Height          =   330
         Left            =   150
         TabIndex        =   1
         Top             =   1200
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
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
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Caption         =   "TGL VALUTA"
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
         Left            =   150
         TabIndex        =   2
         Top             =   465
         Width           =   5160
         _ExtentX        =   9102
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
         Caption         =   "NAMA DEPOSAN"
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
         Left            =   150
         TabIndex        =   3
         Top             =   810
         Width           =   5970
         _ExtentX        =   10530
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
      Begin BiSADateProject.BiSADate dTempo 
         Height          =   330
         Left            =   3105
         TabIndex        =   4
         Top             =   1200
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   582
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
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Caption         =   "JATUH TEMPO"
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
      Begin BiSATextBoxProject.BiSATextBox cCabang 
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   105
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         Text            =   "1234"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         MaxLength       =   4
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   2190
         TabIndex        =   6
         Top             =   105
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
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
      Begin BiSATextBoxProject.BiSATextBox cUrut 
         Height          =   330
         Left            =   3030
         TabIndex        =   7
         Top             =   105
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Text            =   "123456"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
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
      Begin BiSATextBoxProject.BiSATextBox cFrekuensi 
         Height          =   330
         Left            =   3930
         TabIndex        =   8
         Top             =   105
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
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
      Begin BiSANumberBoxProject.BiSANumberBox nNominalDeposito 
         Height          =   330
         Left            =   3105
         TabIndex        =   9
         Top             =   1560
         Width           =   3960
         _ExtentX        =   6985
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
         BackColor       =   -2147483633
         Caption         =   "NOMINAL"
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
         Height          =   3720
         Left            =   90
         TabIndex        =   13
         Top             =   1965
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   6562
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "TANGGAL"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NO. TRANSAKSI"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KETERANGAN"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "JUMLAH"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "USER NAME"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=476"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=397"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1746"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1667"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3757"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3678"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=7832"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=7752"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2752"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2672"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=3043"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2963"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H80000001&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000E&,.bold=0,.fontsize=825,.italic=0,.underline=0"
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
         _StyleDefs(25)  =   "Splits(0).Style:id=75,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=76,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=77,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=78,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=80,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=75"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=76"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=77"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=79"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=75"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=76"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=77"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=79"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=75"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=76"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=77"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=79"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=75,.alignment=3"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=76"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=77"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=79"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=75,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=76"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=77"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=79"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=90,.parent=75"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=87,.parent=76"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=88,.parent=77"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=89,.parent=79"
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
         _StyleDefs(79)  =   ":id=41,.parent=34"
         _StyleDefs(80)  =   "Named:id=42:FilterBar"
         _StyleDefs(81)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label6 
         Caption         =   "Bulan"
         Height          =   195
         Left            =   2460
         TabIndex        =   10
         Top             =   1620
         Width           =   435
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   -30
      Top             =   5775
      Width           =   11550
      _ExtentX        =   20373
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
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9195
         TabIndex        =   11
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
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
         Picture         =   "trKoreksiMutasiDeposito.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10275
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
         Picture         =   "trKoreksiMutasiDeposito.frx":0416
      End
   End
End
Attribute VB_Name = "trKoreksiMutasiDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim cRekening As String
Dim vaArray As New XArrayDB

Private Sub initvalue()
  cGolongan.Default
  cUrut.Default
  cFrekuensi.Default
  cNama.Default
  cAlamat.Default
  nNominalDeposito.Value = 0
  cJangkaWaktu.Default
  frmPesan.Visible = False
  
  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 6
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub

Private Sub cAlamat_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.nama,r.Alamat,d.Rekening", "r.Alamat", sisContent, cAlamat.Text, , "r.Nama", _
                              Array("Left Join RegisterNasabah r on r.Kode=d.Kode"))
  cAlamat.Text = cAlamat.Browse(dbData)
  If Not dbData.eof Then
    cCabang.Text = left(GetNull(dbData!Rekening, ""), 2)
    cGolongan.Text = Mid(GetNull(dbData!Rekening, ""), 4, 2)
    cUrut.Text = Mid(GetNull(dbData!Rekening, ""), 7, 6)
    cFrekuensi.Text = Right(GetNull(dbData!Rekening, ""), 2)
    GetData
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "GolonganDeposito", "Kode", cGolongan, "Kode,Keterangan")
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub GetData()
Dim cFields As String
Dim vaJoin
Dim cRekening As String
  
  cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
  cFields = "d.Rekening,d.Nominaldeposito,d.GolonganDeposito,d.Tgl,d.jthtmp,d.Status,"
  cFields = cFields & " r.Nama,r.Alamat,r.Telepon,r.Path,b.Lama"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode = d.Kode", _
                 "Left Join GolonganDeposito b on b.Kode=d.GolonganDeposito")
  Set dbData = objData.Browse(GetDSN, "Deposito d", cFields, "d.Rekening", sisAssign, cRekening, , , vaJoin)
  If Not dbData.eof Then
      cNama.Text = GetNull(dbData!nama, "")
      cAlamat.Text = GetNull(dbData!alamat, "")
      dValuta.Value = GetNull(dbData!Tgl, "")
      dTempo.Value = GetNull(dbData!jthtmp, "")
      nNominalDeposito.Value = GetNull(dbData!nominaldeposito, "")
      cJangkaWaktu.Value = GetNull(dbData!Lama)
      GetMutasi
  End If
End Sub

Private Sub GetMutasi()
Dim n As Integer
  
  vaArray.ReDim 0, -1, 0, 6
  Set dbData = objData.Browse(GetDSN, "MutasiDeposito", , "Rekening", sisAssign, cRekening, , "ID")
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = 0
      vaArray(n, 1) = GetNull(dbData!Tgl)
      vaArray(n, 2) = GetNull(dbData!Faktur, "")
      vaArray(n, 3) = KeteranganMutasi(GetNull(dbData!KodeMutasi, ""))
      vaArray(n, 4) = GetNull(dbData!Jumlah)
      vaArray(n, 5) = GetNull(dbData!UserName, "")
      vaArray(n, 6) = GetNull(dbData!KodeMutasi, "")
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub

Private Function KeteranganMutasi(ByVal cKode As String) As String
  KeteranganMutasi = ""
  Select Case cKode
    Case "1"
      KeteranganMutasi = "Entri Nominal Deposito"
    Case "2"
      KeteranganMutasi = "Pencairan Pokok Deposito"
    Case "3"
      KeteranganMutasi = "Pencairan Bunga Deposito"
    Case "4"
      KeteranganMutasi = "Finalti Pencairan Pokok Deposito"
  End Select
End Function

Private Sub cmdSimpan_Click()
Dim n As Integer

  If MsgBox("Apakah Rekening Benar-benar akan disimpan?", vbYesNo + vbInformation) = vbYes Then
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 0) = -1 Then
        objData.Delete GetDSN, "MutasiDeposito", "Faktur", sisAssign, vaArray(n, 2)
        objData.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msDeposito, "And Faktur='" & vaArray(n, 2) & "'"
        
        If vaArray(n, 6) = "1" Then
          objData.Edit GetDSN, "Deposito", "Rekening='" & cRekening & "'", Array("NominalDeposito", "Status", "TglCair"), Array(0, "", "")
        End If
        
        If vaArray(n, 6) = "2" Then
          objData.Edit GetDSN, "Deposito", "Rekening='" & cRekening & "'", Array("Status", "TglCair"), Array("", "")
        End If
      End If
    Next
  End If
  GetMutasi
  cGolongan.SetFocus
  Exit Sub
End Sub

Private Sub cFrekuensi_Validate(Cancel As Boolean)
  If cFrekuensi.LastKey = 13 Then
    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    Set dbData = objData.Browse(GetDSN, "Deposito", "Rekening,Status", "Rekening", sisAssign, cRekening)
    If Not dbData.eof Then
      If GetNull(dbData!status, "") = "1" Then
        frmPesan.Visible = True
'        MsgBox "Rekening tersebut sudah Tutup (Cair) !", vbOKOnly, "Blokir Tabungan"
'        Initvalue
'        cGolongan.SetFocus
'        Exit Sub
      End If
      GetData
      Exit Sub
    End If
    MsgBox "Rekening dengan Nomor : " & cRekening & " Tidak ada. Silahkan Ulangi pengisian.", vbOKOnly + vbExclamation, "Blokir Rekening Deposito"
    Cancel = True
    cFrekuensi.Default
    cFrekuensi.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Deposito d", "r.nama,r.Alamat,d.Rekening,d.Status", "r.Nama", sisContent, cNama.Text, , "r.Nama", _
                              Array("Left Join RegisterNasabah r on r.Kode=d.Kode"))
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.eof Then
    cCabang.Text = left(GetNull(dbData!Rekening, ""), 2)
    cGolongan.Text = Mid(GetNull(dbData!Rekening, ""), 4, 2)
    cUrut.Text = Mid(GetNull(dbData!Rekening, ""), 7, 6)
    cFrekuensi.Text = Right(GetNull(dbData!Rekening, ""), 2)
    cRekening = SetNomorRekening(cCabang.Text, cGolongan.Text, cUrut.Text, cFrekuensi.Text)
    If GetNull(dbData!status, "") = "1" Then
      frmPesan.Visible = True
    End If
    GetData
  End If
End Sub

Private Sub cUrut_Validate(Cancel As Boolean)
  cUrut.Text = Padl(cUrut.Text, cUrut.MaxLength, "0")
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me, True
  initvalue
  cCabang.Text = aCfg(msKodeCabang, "")
  
  TabIndex cCabang, n
  TabIndex cGolongan, n
  TabIndex cUrut, n
  TabIndex cFrekuensi, n
  TabIndex cNama, n

  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub TDBGrid1_Click()
Dim n As Integer
  
  n = TDBGrid1.Bookmark
  If vaArray(n, 0) = 0 Then
    vaArray(n, 0) = -1
  Else
    vaArray(n, 0) = 0
  End If
End Sub

