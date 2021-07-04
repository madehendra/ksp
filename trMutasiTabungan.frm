VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form trMutasiTabungan 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin BiSAFramProject.BiSAFrame trMutasiTabungan 
      Height          =   4275
      Left            =   15
      Top             =   15
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   7541
      Caption         =   "MUTASI TABUNGAN"
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
      Begin BiSAFramProject.BiSAFrame BISAPESAN 
         Height          =   840
         Left            =   7050
         Top             =   1530
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   1482
         Caption         =   "REKENING DIBLOKIR SENILAI"
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
         Begin VB.Label lbNilai 
            Caption         =   "Label1"
            Height          =   270
            Left            =   75
            TabIndex        =   19
            Top             =   240
            Width           =   4200
         End
         Begin VB.Label lbKeterangan 
            Caption         =   "Label1"
            Height          =   270
            Left            =   75
            TabIndex        =   18
            Top             =   510
            Width           =   4200
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame1 
         Height          =   1305
         Left            =   7050
         Top             =   180
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   2302
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
         Begin BiSANumberBoxProject.BiSANumberBox nAwal 
            Height          =   330
            Left            =   360
            TabIndex        =   15
            Top             =   120
            Width           =   3675
            _ExtentX        =   6482
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
            Caption         =   "SALDO AWAL"
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
         Begin BiSANumberBoxProject.BiSANumberBox nMutasi 
            Height          =   330
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   3675
            _ExtentX        =   6482
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
            Caption         =   "MUTASI"
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
         Begin BiSANumberBoxProject.BiSANumberBox nAkhir 
            Height          =   330
            Left            =   360
            TabIndex        =   17
            Top             =   840
            Width           =   3675
            _ExtentX        =   6482
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
            Caption         =   "SALDO AKHIR"
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
      End
      Begin BiSATextBoxProject.BiSATextBox cKeteranganTabungan 
         Height          =   330
         Left            =   90
         TabIndex        =   11
         Top             =   2025
         Width           =   6915
         _ExtentX        =   12197
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
         Caption         =   "KETERANGAN"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningJurnal 
         Height          =   330
         Left            =   3585
         TabIndex        =   10
         Top             =   1665
         Width           =   3420
         _ExtentX        =   6033
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
         Left            =   90
         TabIndex        =   6
         Top             =   1305
         Width           =   2130
         _ExtentX        =   3757
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
         Caption         =   "D/K"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaKodeTransaksi 
         Height          =   330
         Left            =   2520
         TabIndex        =   5
         Top             =   945
         Width           =   3420
         _ExtentX        =   6033
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
      Begin BiSATextBoxProject.BiSABrowse cKodeTransaksi 
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   945
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "KODE TRANS."
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoMinimum 
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   585
         Width           =   3270
         _ExtentX        =   5768
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
         Caption         =   "SALDO MIN."
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGolTabungan 
         Height          =   330
         Left            =   2220
         TabIndex        =   1
         Top             =   225
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
      Begin BiSATextBoxProject.BiSATextBox cGolTabungan 
         Height          =   330
         Left            =   90
         TabIndex        =   0
         Top             =   225
         Width           =   2100
         _ExtentX        =   3704
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
         Caption         =   "GOL TABUNGAN"
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
      Begin BiSANumberBoxProject.BiSANumberBox nSetoranMinimum 
         Height          =   330
         Left            =   3390
         TabIndex        =   3
         Top             =   585
         Width           =   3270
         _ExtentX        =   5768
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
         Caption         =   "SETORAN MIN."
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
         Height          =   1455
         Left            =   60
         TabIndex        =   8
         Top             =   2385
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   2566
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
         Columns(1).Caption=   "Nomor Transaksi"
         Columns(1).DataField=   "Faktur"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Keterangan"
         Columns(2).DataField=   "Rekening"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "DK"
         Columns(3).DataField=   "Debet"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Jumlah"
         Columns(4).DataField=   "Kredit"
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=820"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4154"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4075"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=8811"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=8731"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=873"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=794"
         Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(4).Width=4154"
         Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=4075"
         Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=1200,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Times New Roman"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Arial"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin BiSATextBoxProject.BiSABrowse cRekeningJurnal 
         Height          =   330
         Left            =   90
         TabIndex        =   9
         Top             =   1665
         Width           =   3495
         _ExtentX        =   6165
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
         Button          =   -1  'True
         Caption         =   "KODE TRANS."
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
      Begin BiSANumberBoxProject.BiSANumberBox nTotDebet 
         Height          =   330
         Left            =   1800
         TabIndex        =   12
         Top             =   3870
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "DB"
         CaptionWidth    =   300
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
      Begin BiSANumberBoxProject.BiSANumberBox nTotKredit 
         Height          =   330
         Left            =   4275
         TabIndex        =   13
         Top             =   3870
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "CR"
         CaptionWidth    =   300
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoTeller 
         Height          =   330
         Left            =   7020
         TabIndex        =   14
         Top             =   3855
         Width           =   4245
         _ExtentX        =   7488
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
         ForeColor       =   255
         Caption         =   "TOTAL SALDO TELLER"
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
      Begin VB.Label Label1 
         Caption         =   "[K] = Setoran    [D] = Penarikan"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   1365
         Width           =   2775
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame5 
      Height          =   510
      Left            =   0
      Top             =   4290
      Width           =   11520
      _ExtentX        =   20320
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
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9300
         TabIndex        =   20
         Top             =   45
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
         Picture         =   "trMutasiTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10380
         TabIndex        =   21
         Top             =   45
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
         Picture         =   "trMutasiTabungan.frx":0416
      End
   End
End
Attribute VB_Name = "trMutasiTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As Single
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New BisaMyDLL.data
Dim lStatusBlokir As Boolean
Dim nJumlahBlokir As Double
Dim cSql As String
Dim vaarray As New XArrayDB

Private Sub cKodeTransaksi_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "KodeTransaksi k", "k.Kode", cKodeTransaksi, "k.Kode,k.Keterangan,k.DK,k.Kas,k.Rekening", " and (t.Level > 0 or " & nUserLevel & " = 0)", _
               Array("Left Join KodetransaksiTeller t on k.Kode = t.Kode and Level = " & nUserLevel))
  If Not dbData.eof Then
    cNamaKodeTransaksi.Text = GetNull(dbData!Keterangan)
    cDK.Text = GetNull(dbData!DK)
    cRekeningJurnal.Default
    cNamaRekeningJurnal.Default
    
    ' Ambil Rekening Jurnal
    cRekeningJurnal.Text = IIf(GetNull(dbData!Kas) = "K", cKasTeller, GetNull(dbData!Rekening))
    Set dbData = objData.Browse(GetDSN, "Rekening", , "Kode", sisAssign, cRekeningJurnal.Text)
    If Not dbData.eof Then
      cNamaRekeningJurnal.Text = GetNull(dbData!Keterangan)
    End If
    cKeteranganTabungan.Text = cNamaKodeTransaksi.Text & " a.n " & trTeller.cNama.Text
  End If
End Sub

Private Sub cKodeTransaksi_Validate(Cancel As Boolean)
  If cKodeTransaksi.LastKey = 13 Then
     cKodeTransaksi_ButtonClick
  End If
End Sub

Private Sub Form_Activate()
Dim cRekening As String
  BISAPESAN.Visible = False
  Me.Top = 2300
  Me.left = 0
  Me.Width = 11623
  GetDataGolonganTabungan trTeller.cGolongan.Text, cGolTabungan, True, cNamaGolTabungan, True, _
                          , , nSetoranMinimum, True, nSaldoMinimum, True
  cRekening = SetNomorRekening(trTeller.cCabang.Text, trTeller.cGolongan.Text, trTeller.cUrut.Text, trTeller.cFrekuensi.Text)
  Set dbData = objData.Browse(GetDSN, "TABUNGAN", "StatusBlokir,JumlahBlokir,KeteranganBlokir", "Rekening", sisAssign, cRekening)
  If Not dbData.eof Then
    If GetNull(dbData!StatusBlokir) = "1" Then
      lStatusBlokir = True
      nJumlahBlokir = GetNull(dbData!JumlahBlokir)
      BISAPESAN.Visible = True
      lbNilai.Caption = "Rp " & Format(nJumlahBlokir, "###,###,###,###,##0.00")
      lbKeterangan.Caption = GetNull(GetNull(dbData!KeteranganBlokir), "")
      
    End If
  End If
  
End Sub

Private Sub cmdKeluar_Click()
  InitTeller
  Me.Hide
End Sub

Private Sub InitTeller()
  With trTeller
    .Image1.Picture = LoadPicture(GetPicture(""))
    .Image2.Picture = LoadPicture(GetPicture(""))
    '.Height = 2745
    .cShow.Text = "0"
    .cGolongan.Text = ""
    .cUrut.Default
    .cFrekuensi.Default
    .cAlamat.Default
    .cNama.Default
    .cFaktur.Default
    .dTgl.Value = Date
    .cGolongan.SetFocus
    BISAPESAN.Visible = False
  End With
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim cFakturTabungan As String
Dim cRek As String
Dim cRekAdmin As String
Dim cRekBunga As String
Dim cRekPajak As String
Dim cRekPembulatan As String
Dim cRekPenarikan As String
Dim cKodeAdmin As String
Dim cKodeBunga As String
Dim cKodePajak As String
Dim cKodePembulatan As String
Dim cKodePenarikan As String

  If ValidSaving() Then
    If MsgBox("Apakah Data Benar-benar sudah Valid ?", vbYesNo + vbInformation, "Transaksi Mutasi Tabungan") = vbYes Then
       vaField = Array("Faktur", "Tgl", "KodeTransaksi", "Rekening", "Jumlah", "UserName", "DateTime", "Keterangan")
       cFakturTabungan = GetFakturTabungan(objData, trTeller.cCabang.Text, cUserID, trTeller.dTgl.Value)
       cRek = SetNomorRekening(trTeller.cCabang.Text, trTeller.cGolongan.Text, trTeller.cUrut.Text, trTeller.cFrekuensi.Text)
       
'       'simpan di MutasiTabungan
'        If trTeller.OptTransaksi(1).Value = True Then
'
'           'Update Data Tabungan Untuk Memberi Status Tutup
'           objdata.Edit GetDSN, "Tabungan", "Rekening='" & cRek & "'", Array("Awal", "LastUpdate", "Close", "TglPenutupan"), Array(0, trTeller.dTgl.Value, "1", trTeller.dTgl.Value)
'
'           GetKode aCfg(msKodeAdministrasi), cKodeAdmin, cRekAdmin
'           GetKode aCfg(msKodeBunga), cKodeBunga, cRekBunga
'           GetKode aCfg(msPajakBungaTabungan), cKodePajak, cRekPajak
'           GetKode aCfg(msKodePembulatankas), cKodePembulatan, cRekPembulatan
'           GetKode aCfg(msKodePenarikanTunai), cKodePenarikan, cRekPenarikan
'
'           UpdMutasiTabungan objdata, cKodeAdmin, cFakturTabungan, trTeller.dTgl.Value, cRek, nAdministrasi.Value, , "Admin. Tutup Rek Tabungan an. " & trTeller.cNama.Text, True, "D", cRekAdmin, Now
'           UpdMutasiTabungan objdata, cKodeBunga, cFakturTabungan, trTeller.dTgl.Value, cRek, nTotalBunga.Value, , "Bunga Tabungan an. " & trTeller.cNama.Text, True, "K", cRekBunga, Now
'           UpdMutasiTabungan objdata, cKodePajak, cFakturTabungan, trTeller.dTgl.Value, cRek, nTotalPajak.Value, , "Pajak bunga Tabungan an. " & trTeller.cNama.Text, True, "K", cRekPajak, Now
'           UpdMutasiTabungan objdata, cKodePembulatan, cFakturTabungan, trTeller.dTgl.Value, cRek, nPembulatan.Value, , "Pembulatan Kas an. " & trTeller.cNama.Text, True, "D", cRekPembulatan, Now
'           UpdMutasiTabungan objdata, cKodePenarikan, cFakturTabungan, trTeller.dTgl.Value, cRek, nRealPenarikan.Value, , "Penarikan Tutup Tabungan an. " & trTeller.cNama.Text, True, "K", cRekPenarikan, Now
'           trTeller.OptTransaksi(0).Value = True
'           UpdUrutFaktur objdata, cFakturTabungan
'        Else
'          UpdMutasiTabungan objdata, cKodeTransaksi.Text, cFakturTabungan, trTeller.dTgl.Value, cRek, nMutasi.Value, True, cKeteranganTabungan.Text, , cDK.Text, cRekeningJurnal.Text
'          UpdUrutFaktur objdata, cFakturTabungan
'        End If
        
        UpdMutasiTabungan objData, cKodeTransaksi.Text, cFakturTabungan, trTeller.dTgl.Value, cRek, nMutasi.Value, True, cKeteranganTabungan.Text, , cDK.Text, cRekeningJurnal.Text
        UpdUrutFaktur objData, cFakturTabungan
        
'        If MsgBox("Akan mencetak Validasi Tabungan ?", vbYesNo, "Transaksi Mutasi Tabungan") = vbYes Then
'          CetakValidasiTabungan cFakturTabungan, trTeller.dTgl.Value, Now, cRek, trTeller.cNama.Text, cKodeTransaksi.Text, cNamaKodeTransaksi.Text, cDK.Text, IIf(trTeller.OptTransaksi(1).Value = True, nRealPenarikan.Value, nMutasi.Value)
'        End If
'
'        If MsgBox("Akan mencetak Buku Tabungan ?", vbYesNo, "Transaksi Mutasi Tabungan") = vbYes Then
'         CetakValidasiTabungan1 cFakturTabungan, trTeller.dTgl.Value, Now, cRek, trTeller.cNama.Text, cKodeTransaksi.Text, cNamaKodeTransaksi.Text, cDK.Text, IIf(trTeller.OptTransaksi(1).Value = True, nRealPenarikan.Value, nMutasi.Value)
'        End If
        
        
'        If MsgBox("Akan dicetak diBuku Tabungan ?", vbYesNo, "Transaksi Mutasi Tabungan") = vbYes Then
'          Load trCetakMutasiTabungan
'            With trCetakMutasiTabungan
'              .cCabang.Text = trTeller.cCabang.Text
'              .cGolongan.Text = trTeller.cGolongan.Text
'              .cUrut.Text = trTeller.cUrut.Text
'              .cFrekuensi.Text = trTeller.cFrekuensi.Text
'
'              .Show vbModal
'            End With
'        End If
        'BiSAFrame3.Visible = False
        'sisFrame2.Enabled = True
        GetSQL
        
    End If
    Initvalue
    InitTeller
  End If
End Sub

Private Sub GetKode(ByVal cDefault, cKD As String, cKT As String)
  cKD = cDefault
  Set dbData = objData.Browse(GetDSN, "KodeTransaksi", "Kode,Rekening", "Kode", sisAssign, cDefault)
  If Not dbData.eof Then
    cKT = GetNull(dbData!Rekening)
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  If Not CheckData(cKodeTransaksi.Text, "Kode Transaksi Harus Diisi, Ulangi Pengisian..") Then
    ValidSaving = False
    cKodeTransaksi.SetFocus
    Exit Function
  End If
  
  If Not CheckData(nMutasi.Value, "Nilai Mutasi tidak boleh nol, Ulangi Pengisian..") Then
    ValidSaving = False
    nMutasi.SetFocus
    Exit Function
  End If
  
  If Not CheckData(trTeller.cGolongan.Text, "Golongan Pada Nomor Rekening Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    trTeller.cGolongan.SetFocus
    Exit Function
  End If
  
  If Not CheckData(trTeller.cUrut.Text, "Nomor urut Pada Nomor Rekening Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    trTeller.cUrut.SetFocus
    Exit Function
  End If
      
  If Not CheckData(trTeller.cFrekuensi.Text, "Frekuensi Pada Nomor Rekening Harus Diisi, Ulangi Pengisian.....!") Then
    ValidSaving = False
    trTeller.cFrekuensi.SetFocus
    Exit Function
  End If
End Function

Private Sub Initvalue()
  'Tabungan
  cGolTabungan.Default
  cNamaGolTabungan.Default
  nSaldoMinimum.Value = 0
  nSetoranMinimum.Value = 0
  cKodeTransaksi.Default
  cNamaKodeTransaksi.Default
  cDK.Default
  cRekeningJurnal.Default
  cNamaRekeningJurnal.Default
  nAwal.Value = 0
  nMutasi.Value = 0
  nAkhir.Value = 0
  cKeteranganTabungan.Default
End Sub

Private Sub Form_Load()
Dim n As Single

    GetSQL
    Initvalue
    InitGrid TDBGrid1
    TabIndex cKodeTransaksi, n
    TabIndex cKeteranganTabungan, n
    TabIndex nMutasi, n
    TabIndex cmdSimpan, n
    TabIndex cmdKeluar, n
End Sub

Private Sub nMutasi_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nNilaiAkhir As Double

  If KeyCode = 13 Or KeyCode = 40 Then
     If nMutasi.Value = 0 Or nMutasi.Value < 0 Then
        MsgBox "Nilai Mutasi tidak boleh 0 atau lebih kecil 0", vbOKOnly + vbInformation, Me.Caption
        nMutasi.SetFocus
        Exit Sub
     End If
    
     'Setoran Tunai
     If cDK.Text = "K" And nMutasi.Value < nSetoranMinimum.Value And nMutasi.Value <> 0 Then
        MsgBox "Maaf, Setoran Tabungan Minimal : Rp. " & Format((nSetoranMinimum.Value), "#,##,###.00"), vbInformation, Me.Caption
        nMutasi.SetFocus
        Exit Sub
     End If
    
     nNilaiAkhir = nAwal.Value + IIf(cDK.Text = "K", nMutasi.Value, -nMutasi.Value)
     'Penarikan tunai
     If lStatusBlokir = True And cDK.Text = "D" Then
        If nNilaiAkhir < nJumlahBlokir + nSaldoMinimum.Value Then
           MsgBox "Maaf, Saldo Tabungan Anda Tidak Cukup. Silahkan Mengulangi Pengisian !", vbOKOnly + vbInformation, Me.Caption
           nAkhir.Value = 0
           nMutasi.SetFocus
           Exit Sub
        End If
    Else
        If (nNilaiAkhir < nSaldoMinimum.Value) Or nNilaiAkhir < 0 Then
            MsgBox "Maaf, Penarikan tidak boleh melebihi SALDO MINIMUM. Silahkan Mengulangi Pengisian !", vbOKOnly + vbInformation, Me.Caption
            nAkhir.Value = 0
            nMutasi.SetFocus
            Exit Sub
        End If
    End If
    
    nAkhir.Value = nAwal.Value + IIf(cDK.Text = "K", nMutasi.Value, -nMutasi.Value)
  End If
End Sub

'Tabungan
Private Sub GetSQL()
Dim n As Long
Dim nSaldo As Double
Dim cSql As String

  nTotDebet.Value = 0
  nTotKredit.Value = 0
  objData.OpenConnection GetDSN
  cSql = cSql & "Select Awal From SaldoRekening Where Rekening = '" & cKasTeller & "' "
  cSql = cSql & " Union "
  cSql = cSql & "Select Sum(b.Debet-b.Kredit) as Awal From BukuBesar b Where b.Tgl < '" & Format(Date, "yyyy-mm-dd") & "' and b.Rekening = '" & cKasTeller & "'"
  Set dbData = objData.SQL(GetDSN, cSql)
  If Not dbData.eof Then
    nSaldo = nSaldo + GetNull(dbData!Awal)
  End If
  
  Set dbData = objData.Browse(GetDSN, "BukuBesar", "Faktur,Keterangan,Debet,Kredit", "Tgl", sisAssign, Format(Date, "yyyy-mm-dd"), " and Rekening = '" & cKasTeller & "'", "Tgl,Rekening,ID")
               
  objData.CloseConnection GetDSN
  vaarray.ReDim 0, 0, 0, 4
  vaarray(n, 2) = "Saldo Awal Teller"
  vaarray(n, 3) = IIf(nSaldo >= 0, "D", "K")
  vaarray(n, 4) = nSaldo
  
  n = 1
  If Not dbData.eof Then
    Do While Not dbData.eof
        vaarray.InsertRows n
        vaarray(n, 0) = (n)
        vaarray(n, 1) = GetNull(dbData!Faktur)
        vaarray(n, 2) = GetNull(dbData!Keterangan)
        If dbData!Debet <> 0 Then
          vaarray(n, 3) = "D"
          vaarray(n, 4) = GetNull(dbData!Debet)
          nTotDebet.Value = nTotDebet.Value + GetNull(dbData!Debet)
        Else
          vaarray(n, 3) = "K"
          vaarray(n, 4) = GetNull(dbData!Kredit)
          nTotKredit.Value = nTotKredit.Value + GetNull(dbData!Kredit)
        End If
        n = n + 1
      
      dbData.MoveNext
    Loop
  End If

  nSaldoTeller.Value = nTotDebet.Value - nTotKredit.Value + nSaldo
  nSaldoTeller.ForeColor = IIf(nSaldoTeller.Value < 0, &HFF&, &H80000008)
  Set TDBGrid1.Array = vaarray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  If Value = 0 Then
    Value = ""
  Else
    Value = Format(Value, "###,###,###,###,##0.00")
  End If
End Sub

