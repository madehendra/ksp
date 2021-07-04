VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form MstGolonganTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER GOLONGAN SIMPANAN"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   3105
      Left            =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5477
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
      Begin VB.OptionButton OptBunga 
         Caption         =   "&3 Harian"
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
         Index           =   2
         Left            =   4170
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1140
      End
      Begin VB.OptionButton OptBunga 
         Caption         =   "&1 Progressif"
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
         Index           =   0
         Left            =   1770
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1275
         Width           =   1335
      End
      Begin VB.OptionButton OptBunga 
         Caption         =   "&2 Biasa"
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
         Index           =   1
         Left            =   3120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1140
      End
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoMinimum 
         Height          =   330
         Left            =   7260
         TabIndex        =   2
         Top             =   105
         Width           =   4395
         _ExtentX        =   7752
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
         Caption         =   "SALDO MIN. MENGENDAP"
         CaptionWidth    =   2300
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningBunga 
         Height          =   330
         Left            =   3390
         TabIndex        =   3
         Top             =   2445
         Width           =   3645
         _ExtentX        =   6429
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
      Begin BiSATextBoxProject.BiSATextBox cKetSukuBunga 
         Height          =   330
         Left            =   2790
         TabIndex        =   4
         Top             =   2055
         Width           =   4245
         _ExtentX        =   7488
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
         Left            =   3390
         TabIndex        =   5
         Top             =   855
         Width           =   3645
         _ExtentX        =   6429
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
         Left            =   150
         TabIndex        =   6
         Top             =   855
         Width           =   3225
         _ExtentX        =   5689
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
         Caption         =   "REKENING"
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
         Left            =   150
         TabIndex        =   7
         Top             =   135
         Width           =   1920
         _ExtentX        =   3387
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
         Caption         =   "KODE              T"
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   150
         TabIndex        =   8
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
      Begin BiSATextBoxProject.BiSABrowse cSukuBunga 
         Height          =   330
         Left            =   150
         TabIndex        =   9
         Top             =   2055
         Width           =   2610
         _ExtentX        =   4604
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
         Caption         =   "BUNGA PROGR."
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningBunga 
         Height          =   330
         Left            =   150
         TabIndex        =   10
         Top             =   2445
         Width           =   3225
         _ExtentX        =   5689
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
         Caption         =   "REK. BY BUNGA"
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
         Left            =   7260
         TabIndex        =   11
         Top             =   465
         Width           =   4395
         _ExtentX        =   7752
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
         Caption         =   "SETORAN MINIMUM"
         CaptionWidth    =   2300
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoDapatBunga 
         Height          =   330
         Left            =   7260
         TabIndex        =   12
         Top             =   840
         Width           =   4395
         _ExtentX        =   7752
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
         Caption         =   "SALDO MIN. DPT BUNGA"
         CaptionWidth    =   2300
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
         Left            =   7260
         TabIndex        =   13
         Top             =   1215
         Width           =   4395
         _ExtentX        =   7752
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
         Caption         =   "ADM. TUTUP SIMPANAN"
         CaptionWidth    =   2300
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoKenaPajak 
         Height          =   330
         Left            =   7260
         TabIndex        =   14
         Top             =   1590
         Width           =   4395
         _ExtentX        =   7752
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
         Caption         =   "SALDO MIN. KENA PAJAK"
         CaptionWidth    =   2300
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
         Left            =   7260
         TabIndex        =   15
         Top             =   1965
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
         Caption         =   "PAJAK BUNGA (%)"
         CaptionWidth    =   2300
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
         Left            =   150
         TabIndex        =   16
         Top             =   1665
         Width           =   2565
         _ExtentX        =   4524
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
         Caption         =   "BUNGA (%) p.a"
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
      Begin VB.Label Label1 
         Caption         =   "JENIS BUNGA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   17
         Top             =   1290
         Width           =   1590
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3600
      Left            =   0
      Top             =   3090
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6350
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
         Height          =   3465
         Left            =   60
         TabIndex        =   18
         Top             =   75
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   6112
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
         Columns(2).Caption=   "REKENING"
         Columns(2).DataField=   "REKENING"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1879"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1799"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=14235"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=14155"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3836"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3757"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
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
         DefColWidth     =   0
         HeadLines       =   1.5
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   6675
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
         TabIndex        =   19
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
         Picture         =   "MstGolonganTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3825
         TabIndex        =   20
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
         Picture         =   "MstGolonganTabungan.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9510
         TabIndex        =   21
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
         Picture         =   "MstGolonganTabungan.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1605
         TabIndex        =   22
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
         Picture         =   "MstGolonganTabungan.frx":083F
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   540
         TabIndex        =   23
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
         Picture         =   "MstGolonganTabungan.frx":096B
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10590
         TabIndex        =   24
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
         Picture         =   "MstGolonganTabungan.frx":0B16
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   105
         TabIndex        =   25
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
         Picture         =   "MstGolonganTabungan.frx":0BBC
      End
   End
End
Attribute VB_Name = "MstGolonganTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim dbRekening As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim lEdit As Boolean
Dim nPos As SisPos

Private Sub cKode_Validate(Cancel As Boolean)
  If Not dbData.eof Then dbData.MoveFirst
  dbData.Find "Kode = 'T" & cKode.Text & "'"
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
     objData.Delete GetDSN, "GolonganTabungan", "kode", sisAssign, "T" & cKode.Text
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

Private Sub cmdSimpan_Click()
Dim vaField, vaValue

  If ValidSaving() Then
    If MsgBox("Data benar-benar sudah VALID ?'", vbYesNo + vbInformation) = vbYes Then
      vaField = Array("Kode", "Keterangan", "Rekening", "SukuBunga", "Bunga", "RekeningBunga", "SaldoMinimum", "SetoranMinimum", "SaldominimumDapatBunga", "AdministrasiTutup", "SaldoMinimumKenaPajak", "PajakBunga", "JenisBunga")
      vaValue = Array("T" & cKode.Text, cKeterangan.Text, cRekening.Text, cSukuBunga.Text, nBunga.Value, cRekeningBunga.Text, nSaldoMinimum.Value, nSetoranMinimum.Value, nSaldoDapatBunga.Value, nAdministrasi.Value, nSaldoKenaPajak.Value, nPajak.Value, GetOpt(OptBunga))
      objData.Update GetDSN, "GolonganTabungan", "Kode = 'T" & cKode.Text & "'", vaField, vaValue
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
  
  If Not CheckData(cKode.Text, "Kode GolonganTabungan Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKeterangan.Text, "Nama GolonganTabungan Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
  
  ' Check Rekening Perkiraan
  If Not CheckData(cRekening.Text, "Rekening Perkiraan Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cRekening.SetFocus
    Exit Function
  End If
End Function

Private Sub cRekening_ButtonClick()
  Set dbRekening = objData.Pick(GetDSN, "Rekening", "Kode", cRekening, "Kode,Keterangan,Jenis", " and jenis = 'D'")
  If Not dbRekening.eof Then
    cNamaRekening.Text = GetNull(dbRekening!Keterangan, "")
  End If
End Sub

Private Sub cRekening_Validate(Cancel As Boolean)
  If cRekening.LastKey = 13 Then
    cRekening_ButtonClick
  End If
End Sub

Private Sub cRekeningBunga_ButtonClick()
  Set dbRekening = objData.Pick(GetDSN, "Rekening", "Kode", cRekeningBunga, "Kode,Keterangan,Jenis", " and jenis ='D'")
  If Not dbRekening.eof Then
    cNamaRekeningBunga.Text = GetNull(dbRekening!Keterangan, "")
  End If
End Sub

Private Sub cRekeningBunga_Validate(Cancel As Boolean)
  If cRekeningBunga.LastKey = 13 Then
    cRekeningBunga_ButtonClick
  End If
End Sub

Private Sub GetMemory()
Dim vaJoin
Dim cField As String
Dim db As New ADODB.Recordset

  cField = "k.Keterangan,k.Rekening,k.SukuBunga,k.Bunga,k.SaldoMinimum,k.SetoranMinimum,k.SaldoMinimumDapatBunga,"
  cField = cField & "k.AdministrasiTutup,k.SaldoMinimumKenaPajak,k.PajakBunga,k.rekeningBunga,"
  cField = cField & "k.jenisBunga,i.Keterangan as NamaRekeningBunga,r.Keterangan as KeteranganRekening,s.Keterangan as KeteranganSukuBunga,"
  cField = cField & "k.rekeningadministrasi, a.keterangan as namarekeningadministrasi"
  vaJoin = Array("Left Join Rekening r on r.Kode=k.Rekening", _
                 "Left Join Rekening i on i.Kode = k.RekeningBunga", _
                 "Left Join rekening a on a.kode = k.rekeningadministrasi", _
                 "Left Join SukuBunga s on s.Kode=k.SukuBunga")
  Set db = objData.Browse(GetDSN, "GolonganTabungan k", cField, "k.Kode", sisAssign, "T" & cKode.Text, , "k.Kode", vaJoin)
  If Not db.eof Then
     cKeterangan.Text = GetNull(db!Keterangan, "")
     cNamaRekening.Text = GetNull(db!KeteranganRekening, "")
     cRekening.Text = GetNull(db!Rekening, "")
     cSukuBunga.Text = GetNull(db!SukuBunga)
     cKetSukuBunga.Text = GetNull(db!KeteranganSukuBunga, "")
     nSaldoMinimum.Value = GetNull(db!SaldoMinimum)
     nSetoranMinimum.Value = GetNull(db!SetoranMinimum)
     nSaldoDapatBunga.Value = GetNull(db!SaldoMinimumDapatBunga)
     nAdministrasi.Value = GetNull(db!administrasitutup)
     nSaldoKenaPajak.Value = GetNull(db!SaldoMinimumKenaPajak)
     nPajak.Value = GetNull(db!pajakbunga)
     cRekeningBunga.Text = GetNull(db!Rekeningbunga)
     cNamaRekeningBunga.Text = GetNull(db!NamaRekeningBunga)
     nBunga.Value = GetNull(db!bunga)
     SetOpt OptBunga, GetNull(db!JenisBunga, "2")
  End If
End Sub

Private Sub cSukuBunga_ButtonClick()
  Set dbData = objData.Pick(GetDSN, "SukuBunga", "Kode", cSukuBunga, "Kode,Keterangan")
  If Not dbData.eof Then
    cKetSukuBunga.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cSukuBunga_Validate(Cancel As Boolean)
  If cSukuBunga.LastKey = 13 Then
    cSukuBunga_ButtonClick
  End If
End Sub

Private Sub DataGrid1_Click()
  lClick = True
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Then
    lClick = True
  End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If lClick And Not lEdit Then
    If Not dbData.eof Then
      cKode.Default
      cKode.Text = Right(DataGrid1.Columns(0), 1)
      GetMemory
    End If
  End If
  lClick = False
End Sub

Private Sub initvalue()
  OptBunga(0).Value = True
  cKode.Default
  cKeterangan.Default
  cRekening.Default
  cNamaRekening.Default
  nSaldoMinimum.Value = 0
  nSetoranMinimum.Value = 0
  nSaldoDapatBunga.Value = 0
  nAdministrasi.Value = 0
  cSukuBunga.Default
  cKetSukuBunga.Default
  nBunga.Value = 0
  nSaldoKenaPajak.Value = 0
  nPajak.Value = 0
  cRekeningBunga.Default
  cNamaRekeningBunga.Default
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me, True
  GetEdit False
  GetSQL
  initvalue
  
  TabIndex cKode, n
  TabIndex cKeterangan, n
  TabIndex cRekening, n
  TabIndex OptBunga(0), n
  TabIndex OptBunga(1), n
  TabIndex OptBunga(2), n
  TabIndex nBunga, n
  TabIndex cSukuBunga, n
  TabIndex cRekeningBunga, n
  TabIndex nSaldoMinimum, n
  TabIndex nSetoranMinimum, n
  TabIndex nSaldoDapatBunga, n
  TabIndex nAdministrasi, n
  TabIndex nSaldoKenaPajak, n
  TabIndex nPajak, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdPreview, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub GetSQL()
  Set dbData = objData.Browse(GetDSN, "GolonganTabungan", , , , , , "Kode")
  If Not dbData.eof Then
    dbData.MoveFirst
  End If
  Set DataGrid1.DataSource = dbData
End Sub

Private Sub OptBunga_Click(Index As Integer)
  If Index = 0 Then
    cSukuBunga.Enabled = True
    cSukuBunga.BackColor = &H80000005
    nBunga.Enabled = False
    nBunga.BackColor = &H8000000F
  Else
    cSukuBunga.Default
    cKetSukuBunga.Default
    nBunga.Enabled = True
    nBunga.BackColor = &H80000005
    cSukuBunga.Enabled = False
    cSukuBunga.BackColor = &H8000000F
  End If
End Sub

Private Sub OptBunga_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub GetDataRpt()
Dim n As Integer
  
  vaArray.ReDim 0, -1, 0, 9
  Set dbData = objData.Browse(GetDSN, "GolonganTabungan", , , , , , "Kode")
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull(dbData!Kode, "")
      vaArray(n, 1) = GetNull(dbData!Keterangan, "")
      vaArray(n, 2) = IIf(GetNull(dbData!JenisBunga, "") = "1", "Progressif", "Biasa")
      vaArray(n, 3) = GetNull(dbData!bunga)
      vaArray(n, 4) = GetNull(dbData!SaldoMinimum)
      vaArray(n, 5) = GetNull(dbData!SetoranMinimum)
      vaArray(n, 6) = GetNull(dbData!SaldoMinimumDapatBunga)
      vaArray(n, 7) = GetNull(dbData!administrasitutup)
      vaArray(n, 8) = GetNull(dbData!SaldoMinimumKenaPajak)
      vaArray(n, 9) = GetNull(dbData!pajakbunga)
      dbData.MoveNext
    Loop
    GetRpt
  End If
End Sub

Private Sub cmdPreview_Click()
  GetDataRpt
End Sub

Private Sub GetRpt()
With FrmRPT
    .AddPageHeader "DAFTAR GOLONGAN TABUNGAN", tdbHalignCenter, , , True, dbArial, 12, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "KODE", , , , 6, , , , , , , , , , , , , 5
    .AddTableHeader "NAMA TABUNGAN"
    .AddTableHeader "JENIS BUNGA", , , , 7
    .AddTableHeader "BUNGA(%)", , , , 7
    .AddTableHeader "SALDO MIN", , , , 10
    .AddTableHeader "SETORAN MIN", , , , 10
    .AddTableHeader "SALDO MIN DPT BUNGA", , , , 11
    .AddTableHeader "ADM TUTUP TABUNGAN", , , , 11
    .AddTableHeader "SALDO MIN KENA PAJAK", , , , 11
    .AddTableHeader "PAJAK BUNGA", , , , 10
    
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .Preview vaArray, True, , True
  End With
End Sub


