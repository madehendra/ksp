VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptRekapitulasiKeseluruhanTabunganPDL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REKAPITULASI KESELURUHAN SIMPANAN"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11535
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   30
      Top             =   5775
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   1138
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10260
         TabIndex        =   3
         Top             =   90
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
         Picture         =   "rptRekapitulasiKeseluruhanTabunganPDL.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   9105
         TabIndex        =   4
         Top             =   90
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
         Picture         =   "rptRekapitulasiKeseluruhanTabunganPDL.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   435
         Left            =   8385
         TabIndex        =   5
         Top             =   90
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   767
         Caption         =   "  OK"
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
         Picture         =   "rptRekapitulasiKeseluruhanTabunganPDL.frx":032C
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   840
      Left            =   30
      Top             =   15
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   1482
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   255
         Width           =   2685
         _ExtentX        =   4736
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
         ForeColor       =   -2147483640
         Caption         =   "Antara Tgl"
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
         Left            =   2835
         TabIndex        =   1
         Top             =   255
         Width           =   2055
         _ExtentX        =   3625
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
         ForeColor       =   -2147483640
         Caption         =   "sd."
         CaptionWidth    =   0
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4920
      Left            =   30
      Top             =   855
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   8678
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4755
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   8387
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode"
         Columns(0).DataField=   "Nama"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama"
         Columns(1).DataField=   "Alamat"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Resi"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Jumlah"
         Columns(3).DataField=   "Kode"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Saldo Lalu"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Setoran"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Penarikan"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Saldo Baru"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3440"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3360"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1693"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1931"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1852"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2937"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2858"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2487"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2408"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2302"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2223"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         ColumnFooters   =   -1  'True
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14,.alignment=2"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14,.alignment=2"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "rptRekapitulasiKeseluruhanTabunganPDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB
Dim lClick As Boolean

Private Sub GetInitData()
Dim n As Integer
Dim nTmpNasabah As Integer
Dim nTmpSaldoLalu As Double
Dim nTmpSetoran As Double
Dim nTmpPenarikan As Double
Dim nTmpSaldoBaru As Double
Dim nTmpResi As Double

  
  vaArray.ReDim 0, -1, 0, 7
  GetUpdateGrid
  Set dbData = objData.Browse(GetDSN, "PDL p", "p.Kode,p.Keterangan,count(t.Rekening) as Nasabah", , , , "1=1 Group by t.PDL", "p.Kode", Array("Left Join Tabungan t on t.PDL = p.Kode"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Kode, "")
      vaArray(n, 1) = GetNull(dbData!Keterangan, "")
      vaArray(n, 2) = GetResi(objData, vaArray(n, 0), "01")
      vaArray(n, 3) = GetNull(dbData!Nasabah)
      vaArray(n, 4) = GetSaldoLalu(objData, vaArray(n, 0))
      vaArray(n, 5) = GetSaldoMutasi(objData, vaArray(n, 0), "01")
      vaArray(n, 6) = GetSaldoMutasi(objData, vaArray(n, 0), "06")
      vaArray(n, 7) = vaArray(n, 4) + vaArray(n, 5) - vaArray(n, 6)
      nTmpNasabah = nTmpNasabah + vaArray(n, 3)
      nTmpSaldoLalu = nTmpSaldoLalu + vaArray(n, 4)
      nTmpSetoran = nTmpSetoran + vaArray(n, 5)
      nTmpPenarikan = nTmpPenarikan + vaArray(n, 6)
      nTmpSaldoBaru = nTmpSaldoBaru + vaArray(n, 7)
      nTmpResi = nTmpResi + vaArray(n, 2)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    TDBGrid1.Columns(2).FooterText = Format(nTmpResi, "###,###,###,###")
    TDBGrid1.Columns(3).FooterText = Format(nTmpNasabah, "###,###,###,###")
    TDBGrid1.Columns(4).FooterText = Format(nTmpSaldoLalu, "###,###,###,###,##0.00")
    TDBGrid1.Columns(5).FooterText = Format(nTmpSetoran, "###,###,###,###,##0.00")
    TDBGrid1.Columns(6).FooterText = Format(nTmpPenarikan, "###,###,###,###,##0.00")
    TDBGrid1.Columns(7).FooterText = Format(nTmpSaldoBaru, "###,###,###,###,##0.00")
    GetUpdateGrid
  End If
  
End Sub

Private Sub GetUpdateGrid()
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
End Sub

Private Function GetSaldoLalu(ByVal obj As CodeSuiteLibrary.data, ByVal PDL As String) As Double
Dim db As New ADODB.Recordset

  GetSaldoLalu = 0
  Set db = obj.Browse(GetDSN, "MutasiTabungan m", "m.Rekening,m.Jumlah,m.DK,m.Jumlah", "m.tgl", sisLT, Format(dTgl(0).Value, "yyyy-MM-dd"), " and p.Kode ='" & PDL & "'", , Array("Left Join Tabungan t on t.Rekening = m.Rekening", "Left Join PDL p on p.Kode = t.PDL"))
  If Not db.eof Then
    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
      FrmPB.RunPB
      GetSaldoLalu = GetSaldoLalu + IIf((db!DK) = "D", -(db!Jumlah), (db!Jumlah))
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Function

Private Function GetSaldoMutasi(ByVal obj As CodeSuiteLibrary.data, ByVal PDL As String, ByVal KodeTransaksi As String) As Double
Dim db As New ADODB.Recordset

  GetSaldoMutasi = 0
  Set db = obj.Browse(GetDSN, "MutasiTabungan m", "m.Rekening,m.Jumlah,m.DK,m.Jumlah,m.KodeTransaksi", "m.tgl", sisLTEqual, Format(dTgl(1).Value, "yyyy-MM-dd"), " and m.tgl >='" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and p.Kode ='" & PDL & "' and m.KodeTransaksi='" & KodeTransaksi & "'", , Array("Left Join Tabungan t on t.Rekening = m.Rekening", "Left Join PDL p on p.Kode = t.PDL"))
  If Not db.eof Then
    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
      FrmPB.RunPB
      GetSaldoMutasi = GetSaldoMutasi + GetNull(db!Jumlah)
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Function


Private Function GetResi(ByVal obj As CodeSuiteLibrary.data, ByVal PDL As String, ByVal KodeTransaksi) As Double
Dim db As New ADODB.Recordset

  GetResi = 0
  Set db = obj.Browse(GetDSN, "MutasiTabungan m", "count(m.Rekening) as Resi", "m.tgl", sisGTEqual, Format(dTgl(0).Value, "yyyy-MM-dd"), " and m.Tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "' and p.Kode ='" & PDL & "' and m.KodeTransaksi = '" & KodeTransaksi & "'", , Array("Left Join Tabungan t on t.Rekening = m.Rekening", "Left Join PDL p on p.Kode = t.PDL"))
  If Not db.eof Then
    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
      FrmPB.RunPB
      GetResi = GetNull(db!Resi)
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Function

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  GetInitData
End Sub

Private Sub cmdPreview_Click()
  If vaArray.UpperBound(1) >= 0 Then
    rpt
  Else
    MsgBox "DATA TIDAK ADA..", vbInformation, Me.Caption
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  InitGrid TDBGrid1
  vaArray.ReDim 0, -1, 0, 6
  lClick = True
  dTgl(0).Value = BOM(dTgl(0).Value)
  dTgl(1).Value = Date
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cmdOK, n
  TabIndex cmdPreview, n
  CenterForm Me, True
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
  If lClick Then
    Select Case ColIndex
      Case 0
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 0, XORDER_ASCEND, XTYPE_STRING
        lClick = Not lClick
      Case 1
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
        lClick = Not lClick
      Case 2
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 2, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 3
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 3, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 4
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 4, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 5
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 5, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 6
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 6, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 7
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 7, XORDER_ASCEND, XTYPE_DOUBLE
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
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 2, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 3
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 4
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 4, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 5
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 5, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 6
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 6, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 7
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 7, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
    End Select
  End If
  TDBGrid1.ReBind
End Sub

Private Sub rpt()
  With FrmRPT
    .AddPageHeader UCase("Laporan Rekapitulasi Keseluruhan Simpanan PDL"), tdbHalignCenter, , , , , 14, True, True
    .AddPageHeader "Antara Tanggal : " & Format(dTgl(0).Value, "dd-MM-yyyy") & " s.d " & Format(dTgl(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , , , , , , , , , , 10
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "PDL", , , , 5
    .AddTableHeader "NAMA"
    .AddTableHeader "RESI", , , , 7
    .AddTableHeader "NASABAH", , , , 10
    .AddTableHeader "SALDO LALU", , , , 15
    .AddTableHeader "SETORAN", , , , 15
    .AddTableHeader "PENARIKAN", , , , 15
    .AddTableHeader "SALDO BARU", , , , 15
           
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number, tdbHalignRight
    .AddTableBody Sis_Rpt_Number, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .AddTableFooter "Total", , tdbHalignCenter, , , , , , , , , , , , 2
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub
