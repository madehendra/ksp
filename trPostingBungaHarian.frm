VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPostingBungaHarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posting Bunga Harian"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9825
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   690
      Left            =   30
      Top             =   6120
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   1217
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Max             =   1
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   8640
         TabIndex        =   2
         Top             =   120
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
         Picture         =   "trPostingBungaHarian.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPosting 
         Height          =   435
         Left            =   7470
         TabIndex        =   3
         Top             =   120
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
         Picture         =   "trPostingBungaHarian.frx":059A
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   6120
      Left            =   30
      Top             =   15
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   10795
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
         Height          =   360
         Left            =   105
         TabIndex        =   1
         Top             =   255
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   635
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
         Caption         =   "Tgl"
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
      Begin TrueOleDBGrid70.TDBGrid DataGrid1 
         Height          =   5325
         Left            =   15
         TabIndex        =   0
         Top             =   780
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   9393
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
         Columns(2).Caption=   "Saldo Tgl"
         Columns(2).DataField=   "FullName"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Saldo"
         Columns(3).DataField=   "KasTeller"
         Columns(3).NumberFormat=   "###,###,###,###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Bunga Pa."
         Columns(4).DataField=   "Keterangan"
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Jumlah Bunga"
         Columns(5).DataField=   "Plafond"
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Tgl"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "dd-MM-yyyy"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2858"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2778"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3387"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3307"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2328"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2249"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2805"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2725"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2302"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2223"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=3149"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=3069"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
         _StyleDefs(14)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
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
         _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
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
      Begin BiSADateProject.BiSADate dTgl2 
         Height          =   360
         Left            =   2580
         TabIndex        =   4
         Top             =   255
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
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
         Caption         =   "sd"
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
End
Attribute VB_Name = "trPostingBungaHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub PostingBungaHarian()
Dim n, i As Integer
Dim nTotal As Double
Dim nDate As Integer
Dim dTglProses As Date

  nTotal = 0
  vaArray.ReDim 0, -1, 0, 6
  nDate = DateDiff("d", Format(dTgl.Value, "yyyy-MM-dd"), Format(dTgl2.Value, "yyyy-MM-dd"))
  ProgressBar1.Visible = True
  ProgressBar1.Min = 0
  ProgressBar1.Max = IIf(nDate = 0, 1, nDate)
  For i = 0 To nDate
  ProgressBar1.Value = i
    dTglProses = DateAdd("d", i, dTgl.Value)
    Set dbData = objData.Browse(GetDSN, "tabungan t", "t.rekening as rekeningtabungan,r.nama,t.tgl,g.bunga", "g.jenisbunga", sisAssign, "3", , , Array("left join golongantabungan g on g.kode = t.golongantabungan", "left join registernasabah r on r.kode = t.kode"))
    If Not dbData.eof Then
      FrmPB.InitPB dbData.RecordCount
      FrmPB.Caption = "Memproses tgl " & Format(DateAdd("d", i, dTgl.Value), "dd/MM/yyyy")
      Do While Not dbData.eof
        FrmPB.RunPB
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = GetNull(dbData!RekeningTabungan)
        vaArray(n, 1) = GetNull(dbData!nama)
'        vaArray(n, 2) = Format(DateAdd("d", -1, dTglProses), "dd/MM/yyyy")
'        vaArray(n, 3) = GetSaldoTab(objData, GetNull(dbData!RekeningTabungan), DateAdd("d", -1, dTglProses))
        vaArray(n, 2) = Format(dTglProses, "dd/MM/yyyy")
        vaArray(n, 3) = GetSaldoTab(objData, GetNull(dbData!RekeningTabungan), dTglProses)
        
        vaArray(n, 4) = Format(GetNull(dbData!bunga), "###,###,##0.00")
        vaArray(n, 5) = PerhitunganBungaHarian(vaArray(n, 4), vaArray(n, 3))
        vaArray(n, 6) = dTglProses
        nTotal = nTotal + vaArray(n, 5)
        If vaArray(n, 5) <= 0 Then
          vaArray.DeleteRows n
        End If
        dbData.MoveNext
      Loop
      FrmPB.EndPB
    End If
  Next i
  ProgressBar1.Visible = False
  
  
  Set DataGrid1.Array = vaArray
  DataGrid1.Columns(5).FooterText = Format(nTotal, "###,###,##0.00")
  DataGrid1.ReBind
  DataGrid1.Refresh
  
End Sub

Private Function PerhitunganBungaHarian(ByVal nBungaPA As Double, ByVal nSaldoSimpanan As Double) As Double
  PerhitunganBungaHarian = 0
  PerhitunganBungaHarian = (nSaldoSimpanan * (nBungaPA / 100)) / 365
End Function

Private Sub cmdPosting_Click()
  
  PostingBungaHarian
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim n As Integer

  objData.Delete GetDSN, "bungaharian", "tgl", sisAssign, Format(dTgl.Value, "yyyy-MM-dd")
  FrmPB.InitPB vaArray.UpperBound(1)
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    FrmPB.RunPB
    vaField = Array("rekeningsimpanan", "tglposting", "tgl", "saldoakhir", "bunga", "pa", "username", "datetime")
    vaValue = Array(vaArray(n, 0), Date, Format(vaArray(n, 6), "yyyy-MM-dd"), vaArray(n, 3), vaArray(n, 5), vaArray(n, 4), GetRegistry(reg_UserName), SNow)
    objData.Add GetDSN, "bungaharian", vaField, vaValue
    
  Next n
  FrmPB.EndPB
  MsgBox "Data sudah disimpan!!"
End Sub

Private Sub Form_Load()
  CenterForm Me
End Sub


