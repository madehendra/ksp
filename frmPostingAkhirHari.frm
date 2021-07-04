VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form frmPostingAkhirHari 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POSTING PERPANJANGAN DEPOSITO ARO"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11490
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5460
      Left            =   0
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   9631
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
         Height          =   4245
         Left            =   105
         TabIndex        =   5
         Top             =   1155
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
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
         Columns(1).Caption=   "JTH TMP"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "REKENING"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NAMA"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "LAMA"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,##0"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "NOMINAL"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "KE-"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,##0"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "JTH TMP"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1905"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1826"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2910"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2831"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=5239"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=5159"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1270"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1191"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2858"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2778"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=979"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=900"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2117"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2037"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=513"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Named:id=33:Normal"
         _StyleDefs(70)  =   ":id=33,.parent=0"
         _StyleDefs(71)  =   "Named:id=34:Heading"
         _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   ":id=34,.wraptext=-1"
         _StyleDefs(74)  =   "Named:id=35:Footing"
         _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   "Named:id=36:Selected"
         _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=37:Caption"
         _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(80)  =   "Named:id=38:HighlightRow"
         _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(82)  =   "Named:id=39:EvenRow"
         _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(84)  =   "Named:id=40:OddRow"
         _StyleDefs(85)  =   ":id=40,.parent=33"
         _StyleDefs(86)  =   "Named:id=41:RecordSelector"
         _StyleDefs(87)  =   ":id=41,.parent=34"
         _StyleDefs(88)  =   "Named:id=42:FilterBar"
         _StyleDefs(89)  =   ":id=42,.parent=33"
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
      Height          =   615
      Left            =   0
      Top             =   5445
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   1085
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
         Left            =   7470
         TabIndex        =   6
         Top             =   150
         Width           =   2430
         _ExtentX        =   4286
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
         BackColor       =   12632256
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
         Height          =   435
         Left            =   10875
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   90
         Width           =   465
         _ExtentX        =   820
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
         Picture         =   "frmPostingAkhirHari.frx":0000
      End
      Begin MSComctlLib.ProgressBar pr 
         Height          =   345
         Left            =   195
         TabIndex        =   2
         Top             =   165
         Visible         =   0   'False
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   609
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin BiSAButtonProject.BiSAButton cmdRefresh 
         Height          =   435
         Left            =   9930
         TabIndex        =   3
         ToolTipText     =   "Refresh"
         Top             =   90
         Width           =   480
         _ExtentX        =   847
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
         Picture         =   "frmPostingAkhirHari.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   435
         Left            =   10425
         TabIndex        =   4
         ToolTipText     =   "Proses"
         Top             =   90
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
         Picture         =   "frmPostingAkhirHari.frx":0250
      End
   End
End
Attribute VB_Name = "frmPostingAkhirHari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaDP As New XArrayDB

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
  If vaDP.UpperBound(1) < 0 Then
    MsgBox "Tidak ada data yang disimpan...", vbOKOnly + vbInformation, "POSTING PERPANJANGAN DEPOSITO ARO"
    cmdKeluar.SetFocus
    Exit Sub
  End If
  
  If MsgBox("Posting akan dilakukan...", vbOKCancel + vbInformation, "POSTING PERPANJANGAN DEPOSITO ARO") = vbOK Then
    ProsesSimpan
    MsgBox "Proses selesai...", vbOKOnly + vbInformation, "POSTING PERPANJANGAN DEPOSITO ARO"
    initvalue
    cmdKeluar.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cmdRefresh_Click()
  GetDeposito
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
  cText = "Posting ini akan melakukan perpanjangan Deposito ARO secara otomatis." & vbCrLf
  cText = cText & "Perpanjangan hanya akan merubah tanggal jatuh tempo"
  Label1.Caption = cText
  CenterForm Me, True
  TabIndex cmdRefresh, n
  TabIndex cmdOK, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetDeposito()
Dim cField As String
Dim cWhere As String
Dim vaJoin
Dim n As Integer
  
  vaDP.ReDim 0, -1, 0, 8
  cField = "d.Rekening,d.Tgl,d.JthTmp,d.NominalDeposito,r.nama,d.Lama,d.jumlahperpanjangan"
  vaJoin = Array("Left Join RegisterNasabah r on r.Kode = d.Kode", _
                 "Left Join GolonganDeposito g on g.Kode = d.GolonganDeposito")
  
  cWhere = "And d.JthTmp <= '" & Format(dTgl.Value, "yyyy-mm-dd") & "'"
  cWhere = cWhere & "And d.SistemARO = 'Y'"
  cWhere = cWhere & "And d.Nominaldeposito > 0"
  
  Set dbData = objData.Browse(GetDSN, "Deposito d", cField, "Status", sisDifference, "1", cWhere, "d.rekening", vaJoin)
  If Not dbData.eof Then
    InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.eof
      RunPB
      vaDP.InsertRows vaDP.UpperBound(1) + 1
      n = vaDP.UpperBound(1)
      
      vaDP(n, 0) = GetNull(dbData!Tgl)
      vaDP(n, 1) = GetNull(dbData!jthtmp)
      vaDP(n, 2) = GetNull(dbData!Rekening, "")
      vaDP(n, 3) = GetNull(dbData!nama, "")
      vaDP(n, 4) = GetNull(dbData!Lama)
      vaDP(n, 5) = GetNull(dbData!nominaldeposito)
      vaDP(n, 6) = DateDiff("m", vaDP(n, 0), vaDP(n, 1)) / vaDP(n, 4)
      vaDP(n, 7) = DateAdd("m", vaDP(n, 4), vaDP(n, 1))
      vaDP(n, 8) = GetNull(dbData!jumlahperpanjangan)
      dbData.MoveNext
    Loop
    EndPB
  End If
  Set TDBGrid2.Array = vaDP
  TDBGrid2.ReBind
End Sub

Private Sub initvalue()
  vaDP.ReDim 0, -1, 0, 8
  Set TDBGrid2.Array = vaDP
  TDBGrid2.ReBind
End Sub

Private Sub ProsesSimpan()
Dim n As Integer
  
  InitPB vaDP.UpperBound(1) + 1
  For n = 0 To vaDP.UpperBound(1)
    RunPB
    'objData.Edit GetDSN, "Deposito", "Rekening='" & vaDP(n, 2) & "'", Array("JthTmp", "LastPerpanjangan"), Array(vaDP(n, 7), dTgl.Value)
    objData.Edit GetDSN, "Deposito", "Rekening='" & vaDP(n, 2) & "'", Array("JthTmp", "LastPerpanjangan", "jumlahperpanjangan"), Array(vaDP(n, 7), vaDP(n, 1), Val(vaDP(n, 8)) + 1)
  Next
  EndPB
End Sub
