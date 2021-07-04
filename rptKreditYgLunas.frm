VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form rptKreditYgLunas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN PINJAMAN YG SUDAH LUNAS"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11790
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   840
      Left            =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   3060
         TabIndex        =   3
         Top             =   360
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         Caption         =   "OK"
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
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   360
         Width           =   2940
         _ExtentX        =   5186
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
         Caption         =   "Kriteria Nama Debitur :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   2010
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   5595
      Left            =   0
      Top             =   825
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9869
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
      Begin TrueOleDBGrid70.TDBGrid DataGrid 
         Height          =   5475
         Left            =   60
         TabIndex        =   0
         Top             =   75
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   9657
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "REKENING"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KODE"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "###,###,###,##0.00"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TGL VALUTA"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "PLAFOND"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "LAMA"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "SUKU BUNGA"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "AO"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3572"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3493"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3016"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2937"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=5530"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5450"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2381"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2302"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3519"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3440"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2011"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1931"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=1958"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1879"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=5689"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=5609"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=512"
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
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=0"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=0"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(71)  =   "Named:id=33:Normal"
         _StyleDefs(72)  =   ":id=33,.parent=0"
         _StyleDefs(73)  =   "Named:id=34:Heading"
         _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   ":id=34,.wraptext=-1"
         _StyleDefs(76)  =   "Named:id=35:Footing"
         _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   "Named:id=36:Selected"
         _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=37:Caption"
         _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(82)  =   "Named:id=38:HighlightRow"
         _StyleDefs(83)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(84)  =   "Named:id=39:EvenRow"
         _StyleDefs(85)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(86)  =   "Named:id=40:OddRow"
         _StyleDefs(87)  =   ":id=40,.parent=33"
         _StyleDefs(88)  =   "Named:id=41:RecordSelector"
         _StyleDefs(89)  =   ":id=41,.parent=34"
         _StyleDefs(90)  =   "Named:id=42:FilterBar"
         _StyleDefs(91)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "rptKreditYgLunas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB
Dim lClick As Boolean

Private Sub initvalue()
  cNama.Default
  vaArray.ReDim 0, -1, 0, 7
  Set DataGrid.Array = vaArray
  DataGrid.ReBind
  DataGrid.Refresh
End Sub

Private Sub cmdOK_Click()
  If Trim(cNama.Text) = "" Then
    If MsgBox("Anda tidak memasukkan kriteria, apakah data akan ditampilkan seluruhnya?", vbInformation + vbYesNo) = vbYes Then
      GetData False
    Else
      initvalue
    End If
  Else
    GetData True, cNama.Text
  End If
End Sub

Private Sub GetData(ByVal lFilter As Boolean, Optional ByVal Kriteria As String = "")
Dim cWhere As String
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 7
  
  If lFilter = True Then
    cWhere = " and r.Nama like '%" & Kriteria & "%'"
  Else
    cWhere = ""
  End If
  
  
  Set dbData = objData.Browse(GetDSN, "Debitur d", _
  "d.Rekening,d.Kode,r.Nama,d.Tgl,d.Plafond,d.Lama,d.TotalBunga,a.Nama as NamaAO", "d.Status", sisAssign, "1", cWhere, "Rekening", _
  Array("Left Join RegisterNasabah r on r.Kode = d.Kode", "Left Join AO a on a.Kode = d.AO"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Rekening, "")
      vaArray(n, 1) = GetNull(dbData!Kode, "")
      vaArray(n, 2) = GetNull(dbData!nama, "")
      vaArray(n, 3) = GetNull(dbData!Tgl)
      vaArray(n, 4) = GetNull(dbData!plafond)
      vaArray(n, 5) = GetNull(dbData!Lama)
      vaArray(n, 6) = GetNull(dbData!totalBunga)
      vaArray(n, 7) = GetNull(dbData!namaao)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  Else
    MsgBox "Laporan kredit yang sudah lunas tidak bisa ditampilkan karena data tidak ada..", vbInformation, Me.Caption
  End If
  
  Set DataGrid.Array = vaArray
  DataGrid.ReBind
  DataGrid.Refresh
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
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 4, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 5
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 5, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 6
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 6, XORDER_ASCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 7
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 7, XORDER_ASCEND, XTYPE_STRING
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
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 4, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 5
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 5, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 6
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 6, XORDER_DESCEND, XTYPE_DOUBLE
        lClick = Not lClick
      Case 7
        vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 7, XORDER_DESCEND, XTYPE_STRING
        lClick = Not lClick
    End Select
  End If
  DataGrid.ReBind
End Sub

Private Sub Form_Load()
Dim n As Single
  lClick = True
  'INISIALISASI TAB INDEX
  TabIndex cNama, n
  TabIndex cmdOK, n
  'INITVALUE
  initvalue
End Sub
