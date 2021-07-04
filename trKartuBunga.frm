VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form trKartuBunga 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KARTU BUNGA DEPOSITO"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   11790
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   1508
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
         Height          =   360
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   635
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
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4725
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   8334
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO"
      Columns(0).DataField=   ""
      Columns(0).NumberFormat=   "FormatText Event"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "REK"
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
      Columns(3).NumberFormat=   "###,###,###,##0.00"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TGL VALUTA"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "LAMA"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "###,###,###,###,##0.00"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "BUNGA(PA)"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "###,###,###,##0.00"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "TGL BUNGA"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "BUNGA"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "###,###,###,###,##0.00"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2461"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2381"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1).Merge=1"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=3836"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3757"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(2).Merge=1"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2593"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2514"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(3).Merge=1"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2143"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2064"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(4).Merge=1"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=1164"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=1085"
      Splits(0)._ColumnProps(33)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(5).Merge=1"
      Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(41)=   "Column(6).Merge=1"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=2275"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2196"
      Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(7).Merge=1"
      Splits(0)._ColumnProps(48)=   "Column(8).Width=2064"
      Splits(0)._ColumnProps(49)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(8)._WidthInPix=1984"
      Splits(0)._ColumnProps(51)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(52)=   "Column(8).Order=9"
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
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HEBDACB&"
      _StyleDefs(15)  =   ":id=3,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
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
      _StyleDefs(26)  =   "Splits(0).Style:id=95,.parent=1"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=104,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=96,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=97,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=98,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=100,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=99,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=101,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=102,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=103,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=105,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=106,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=110,.parent=95,.alignment=1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=107,.parent=96"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=108,.parent=97"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=109,.parent=99"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=114,.parent=95,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=111,.parent=96"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=112,.parent=97"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=113,.parent=99"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=95,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=96"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=97"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=99"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=122,.parent=95,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=119,.parent=96"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=120,.parent=97"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=121,.parent=99"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=126,.parent=95,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=123,.parent=96"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=124,.parent=97"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=125,.parent=99"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=46,.parent=95"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=96"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=97"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=99"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=32,.parent=95,.alignment=1"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=96"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=97"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=99"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=130,.parent=95,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=127,.parent=96"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=128,.parent=97"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=129,.parent=99"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=150,.parent=95,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=147,.parent=96"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=148,.parent=97"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=149,.parent=99"
      _StyleDefs(74)  =   "Named:id=33:Normal"
      _StyleDefs(75)  =   ":id=33,.parent=0"
      _StyleDefs(76)  =   "Named:id=34:Heading"
      _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   ":id=34,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=35:Footing"
      _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=36:Selected"
      _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=37:Caption"
      _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(85)  =   "Named:id=38:HighlightRow"
      _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=39:EvenRow"
      _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(89)  =   "Named:id=40:OddRow"
      _StyleDefs(90)  =   ":id=40,.parent=33"
      _StyleDefs(91)  =   "Named:id=41:RecordSelector"
      _StyleDefs(92)  =   ":id=41,.parent=34"
      _StyleDefs(93)  =   "Named:id=42:FilterBar"
      _StyleDefs(94)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "trKartuBunga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB

Private Sub initvalue()
  vaArray.ReDim 0, -1, 0, 8
  Set TDBGrid1.Array = vaArray
  TDBGrid1.Refresh
End Sub

Private Sub GetData()
Dim n As Integer
initvalue
Set dbData = objData.Browse(GetDSN, "MutasiDeposito m", "m.Rekening,r.Nama,d.NominalDeposito,d.Tgl as TglValuta,d.Lama,d.SukuBunga,m.Tgl as TglMutasi,m.Jumlah as Bunga", "m.KodeMutasi", sisAssign, "3", , "m.Rekening,m.Tgl", Array("Left Join deposito d on d.Rekening = m.Rekening", "Left Join RegisterNasabah r on r.Kode = d.Kode"))  'mutasi bunga
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!Rekening)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!nominaldeposito)
      vaArray(n, 4) = Format(GetNull(dbData!TglValuta), "dd-MM-yyyy")
      vaArray(n, 5) = GetNull(dbData!Lama)
      vaArray(n, 6) = GetNull(dbData!SukuBunga)
      vaArray(n, 7) = Format(GetNull(dbData!TglMutasi), "dd-MM-yyyy")
      vaArray(n, 8) = GetNull(dbData!bunga)
      dbData.MoveNext
    Loop
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
    FrmPB.EndPB
  End If
End Sub

Private Sub cmdOK_Click()
  GetData
End Sub
