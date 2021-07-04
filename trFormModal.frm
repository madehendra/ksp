VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form trFormModal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Modal"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   11655
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3720
      Left            =   0
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   6562
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
         Height          =   3600
         Left            =   75
         TabIndex        =   0
         Top             =   60
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   6350
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Kode Reg"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=476"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=397"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4392"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4313"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=13203"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=13123"
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
         Appearance      =   2
         DataMode        =   4
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
      Top             =   3705
      Width           =   11670
      _ExtentX        =   20585
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
      Begin BiSAButtonProject.BiSAButton cmdDelete 
         Height          =   435
         Left            =   9435
         TabIndex        =   1
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         Caption         =   "    &Update"
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
         Picture         =   "trFormModal.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10515
         TabIndex        =   2
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
         Picture         =   "trFormModal.frx":0416
      End
   End
End
Attribute VB_Name = "trFormModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cmdDelete_Click()
Dim n As Integer

  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    If vaArray(n, 0) <> 0 Then
      objData.Edit GetDSN, "registernasabah", " kode = '" & vaArray(n, 1) & "'", Array("modalpenyertaan"), Array(1)
    End If
  Next n
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  TabIndex cmdDelete, n
  TabIndex cmdKeluar, n
  
  vaArray.ReDim 0, -1, 0, 2
  Set dbData = objData.Browse(GetDSN, "registernasabah", , , , , , "kode")
  If Not dbData.eof Then
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      If GetNull(dbData!modalpenyertaan) <> 0 Then
        vaArray(n, 0) = -1
      Else
        vaArray(n, 0) = 0
      End If
      vaArray(n, 1) = GetNull(dbData!Kode)
      vaArray(n, 2) = GetNull(dbData!nama)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub
