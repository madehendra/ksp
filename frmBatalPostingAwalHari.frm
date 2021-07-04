VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form frmBatalPostingAwalHari 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMBATALAN POSTING AWAL HARI"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   5025
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   4110
      TabIndex        =   6
      Top             =   390
      Width           =   465
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   375
      Left            =   2490
      TabIndex        =   3
      Top             =   5160
      Width           =   1005
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&BATAL"
      Height          =   375
      Left            =   1485
      TabIndex        =   2
      Top             =   5160
      Width           =   1005
   End
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   360
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   375
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   635
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
      Caption         =   "Tgl"
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
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   345
      Index           =   1
      Left            =   2265
      TabIndex        =   4
      Top             =   375
      Width           =   1740
      _ExtentX        =   3069
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3270
      Left            =   210
      TabIndex        =   5
      Top             =   810
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5768
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Tgl"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Faktur"
      Columns(2).DataField=   ""
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=661"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2408"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2328"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4445"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4366"
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
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=112,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HC0C0C0&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000007&,.bold=0,.fontsize=825,.italic=0,.underline=0"
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
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
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
      _StyleDefs(67)  =   ":id=41,.parent=34,.alignment=3"
      _StyleDefs(68)  =   "Named:id=42:FilterBar"
      _StyleDefs(69)  =   ":id=42,.parent=33,.alignment=3"
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   795
      Left            =   255
      TabIndex        =   7
      Top             =   4260
      Width           =   4560
   End
   Begin VB.Label Label1 
      Caption         =   "TANGGAL TRANSAKSI YG MAU DIBATALKAN"
      Height          =   240
      Left            =   615
      TabIndex        =   1
      Top             =   90
      Width           =   3540
   End
End
Attribute VB_Name = "frmBatalPostingAwalHari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB


Private Sub cmdBatal_Click()
Dim vaField, vaValue
Dim n As Integer
  
  If MsgBox("Yakin akan melakukan pembatalan?", vbExclamation + vbYesNo) = vbYes Then
    TDBGrid1.MoveFirst
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 0) = -1 Then
        objData.Delete GetDSN, "MutasiBungaDeposito", "Faktur", sisAssign, vaArray(n, 2)
        objData.Delete GetDSN, "BungaDeposito", "Faktur", sisAssign, vaArray(n, 2)
        objData.Delete GetDSN, "BukuBesar", "Faktur", sisAssign, vaArray(n, 2)
      End If
    Next
    GetData
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single
  
  Label2.Caption = "Pembatalan proses awal hari akan melakukan pembatalan dari posting yang telah dilakukan pada tgl yang bersangkutan" & vbCrLf & _
                   "Bunga hasil posting yg sudah terlanjur dicairkan tidak bisa dibatalkan"
  
  GetData
  CenterForm Me
  Me.Top = 0
  dTgl(0).Value = Date
  dTgl(1).Value = Date
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cmdOK, n
  TabIndex cmdBatal, n
  TabIndex cmdExit, n
End Sub

Private Sub GetData()
Dim n As Integer
  
  vaArray.ReDim 0, -1, 0, 2
  Set dbData = objData.Browse(GetDSN, "MutasiBungaDeposito", "Distinct(Faktur),Tgl", "Tgl", sisGTEqual, Format(dTgl(0).Value, "yyyy-MM-dd"), " and Tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'", "Tgl")
  If Not dbData.eof Then
    dbData.MoveFirst
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = 0
      vaArray(n, 1) = GetNull(dbData!Tgl)
      vaArray(n, 2) = GetNull(dbData!Faktur)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub
