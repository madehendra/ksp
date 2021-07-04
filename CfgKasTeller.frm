VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form CfgKasTeller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETUP KAS TELLER"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10290
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3270
      Left            =   0
      Top             =   1305
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   5768
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
         Height          =   3150
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   5556
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "USER NAME"
         Columns(0).DataField=   "UserName"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "FULL NAME"
         Columns(1).DataField=   "FullName"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "REKENING TELLER"
         Columns(2).DataField=   "KasTeller"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NAMA PERKIARAAN"
         Columns(3).DataField=   "Keterangan"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "MAX PLAFOND"
         Columns(4).DataField=   "Plafond"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2672"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=4101"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4022"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2937"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2858"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=4657"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4577"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2858"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2778"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HC0C0C0&"
         _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000012&,.appearance=0,.ellipsis=0,.borderColor=&HFF8000&"
         _StyleDefs(11)  =   ":id=4,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
         _StyleDefs(14)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(15)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(17)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(18)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(19)  =   ":id=3,.fontname=MS Sans Serif"
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
         _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1305
      Left            =   0
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
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
      BackColor       =   -2147483633
      Begin BiSANumberBoxProject.BiSANumberBox nPlafond 
         Height          =   345
         Left            =   315
         TabIndex        =   6
         Top             =   810
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   609
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "MAX PLAFOND"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekening 
         Height          =   330
         Left            =   4125
         TabIndex        =   3
         Top             =   465
         Width           =   5145
         _ExtentX        =   9075
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Left            =   330
         TabIndex        =   2
         Top             =   465
         Width           =   3795
         _ExtentX        =   6694
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
         Caption         =   "KAS TELLER"
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
      Begin BiSATextBoxProject.BiSATextBox cUser 
         Height          =   330
         Left            =   330
         TabIndex        =   1
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
         FontName        =   "Verdana"
         Caption         =   "USER NAME"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   4560
      Width           =   10275
      _ExtentX        =   18124
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
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   8010
         TabIndex        =   4
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
         Picture         =   "CfgKasTeller.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9090
         TabIndex        =   5
         Top             =   120
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
         Picture         =   "CfgKasTeller.frx":0416
      End
   End
End
Attribute VB_Name = "CfgKasTeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim db As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  objData.Edit GetDSN, "UserName", "Username = '" & cUser.Text & "'", Array("KasTeller", "Plafond"), Array(cRekening.Text, nPlafond.Value)
  GetSQL
End Sub

Private Sub cRekening_ButtonClick()
  Set db = objData.Browse(GetDSN, "Rekening", "Kode,Keterangan", "Kode", sisContent, cRekening.Text, , "Kode")
  cRekening.Text = cRekening.Browse(db)
  If Not db.eof Then
    cNamaRekening.Text = GetNull(db!Keterangan, "")
  End If
End Sub

Private Sub cRekening_Validate(Cancel As Boolean)
  If cRekening.LastKey = 13 Or cRekening.LastKey = 40 Then
    cRekening_ButtonClick
  End If
End Sub

Private Sub cUser_Change()
  cmdSimpan.Enabled = Trim(cUser.Text) <> ""
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
  If lClick Then
    If Not dbData.eof Then
      cUser.Text = GetNull(dbData!UserName, "")
      cRekening.Text = GetNull(dbData!KasTeller, "")
      cNamaRekening.Text = GetNull(dbData!Keterangan, "")
      nPlafond.Value = GetNull(dbData!plafond)
    End If
  End If
  lClick = False
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  GetSQL
  
  TabIndex cUser, n
  TabIndex cRekening, n
  TabIndex cNamaRekening, n
  TabIndex nPlafond, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSQL()
  Set dbData = objData.Browse(GetDSN, "Username u", "u.UserName,u.FullName,u.KasTeller,r.Keterangan,u.Plafond", , , , , "u.Username", _
               Array("Left Join Rekening r on u.KasTeller = r.Kode"))
  Set DataGrid1.DataSource = dbData
End Sub
