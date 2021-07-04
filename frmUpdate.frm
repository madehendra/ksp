VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Deposito"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8865
   Begin BiSAButtonProject.BiSAButton BiSAButton3 
      Height          =   480
      Left            =   315
      TabIndex        =   4
      Top             =   4860
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   847
      Caption         =   "Label1"
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
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   495
      Left            =   1395
      TabIndex        =   1
      Top             =   750
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   873
      Caption         =   "Update"
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
   Begin BiSADateProject.BiSADate dDate 
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   661
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
      Caption         =   "Tgl Posting Bunga"
      CaptionWidth    =   1700
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Height          =   3330
      Left            =   30
      TabIndex        =   2
      Top             =   1395
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   5874
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Rekening"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Faktur"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Bunga"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tabungan"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=556"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=476"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3413"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3334"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3969"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3889"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2937"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2858"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=4763"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=4683"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(57)  =   "Named:id=33:Normal"
      _StyleDefs(58)  =   ":id=33,.parent=0"
      _StyleDefs(59)  =   "Named:id=34:Heading"
      _StyleDefs(60)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   ":id=34,.wraptext=-1"
      _StyleDefs(62)  =   "Named:id=35:Footing"
      _StyleDefs(63)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   "Named:id=36:Selected"
      _StyleDefs(65)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(66)  =   "Named:id=37:Caption"
      _StyleDefs(67)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(68)  =   "Named:id=38:HighlightRow"
      _StyleDefs(69)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(70)  =   "Named:id=39:EvenRow"
      _StyleDefs(71)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(72)  =   "Named:id=40:OddRow"
      _StyleDefs(73)  =   ":id=40,.parent=33"
      _StyleDefs(74)  =   "Named:id=41:RecordSelector"
      _StyleDefs(75)  =   ":id=41,.parent=34"
      _StyleDefs(76)  =   "Named:id=42:FilterBar"
      _StyleDefs(77)  =   ":id=42,.parent=33"
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   495
      Left            =   6900
      TabIndex        =   3
      Top             =   4860
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   873
      Caption         =   "Simpan"
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
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
'cari data posting di tabel bungadeposito
'ambil data posting di tabel mutasitabungan
'bandingkan
'masukkan data yg tidak masuk di table mutasitabungan dengan menggunakan referensi data di table mutasideposito
'set dbdata = objdata.Browse(getdsn,""

Dim cKodeTransaksi As String
Dim cDK As String
Dim cKeterangan As String
Dim n As Single

Dim cSQL As String
  cSQL = "SELECT m.rekening,m.Faktur,m.bunga,m.tgl from bungadeposito m " & _
         "left JOIN mutasitabungan t on t.FAKTUR = m.Faktur and (RIGHT(t.Rekening,9)=RIGHT(m.REKENING,9)) " & _
         "Where M.Tgl = '" & Format(dDate.Value, "yyyy-MM-dd") & "' And t.Rekening Is Null"
  
  cKodeTransaksi = 13
  cDK = "K"
  cKeterangan = "Pencairan Deposito Ke Tabungan"
  
  vaArray.ReDim 0, -1, 0, 4
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = -1
      vaArray(n, 1) = GetNull(dbData!Rekening)
      vaArray(n, 2) = GetNull(dbData!Faktur)
      vaArray(n, 3) = GetNull(dbData!bunga)
      vaArray(n, 4) = GetRekSimpananDeposito(GetNull(dbData!Rekening))
      dbData.MoveNext
    Loop
    Set TDBGrid2.Array = vaArray
    TDBGrid2.ReBind
    TDBGrid2.Refresh
  End If
End Sub

Private Function GetRekSimpananDeposito(cRekDeposito As String) As String
Dim db As New ADODB.Recordset

  GetRekSimpananDeposito = ""
  Set db = objData.Browse(GetDSN, "deposito", , "rekening", sisAssign, cRekDeposito)
  If Not db.eof Then
    GetRekSimpananDeposito = GetNull(db!rekeningsimpanan)
  End If
End Function

Private Sub BiSAButton2_Click()
Dim n As Single

 For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
  If vaArray(n, 0) = -1 Then
    UpdMutasiTabungan objData, aCfg(msKodeTransaksiPB), vaArray(n, 2), Format(dDate.Value, "yyyy-MM-dd"), vaArray(n, 4), vaArray(n, 3), False, "Pencairan Bunga Deposito ke Tabungan", False
  End If
 Next n
 MsgBox "Data telah berhasil di update"
End Sub

Private Sub BiSAButton3_Click()
Dim n As Single

  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    MsgBox vaArray(n, 1) & " " & vaArray(n, 0)
  Next n
End Sub

Private Sub TDBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid2.Update
  TDBGrid2.ReBind
End Sub

