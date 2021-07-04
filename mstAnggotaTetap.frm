VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form mstAnggotaTetap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Anggota Tetap"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7170
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   600
      Left            =   0
      Top             =   4500
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   1058
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
         Left            =   4905
         TabIndex        =   7
         Top             =   90
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
         Picture         =   "mstAnggotaTetap.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5985
         TabIndex        =   8
         Top             =   90
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
         Picture         =   "mstAnggotaTetap.frx":0416
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3540
      Left            =   0
      Top             =   975
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   6244
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
      Begin BiSATextBoxProject.BiSABrowse cKodeAnggotaFiktif 
         Height          =   330
         Left            =   4575
         TabIndex        =   0
         Top             =   60
         Width           =   1995
         _ExtentX        =   3519
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
         Appearance      =   0
         Button          =   -1  'True
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
      Begin BiSATextBoxProject.BiSABrowse cNamaAnggotaFiktif 
         Height          =   330
         Left            =   615
         TabIndex        =   1
         Top             =   60
         Width           =   3945
         _ExtentX        =   6959
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
         Appearance      =   0
         Button          =   -1  'True
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
      Begin BiSANumberBoxProject.BiSANumberBox nNomor 
         Height          =   330
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         DecimalPoint    =   ""
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " "
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
         Height          =   3075
         Left            =   60
         TabIndex        =   3
         Top             =   405
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   5424
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO."
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA ANGGOTA"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KODE"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=7064"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6985"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1482"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1402"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
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
         Appearance      =   0
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   0
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000007&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   6600
         TabIndex        =   4
         Top             =   60
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
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
         Picture         =   "mstAnggotaTetap.frx":04BC
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1005
      Left            =   0
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   1773
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
      Begin BiSATextBoxProject.BiSATextBox cKodeAnggota 
         Height          =   345
         Left            =   1275
         TabIndex        =   6
         Top             =   450
         Width           =   2820
         _ExtentX        =   4974
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
         FontName        =   "Verdana"
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
      Begin BiSATextBoxProject.BiSABrowse cNamaAnggota 
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   105
         Width           =   3945
         _ExtentX        =   6959
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
         Appearance      =   0
         Button          =   -1  'True
         Caption         =   "Anggota"
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
Attribute VB_Name = "mstAnggotaTetap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cKodeAnggotaFiktif_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "registernasabah", "kode,nama,alamat", , , , , "kode")
  If Not dbData.eof Then
    cKodeAnggotaFiktif.Text = cKodeAnggotaFiktif.Browse(dbData)
    cNamaAnggotaFiktif.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
Dim n As Integer
    
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.Value Then
      vaArray.ReDim 0, nNomor.Value - 1, 0, 2
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.Value = 1
      vaArray.ReDim 0, nNomor.Value - 1, 0, 2
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.Value - 1
    End If
        
    vaArray(n, 0) = nNomor.Value
    vaArray(n, 1) = cNamaAnggotaFiktif.Text
    vaArray(n, 2) = cKodeAnggotaFiktif.Text
    
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    
    nNomor.Value = vaArray.UpperBound(1) + 2
    cNamaAnggotaFiktif.SetFocus
End Sub

Private Sub cmdSimpan_Click()
Dim n As Integer

  If MsgBox("Data akan disimpan?", vbYesNo + vbInformation) = vbYes Then
    objData.Delete GetDSN, "anggotafiktif", "pemegangsaham", sisAssign, cKodeAnggota.Text
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      objData.Add GetDSN, "anggotafiktif", Array("pemegangsaham", "anggota"), Array(cKodeAnggota.Text, vaArray(n, 2))
    Next n
    initvalue
  End If
End Sub

Private Sub cNamaAnggota_ButtonClick()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 2
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
      
  Set dbData = objData.Browse(GetDSN, "registernasabah", "nama,kode,alamat", , , , , "nama")
  If Not dbData.eof Then
    cNamaAnggota.Text = cNamaAnggota.Browse(dbData)
    cKodeAnggota.Text = GetNull(dbData!Kode)
    
    Set dbData = objData.Browse(GetDSN, "anggotafiktif f", "f.anggota,r.nama", "pemegangsaham", sisAssign, cKodeAnggota.Text, , , Array("LEFT JOIN registernasabah r on r.kode = f.anggota"))
    If Not dbData.eof Then
      Do While Not dbData.eof
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(dbData!nama)
        vaArray(n, 2) = GetNull(dbData!anggota)
        dbData.MoveNext
      Loop
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      TDBGrid1.Refresh
    End If
  End If
End Sub

Private Sub cNamaAnggotaFiktif_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "registernasabah", "nama,kode,alamat", , , , , "nama")
  If Not dbData.eof Then
    cNamaAnggotaFiktif.Text = cNamaAnggotaFiktif.Browse(dbData)
    cKodeAnggotaFiktif.Text = GetNull(dbData!Kode)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  
  vaArray.ReDim 0, -1, 0, 2
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  
  TabIndex cNamaAnggota, n
  TabIndex cKodeAnggota, n
  TabIndex nNomor, n
  TabIndex cNamaAnggotaFiktif, n
  TabIndex cKodeAnggotaFiktif, n
  TabIndex cmdOK, n
  initvalue
End Sub

Private Sub initvalue()
  cNamaAnggota.Default
  cKodeAnggota.Default
  nNomor.Value = 1
  cNamaAnggotaFiktif.Default
  cKodeAnggotaFiktif.Default
  
  vaArray.ReDim 0, -1, 0, 2
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind

End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.Value - 1
    If n <= vaArray.UpperBound(1) Then
      cNamaAnggotaFiktif.Text = vaArray(n, 1)
      cKodeAnggotaFiktif.Text = vaArray(n, 2)
    End If
  End If
End Sub

Function GetValidNomorUrut(ByVal oNomor As Object, ByVal vaNomor As XArrayDB) As Boolean
GetValidNomorUrut = True
'Desc.
'===============================================================================================
': Function untuk melakukan validasi nomor urut ketika data dimasukkan ke dalam grid
'digunakan untuk menghindari kesalahan seperti "subsbrice out of range" dan error yang lain
'===============================================================================================
'by made hendra, made.hendra@gmail.com

  If oNomor.Value - 2 > vaNomor.UpperBound(1) Or oNomor.Value <= 0 Then
    oNomor.Value = vaNomor.UpperBound(1) + 2
  End If
  If vaNomor.UpperBound(1) < 0 Then
    GetValidNomorUrut = False
  End If
End Function

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNomor.Value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub
