VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form frmSHU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM SHU"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   12030
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   6060
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2145
      Width           =   12030
      _cx             =   21220
      _cy             =   10689
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin TrueOleDBGrid70.TDBGrid tdbgrid1 
         Height          =   6000
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   10583
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NO ANGGOTA"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TOTAL BUNGA PINJAMAN"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "###,###,###,###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "KONTRIBUSI"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "SHU PINJAMAN"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "% PINJAMAN"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "TOTAL SIM HARIAN MENGENDAP"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###,###,##0.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "KONTRIBUSI"
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "###,###,###,###,##0.00"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "SHU SIM HARIAN"
         Columns(9).DataField=   ""
         Columns(9).NumberFormat=   "###,###,###,###,##0.00"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "% SIM HARIAN"
         Columns(10).DataField=   ""
         Columns(10).NumberFormat=   "###,###,###,###,##0.00"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "TOTAL BUNGA SIM BERJANGKA"
         Columns(11).DataField=   ""
         Columns(11).NumberFormat=   "###,###,###,###,##0.00"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "KONTRIBUSI"
         Columns(12).DataField=   ""
         Columns(12).NumberFormat=   "###,###,###,###,##0.00"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "SHU SIM BERJANGKA"
         Columns(13).DataField=   ""
         Columns(13).NumberFormat=   "###,###,###,###,##0.00"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "% SIM BERJANGKA"
         Columns(14).DataField=   ""
         Columns(14).NumberFormat=   "###,###,###,###,##0.00"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "SHU MODAL"
         Columns(15).DataField=   ""
         Columns(15).NumberFormat=   "###,###,###,###,##0.00"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "SHU"
         Columns(16).DataField=   ""
         Columns(16).NumberFormat=   "###,###,###,###,##0.00"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   17
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).Size  =   354
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=17"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1164"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1058"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2937"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2831"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5345"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5239"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2831"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2725"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=2461"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2355"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2090"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1984"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=2937"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2831"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(49)=   "Column(8).Width=2593"
         Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=2487"
         Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(55)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=2619"
         Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(61)=   "Column(10).Width=2355"
         Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=2249"
         Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(10)._ColStyle=516"
         Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(67)=   "Column(11).Width=4419"
         Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=4313"
         Splits(0)._ColumnProps(70)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(11)._ColStyle=516"
         Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(73)=   "Column(12).Width=2646"
         Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=2540"
         Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(77)=   "Column(12)._ColStyle=516"
         Splits(0)._ColumnProps(78)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(79)=   "Column(13).Width=3043"
         Splits(0)._ColumnProps(80)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(13)._WidthInPix=2937"
         Splits(0)._ColumnProps(82)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(83)=   "Column(13)._ColStyle=516"
         Splits(0)._ColumnProps(84)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(85)=   "Column(14).Width=2910"
         Splits(0)._ColumnProps(86)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(14)._WidthInPix=2805"
         Splits(0)._ColumnProps(88)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(89)=   "Column(14)._ColStyle=516"
         Splits(0)._ColumnProps(90)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(91)=   "Column(15).Width=3281"
         Splits(0)._ColumnProps(92)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(93)=   "Column(15)._WidthInPix=3175"
         Splits(0)._ColumnProps(94)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(95)=   "Column(15)._ColStyle=516"
         Splits(0)._ColumnProps(96)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(97)=   "Column(16).Width=4180"
         Splits(0)._ColumnProps(98)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(99)=   "Column(16)._WidthInPix=4075"
         Splits(0)._ColumnProps(100)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(101)=   "Column(16)._ColStyle=514"
         Splits(0)._ColumnProps(102)=   "Column(16).FetchStyle=1"
         Splits(0)._ColumnProps(103)=   "Column(16).Order=17"
         Splits(1)._UserFlags=   0
         Splits(1).MarqueeStyle=   3
         Splits(1).Size  =   438
         Splits(1).Size.vt=   2
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   873
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1).AlternatingRowStyle=   -1  'True
         Splits(1).DividerColor=   15790320
         Splits(1).FilterBar=   -1  'True
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=17"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=1164"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1058"
         Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(8)=   "Column(1).Width=2937"
         Splits(1)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(10)=   "Column(1)._WidthInPix=2831"
         Splits(1)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(1)._ColumnProps(12)=   "Column(1)._ColStyle=516"
         Splits(1)._ColumnProps(13)=   "Column(1).Visible=0"
         Splits(1)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(15)=   "Column(2).Width=5345"
         Splits(1)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(17)=   "Column(2)._WidthInPix=5239"
         Splits(1)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(1)._ColumnProps(19)=   "Column(2)._ColStyle=513"
         Splits(1)._ColumnProps(20)=   "Column(2).Visible=0"
         Splits(1)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(22)=   "Column(3).Width=3281"
         Splits(1)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(24)=   "Column(3)._WidthInPix=3175"
         Splits(1)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(1)._ColumnProps(26)=   "Column(3)._ColStyle=514"
         Splits(1)._ColumnProps(27)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(28)=   "Column(4).Width=2831"
         Splits(1)._ColumnProps(29)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(30)=   "Column(4)._WidthInPix=2725"
         Splits(1)._ColumnProps(31)=   "Column(4)._EditAlways=0"
         Splits(1)._ColumnProps(32)=   "Column(4)._ColStyle=514"
         Splits(1)._ColumnProps(33)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(34)=   "Column(5).Width=2461"
         Splits(1)._ColumnProps(35)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(36)=   "Column(5)._WidthInPix=2355"
         Splits(1)._ColumnProps(37)=   "Column(5)._EditAlways=0"
         Splits(1)._ColumnProps(38)=   "Column(5)._ColStyle=514"
         Splits(1)._ColumnProps(39)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(40)=   "Column(6).Width=2090"
         Splits(1)._ColumnProps(41)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(42)=   "Column(6)._WidthInPix=1984"
         Splits(1)._ColumnProps(43)=   "Column(6)._EditAlways=0"
         Splits(1)._ColumnProps(44)=   "Column(6)._ColStyle=514"
         Splits(1)._ColumnProps(45)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(46)=   "Column(7).Width=2937"
         Splits(1)._ColumnProps(47)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(48)=   "Column(7)._WidthInPix=2831"
         Splits(1)._ColumnProps(49)=   "Column(7)._EditAlways=0"
         Splits(1)._ColumnProps(50)=   "Column(7)._ColStyle=514"
         Splits(1)._ColumnProps(51)=   "Column(7).Order=8"
         Splits(1)._ColumnProps(52)=   "Column(8).Width=2593"
         Splits(1)._ColumnProps(53)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(54)=   "Column(8)._WidthInPix=2487"
         Splits(1)._ColumnProps(55)=   "Column(8)._EditAlways=0"
         Splits(1)._ColumnProps(56)=   "Column(8)._ColStyle=514"
         Splits(1)._ColumnProps(57)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(58)=   "Column(9).Width=2725"
         Splits(1)._ColumnProps(59)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(60)=   "Column(9)._WidthInPix=2619"
         Splits(1)._ColumnProps(61)=   "Column(9)._EditAlways=0"
         Splits(1)._ColumnProps(62)=   "Column(9)._ColStyle=514"
         Splits(1)._ColumnProps(63)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(64)=   "Column(10).Width=2355"
         Splits(1)._ColumnProps(65)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(66)=   "Column(10)._WidthInPix=2249"
         Splits(1)._ColumnProps(67)=   "Column(10)._EditAlways=0"
         Splits(1)._ColumnProps(68)=   "Column(10)._ColStyle=514"
         Splits(1)._ColumnProps(69)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(70)=   "Column(11).Width=4419"
         Splits(1)._ColumnProps(71)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(72)=   "Column(11)._WidthInPix=4313"
         Splits(1)._ColumnProps(73)=   "Column(11)._EditAlways=0"
         Splits(1)._ColumnProps(74)=   "Column(11)._ColStyle=514"
         Splits(1)._ColumnProps(75)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(76)=   "Column(12).Width=2646"
         Splits(1)._ColumnProps(77)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(78)=   "Column(12)._WidthInPix=2540"
         Splits(1)._ColumnProps(79)=   "Column(12)._EditAlways=0"
         Splits(1)._ColumnProps(80)=   "Column(12)._ColStyle=514"
         Splits(1)._ColumnProps(81)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(82)=   "Column(13).Width=3043"
         Splits(1)._ColumnProps(83)=   "Column(13).DividerColor=0"
         Splits(1)._ColumnProps(84)=   "Column(13)._WidthInPix=2937"
         Splits(1)._ColumnProps(85)=   "Column(13)._EditAlways=0"
         Splits(1)._ColumnProps(86)=   "Column(13)._ColStyle=514"
         Splits(1)._ColumnProps(87)=   "Column(13).Order=14"
         Splits(1)._ColumnProps(88)=   "Column(14).Width=2910"
         Splits(1)._ColumnProps(89)=   "Column(14).DividerColor=0"
         Splits(1)._ColumnProps(90)=   "Column(14)._WidthInPix=2805"
         Splits(1)._ColumnProps(91)=   "Column(14)._EditAlways=0"
         Splits(1)._ColumnProps(92)=   "Column(14)._ColStyle=514"
         Splits(1)._ColumnProps(93)=   "Column(14).Order=15"
         Splits(1)._ColumnProps(94)=   "Column(15).Width=3281"
         Splits(1)._ColumnProps(95)=   "Column(15).DividerColor=0"
         Splits(1)._ColumnProps(96)=   "Column(15)._WidthInPix=3175"
         Splits(1)._ColumnProps(97)=   "Column(15)._EditAlways=0"
         Splits(1)._ColumnProps(98)=   "Column(15)._ColStyle=514"
         Splits(1)._ColumnProps(99)=   "Column(15).Order=16"
         Splits(1)._ColumnProps(100)=   "Column(16).Width=4180"
         Splits(1)._ColumnProps(101)=   "Column(16).DividerColor=0"
         Splits(1)._ColumnProps(102)=   "Column(16)._WidthInPix=4075"
         Splits(1)._ColumnProps(103)=   "Column(16)._EditAlways=0"
         Splits(1)._ColumnProps(104)=   "Column(16)._ColStyle=514"
         Splits(1)._ColumnProps(105)=   "Column(16).FetchStyle=1"
         Splits(1)._ColumnProps(106)=   "Column(16).Order=17"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   0
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   15790320
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=825"
         _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=90,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=87,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=88,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=89,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=94,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=91,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=92,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=93,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=50,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=98,.parent=13"
         _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=95,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=96,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=97,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=54,.parent=13"
         _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=51,.parent=14"
         _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=52,.parent=15"
         _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=53,.parent=17"
         _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=178,.parent=13"
         _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=175,.parent=14"
         _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=176,.parent=15"
         _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=177,.parent=17"
         _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=63,.parent=14"
         _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=64,.parent=15"
         _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=65,.parent=17"
         _StyleDefs(102) =   "Splits(1).Style:id=99,.parent=1"
         _StyleDefs(103) =   "Splits(1).CaptionStyle:id=108,.parent=4"
         _StyleDefs(104) =   "Splits(1).HeadingStyle:id=100,.parent=2"
         _StyleDefs(105) =   "Splits(1).FooterStyle:id=101,.parent=3"
         _StyleDefs(106) =   "Splits(1).InactiveStyle:id=102,.parent=5"
         _StyleDefs(107) =   "Splits(1).SelectedStyle:id=104,.parent=6"
         _StyleDefs(108) =   "Splits(1).EditorStyle:id=103,.parent=7"
         _StyleDefs(109) =   "Splits(1).HighlightRowStyle:id=105,.parent=8"
         _StyleDefs(110) =   "Splits(1).EvenRowStyle:id=106,.parent=9"
         _StyleDefs(111) =   "Splits(1).OddRowStyle:id=107,.parent=10"
         _StyleDefs(112) =   "Splits(1).RecordSelectorStyle:id=109,.parent=11"
         _StyleDefs(113) =   "Splits(1).FilterBarStyle:id=110,.parent=12"
         _StyleDefs(114) =   "Splits(1).Columns(0).Style:id=114,.parent=99"
         _StyleDefs(115) =   "Splits(1).Columns(0).HeadingStyle:id=111,.parent=100"
         _StyleDefs(116) =   "Splits(1).Columns(0).FooterStyle:id=112,.parent=101"
         _StyleDefs(117) =   "Splits(1).Columns(0).EditorStyle:id=113,.parent=103"
         _StyleDefs(118) =   "Splits(1).Columns(1).Style:id=118,.parent=99"
         _StyleDefs(119) =   "Splits(1).Columns(1).HeadingStyle:id=115,.parent=100"
         _StyleDefs(120) =   "Splits(1).Columns(1).FooterStyle:id=116,.parent=101"
         _StyleDefs(121) =   "Splits(1).Columns(1).EditorStyle:id=117,.parent=103"
         _StyleDefs(122) =   "Splits(1).Columns(2).Style:id=122,.parent=99,.alignment=2"
         _StyleDefs(123) =   "Splits(1).Columns(2).HeadingStyle:id=119,.parent=100"
         _StyleDefs(124) =   "Splits(1).Columns(2).FooterStyle:id=120,.parent=101"
         _StyleDefs(125) =   "Splits(1).Columns(2).EditorStyle:id=121,.parent=103"
         _StyleDefs(126) =   "Splits(1).Columns(3).Style:id=126,.parent=99,.alignment=1"
         _StyleDefs(127) =   "Splits(1).Columns(3).HeadingStyle:id=123,.parent=100"
         _StyleDefs(128) =   "Splits(1).Columns(3).FooterStyle:id=124,.parent=101"
         _StyleDefs(129) =   "Splits(1).Columns(3).EditorStyle:id=125,.parent=103"
         _StyleDefs(130) =   "Splits(1).Columns(4).Style:id=130,.parent=99,.alignment=1"
         _StyleDefs(131) =   "Splits(1).Columns(4).HeadingStyle:id=127,.parent=100"
         _StyleDefs(132) =   "Splits(1).Columns(4).FooterStyle:id=128,.parent=101"
         _StyleDefs(133) =   "Splits(1).Columns(4).EditorStyle:id=129,.parent=103"
         _StyleDefs(134) =   "Splits(1).Columns(5).Style:id=134,.parent=99,.alignment=1"
         _StyleDefs(135) =   "Splits(1).Columns(5).HeadingStyle:id=131,.parent=100"
         _StyleDefs(136) =   "Splits(1).Columns(5).FooterStyle:id=132,.parent=101"
         _StyleDefs(137) =   "Splits(1).Columns(5).EditorStyle:id=133,.parent=103"
         _StyleDefs(138) =   "Splits(1).Columns(6).Style:id=138,.parent=99,.alignment=1"
         _StyleDefs(139) =   "Splits(1).Columns(6).HeadingStyle:id=135,.parent=100"
         _StyleDefs(140) =   "Splits(1).Columns(6).FooterStyle:id=136,.parent=101"
         _StyleDefs(141) =   "Splits(1).Columns(6).EditorStyle:id=137,.parent=103"
         _StyleDefs(142) =   "Splits(1).Columns(7).Style:id=142,.parent=99,.alignment=1"
         _StyleDefs(143) =   "Splits(1).Columns(7).HeadingStyle:id=139,.parent=100"
         _StyleDefs(144) =   "Splits(1).Columns(7).FooterStyle:id=140,.parent=101"
         _StyleDefs(145) =   "Splits(1).Columns(7).EditorStyle:id=141,.parent=103"
         _StyleDefs(146) =   "Splits(1).Columns(8).Style:id=146,.parent=99,.alignment=1"
         _StyleDefs(147) =   "Splits(1).Columns(8).HeadingStyle:id=143,.parent=100"
         _StyleDefs(148) =   "Splits(1).Columns(8).FooterStyle:id=144,.parent=101"
         _StyleDefs(149) =   "Splits(1).Columns(8).EditorStyle:id=145,.parent=103"
         _StyleDefs(150) =   "Splits(1).Columns(9).Style:id=150,.parent=99,.alignment=1"
         _StyleDefs(151) =   "Splits(1).Columns(9).HeadingStyle:id=147,.parent=100"
         _StyleDefs(152) =   "Splits(1).Columns(9).FooterStyle:id=148,.parent=101"
         _StyleDefs(153) =   "Splits(1).Columns(9).EditorStyle:id=149,.parent=103"
         _StyleDefs(154) =   "Splits(1).Columns(10).Style:id=154,.parent=99,.alignment=1"
         _StyleDefs(155) =   "Splits(1).Columns(10).HeadingStyle:id=151,.parent=100"
         _StyleDefs(156) =   "Splits(1).Columns(10).FooterStyle:id=152,.parent=101"
         _StyleDefs(157) =   "Splits(1).Columns(10).EditorStyle:id=153,.parent=103"
         _StyleDefs(158) =   "Splits(1).Columns(11).Style:id=158,.parent=99,.alignment=1"
         _StyleDefs(159) =   "Splits(1).Columns(11).HeadingStyle:id=155,.parent=100"
         _StyleDefs(160) =   "Splits(1).Columns(11).FooterStyle:id=156,.parent=101"
         _StyleDefs(161) =   "Splits(1).Columns(11).EditorStyle:id=157,.parent=103"
         _StyleDefs(162) =   "Splits(1).Columns(12).Style:id=162,.parent=99,.alignment=1"
         _StyleDefs(163) =   "Splits(1).Columns(12).HeadingStyle:id=159,.parent=100"
         _StyleDefs(164) =   "Splits(1).Columns(12).FooterStyle:id=160,.parent=101"
         _StyleDefs(165) =   "Splits(1).Columns(12).EditorStyle:id=161,.parent=103"
         _StyleDefs(166) =   "Splits(1).Columns(13).Style:id=166,.parent=99,.alignment=1"
         _StyleDefs(167) =   "Splits(1).Columns(13).HeadingStyle:id=163,.parent=100"
         _StyleDefs(168) =   "Splits(1).Columns(13).FooterStyle:id=164,.parent=101"
         _StyleDefs(169) =   "Splits(1).Columns(13).EditorStyle:id=165,.parent=103"
         _StyleDefs(170) =   "Splits(1).Columns(14).Style:id=170,.parent=99,.alignment=1"
         _StyleDefs(171) =   "Splits(1).Columns(14).HeadingStyle:id=167,.parent=100"
         _StyleDefs(172) =   "Splits(1).Columns(14).FooterStyle:id=168,.parent=101"
         _StyleDefs(173) =   "Splits(1).Columns(14).EditorStyle:id=169,.parent=103"
         _StyleDefs(174) =   "Splits(1).Columns(15).Style:id=182,.parent=99,.alignment=1"
         _StyleDefs(175) =   "Splits(1).Columns(15).HeadingStyle:id=179,.parent=100"
         _StyleDefs(176) =   "Splits(1).Columns(15).FooterStyle:id=180,.parent=101"
         _StyleDefs(177) =   "Splits(1).Columns(15).EditorStyle:id=181,.parent=103"
         _StyleDefs(178) =   "Splits(1).Columns(16).Style:id=174,.parent=99,.alignment=1"
         _StyleDefs(179) =   "Splits(1).Columns(16).HeadingStyle:id=171,.parent=100"
         _StyleDefs(180) =   "Splits(1).Columns(16).FooterStyle:id=172,.parent=101"
         _StyleDefs(181) =   "Splits(1).Columns(16).EditorStyle:id=173,.parent=103"
         _StyleDefs(182) =   "Named:id=33:Normal"
         _StyleDefs(183) =   ":id=33,.parent=0"
         _StyleDefs(184) =   "Named:id=34:Heading"
         _StyleDefs(185) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(186) =   ":id=34,.wraptext=-1"
         _StyleDefs(187) =   "Named:id=35:Footing"
         _StyleDefs(188) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(189) =   "Named:id=36:Selected"
         _StyleDefs(190) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(191) =   "Named:id=37:Caption"
         _StyleDefs(192) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(193) =   "Named:id=38:HighlightRow"
         _StyleDefs(194) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(195) =   "Named:id=39:EvenRow"
         _StyleDefs(196) =   ":id=39,.parent=33,.bgcolor=&HE9E9E9&"
         _StyleDefs(197) =   "Named:id=40:OddRow"
         _StyleDefs(198) =   ":id=40,.parent=33"
         _StyleDefs(199) =   "Named:id=41:RecordSelector"
         _StyleDefs(200) =   ":id=41,.parent=34"
         _StyleDefs(201) =   "Named:id=42:FilterBar"
         _StyleDefs(202) =   ":id=42,.parent=33"
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8205
      Width           =   12030
      _cx             =   21220
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      AutoSizeChildren=   8
      BorderWidth     =   1
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   6
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmSHU.frx":0000
      Begin BiSAButtonProject.BiSAButton cmdRefresh 
         Height          =   435
         Left            =   9795
         TabIndex        =   15
         Top             =   90
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   767
         Caption         =   "Refresh"
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10905
         TabIndex        =   16
         Top             =   90
         Width           =   1050
         _ExtentX        =   1852
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
         Picture         =   "frmSHU.frx":007D
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   8595
         TabIndex        =   20
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "frmSHU.frx":0123
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   2145
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12030
      _cx             =   21220
      _cy             =   3784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   10245
         TabIndex        =   13
         Top             =   645
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         Appearance      =   0
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
         Caption         =   "Tg"
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
      Begin BiSANumberBoxProject.BiSANumberBox nKeuntungan 
         Height          =   360
         Left            =   3360
         TabIndex        =   8
         Top             =   1665
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   635
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin BiSANumberBoxProject.BiSANumberBox nPinjaman 
         Height          =   360
         Left            =   3360
         TabIndex        =   9
         Top             =   540
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   635
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin BiSANumberBoxProject.BiSANumberBox nSimpananHarian 
         Height          =   360
         Left            =   3360
         TabIndex        =   10
         Top             =   915
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   635
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin BiSANumberBoxProject.BiSANumberBox nSimpananBerjangka 
         Height          =   360
         Left            =   3360
         TabIndex        =   11
         Top             =   1290
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   635
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   10245
         TabIndex        =   14
         Top             =   1005
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         Appearance      =   0
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
      Begin BiSANumberBoxProject.BiSANumberBox nKeuntunganModal 
         Height          =   360
         Left            =   6480
         TabIndex        =   19
         Top             =   540
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   635
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin BiSANumberBoxProject.BiSANumberBox nKeuntunganModalSukarela 
         Height          =   360
         Left            =   6480
         TabIndex        =   21
         Top             =   915
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   635
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin VB.Label Label7 
         Caption         =   "MODAL USAHA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   18
         Top             =   165
         Width           =   2100
      End
      Begin VB.Label Label6 
         Caption         =   "JASA USAHA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   17
         Top             =   150
         Width           =   1590
      End
      Begin VB.Label Label5 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6225
         TabIndex        =   12
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label Label4 
         Caption         =   "3. Porsi Keuntungan untuk Sim Berjangka"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   7
         Top             =   1335
         Width           =   2955
      End
      Begin VB.Label Label3 
         Caption         =   "2. Porsi Keuntungan untuk Sim Harian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   6
         Top             =   975
         Width           =   2940
      End
      Begin VB.Label Label2 
         Caption         =   "1. Porsi Keuntungan untuk Pinjaman"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   5
         Top             =   630
         Width           =   2940
      End
      Begin VB.Label Label1 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   4
         Top             =   1740
         Width           =   2940
      End
   End
End
Attribute VB_Name = "frmSHU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB
Dim vaSHU As New XArrayDB
Dim nLebarAwalKolomNamaAnggota As Double

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetRpt
End Sub

Private Sub cmdRefresh_Click()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim i As Integer
Dim nTotalAngsuran
Dim nSaldoCol4
Dim nSaldoCol5
Dim nSaldoCol8
Dim nSaldoCol9
Dim nSaldoCol15
Dim nSaldoCol16
Dim nJumlahAnggota As Double

  vaArray.ReDim 0, 0, 0, 16
  
  nKeuntungan.Value = GetLaba(dTgl(0).Value, dTgl(1).Value, True)
  nKeuntunganModal.Value = GetLaba(dTgl(0).Value, dTgl(1).Value, False)
  nKeuntunganModalSukarela.Value = GetLabaModalSukarela(dTgl(0).Value, dTgl(1).Value)
  nPinjaman.Value = nKeuntungan.Value * aCfg(msSHUPinjaman) / 100
  nSimpananHarian.Value = nKeuntungan.Value * aCfg(msSHUSimpanan) / 100
  nSimpananBerjangka.Value = nKeuntungan.Value * aCfg(msSHUDeposito) / 100
  
  nTotalAngsuran = GetTotalBungaAngsuran(dTgl(0).Value, dTgl(1).Value)
  
  Set db = objData.Browse(GetDSN, "registernasabah", , "jenisanggota", sisAssign, "1")
  If Not db.eof Then
    nJumlahAnggota = db.RecordCount
    Me.MousePointer = vbHourglass
    Do While Not db.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n
      vaArray(n, 1) = GetNull(db!Kode)
      vaArray(n, 2) = GetNull(db!nama)
      vaArray(n, 3) = nTotalAngsuran
      vaArray(n, 4) = GetKontribusiPinjaman(vaArray(n, 1), dTgl(0).Value, dTgl(1).Value)
      vaArray(n, 6) = Devide(vaArray(n, 4), vaArray(n, 3)) * 100
      vaArray(n, 5) = Devide(vaArray(n, 6), 100) * nPinjaman.Value
      
      vaArray(n, 7) = 0
      vaArray(n, 8) = GetKontribusiSimpanan(vaArray(n, 1), aCfg(msSHUKodeGolonganSimpananHarian, ""), dTgl(0).Value)
      vaArray(n, 9) = 0
      vaArray(n, 10) = 0
      vaArray(n, 11) = IIf(db!modalpenyertaan = 1, Devide(nKeuntunganModalSukarela.Value, 20), 0)
      vaArray(n, 12) = 0
      vaArray(n, 13) = 0
      vaArray(n, 14) = 0
      vaArray(n, 15) = Devide(nKeuntunganModal.Value, nJumlahAnggota)
      nSaldoCol4 = nSaldoCol4 + vaArray(n, 4)
      nSaldoCol5 = nSaldoCol5 + vaArray(n, 5)
      nSaldoCol8 = nSaldoCol8 + vaArray(n, 8)
      nSaldoCol15 = nSaldoCol15 + vaArray(n, 15)
      db.MoveNext
    Loop
  End If
  
  If vaArray.UpperBound(1) >= 0 Then
    For i = vaArray.LowerBound(1) + 1 To vaArray.UpperBound(1)
      vaArray(i, 7) = nSaldoCol8
      vaArray(i, 10) = Devide(vaArray(i, 8), vaArray(i, 7)) * 100
      vaArray(i, 9) = Devide(vaArray(i, 10), 100) * nSimpananHarian.Value
      nSaldoCol9 = nSaldoCol9 + vaArray(i, 9)
      vaArray(i, 16) = vaArray(i, 5) + vaArray(i, 9) + vaArray(i, 13) + vaArray(i, 15) + vaArray(i, 11)
      nSaldoCol16 = nSaldoCol16 + vaArray(i, 16)
    Next i
  End If
  
  Me.MousePointer = vbIconPointer
  TDBGrid1.Columns(4).FooterText = Format(nSaldoCol4, "###,###,###,##0.00")
  TDBGrid1.Columns(5).FooterText = Format(nSaldoCol5, "###,###,###,##0.00")
  TDBGrid1.Columns(8).FooterText = Format(nSaldoCol8, "###,###,###,##0.00")
  TDBGrid1.Columns(9).FooterText = Format(nSaldoCol9, "###,###,###,##0.00")
  TDBGrid1.Columns(15).FooterText = Format(nSaldoCol15, "###,###,###,##0.00")
  TDBGrid1.Columns(16).FooterText = Format(nSaldoCol16, "###,###,###,##0.00")
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Function GetTotalBungaAngsuran(dAwal As Date, dAkhir As Date) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim nTotal As Double
  
  nTotal = 0
  GetTotalBungaAngsuran = 0
  cSQL = "select a.rekening,sum(a.total) as total from angsuran a"
  cSQL = cSQL & " left join debitur d on d.rekening = a.rekening"
  cSQL = cSQL & " LEFT JOIN registernasabah r on r.kode = d.kode"
  cSQL = cSQL & " WHERE  a.tgl >= '" & Format(dAwal, "yyyy-MM-dd") & "' AND a.tgl <= '" & Format(dAkhir, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " AND r.jenisanggota = '1'"
  cSQL = cSQL & " GROUP BY d.kode"
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.eof Then
    Do While Not db.eof
      nTotal = nTotal + GetNull(db!Total)
      db.MoveNext
    Loop
  End If
  GetTotalBungaAngsuran = nTotal
End Function

Private Sub Form_Load()
Dim i As Integer

  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = EOM(Date)
  CenterForm Me
  nLebarAwalKolomNamaAnggota = TDBGrid1.Columns(2).Width
  Me.WindowState = 2
  For i = 3 To 16
    TDBGrid1.Splits(0).Columns(i).Visible = False
  Next i
End Sub

Private Function GetLaba(dAwal As Date, dAkhir As Date, lUsaha As Boolean) As Double
Dim db As New ADODB.Recordset

  GetLaba = 0
  Set db = objData.Browse(GetDSN, "bukubesar", "sum(kredit-debet) as laba", "tgl", sisGTEqual, Format(dAwal, "yyyy-MM-dd"), " and tgl <='" & Format(dAkhir, "yyyy-MM-dd") & "' AND (left(rekening,1) = '4' OR left(rekening,1) = '5')")
  If Not db.eof Then
'    If lUsaha Then
'      GetLaba = (GetNull(db!laba) + 2936287) * aCfg(msSHUJasaUsaha) / 100
'    Else
'      GetLaba = (GetNull(db!laba) + 2936287) * aCfg(msSHUModal) / 100
'    End If

    If lUsaha Then
'      GetLaba = (GetNull(db!laba) + 2936287) * aCfg(msSHUJasaUsaha) / 100
      GetLaba = 1040825
    Else
'      GetLaba = (GetNull(db!laba) + 2936287) * aCfg(msSHUModal) / 100
      GetLaba = 780619
    End If
  
  End If
End Function

Private Function GetLabaModalSukarela(dAwal As Date, dAkhir As Date) As Double
Dim db As New ADODB.Recordset

  GetLabaModalSukarela = 0
  Set db = objData.Browse(GetDSN, "bukubesar", "sum(kredit-debet) as laba", "tgl", sisGTEqual, Format(dAwal, "yyyy-MM-dd"), " and tgl <='" & Format(dAkhir, "yyyy-MM-dd") & "' AND (left(rekening,1) = '4' OR left(rekening,1) = '5')")
  If Not db.eof Then
    GetLabaModalSukarela = (GetNull(db!laba) + 2936287) * 5 / 100
  End If
End Function

Private Function GetKontribusiSimpanan(cKodeAnggota As String, GolonganTabungan As String, dAwal As Date) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim TotalSimpanan As Double
  
  GetKontribusiSimpanan = 0
  TotalSimpanan = 0
  cSQL = "SELECT t.rekening FROM tabungan t"
  cSQL = cSQL & " LEFT JOIN registernasabah r ON r.kode = t.kode"
  cSQL = cSQL & " LEFT JOIN golongantabungan g ON g.kode = '" & GolonganTabungan & "'"
  cSQL = cSQL & " WHERE r.kode = '" & cKodeAnggota & "'"
  cSQL = cSQL & " AND t.golongantabungan = '" & GolonganTabungan & "'"
  
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.eof Then
    Do While Not db.eof
      TotalSimpanan = TotalSimpanan + GetSaldo(GetNull(db!Rekening), Format(dTgl(0).Value, "yyyy-MM-dd"), Format(dTgl(1).Value, "yyyy-MM-dd"))
      db.MoveNext
    Loop
  End If
  GetKontribusiSimpanan = TotalSimpanan
End Function

Private Function GetKontribusiPinjaman(cKodeAnggota As String, dAwal As Date, dAkhir As Date) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim TotalPinjaman As Double
  
  GetKontribusiPinjaman = 0
  TotalPinjaman = 0
  cSQL = "SELECT SUM(total) as kontribusi FROM angsuran a"
  cSQL = cSQL & " LEFT JOIN debitur d ON d.rekening = a.rekening"
  cSQL = cSQL & " WHERE d.kode = '" & cKodeAnggota & "'"
  cSQL = cSQL & " AND a.tgl >= '" & Format(dAwal, "yyyy-MM-dd") & "' and a.tgl <= '" & Format(dAkhir, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " GROUP BY d.kode"
  
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.eof Then
    TotalPinjaman = GetNull(db!kontribusi)
  End If
  GetKontribusiPinjaman = TotalPinjaman
End Function

Private Sub Form_Resize()
Dim nSisaLebar As Double
Dim n As Integer

    Me.Refresh
    With TDBGrid1
      nSisaLebar = .Width - TDBGrid1.Columns(0).Width
      nSisaLebar = nSisaLebar - TDBGrid1.Columns(1).Width
      For n = 3 To 15
        nSisaLebar = nSisaLebar - TDBGrid1.Columns(n).Width
      Next n
    End With
    
    If Me.WindowState = 2 Then
      TDBGrid1.Columns(2).Width = nSisaLebar + 10000
    Else
      TDBGrid1.Columns(2).Width = nLebarAwalKolomNamaAnggota
    End If
End Sub

Private Function GetSaldo(ByVal cRekening As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As Double
Dim cWhere As String

  GetSaldo = 0
  cWhere = " and t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "' Group by t.Rekening,t.KodeTransaksi"
  Set dbData = objData.Browse(GetDSN, "MutasiTabungan t", "t.Rekening,k.DK,Sum(t.Jumlah) as Jumlah,t.KodeTransaksi", "t.rekening", sisAssign, cRekening, cWhere, , _
               Array("Left Join KodeTransaksi k on k.Kode = t.KodeTransaksi"))
  If Not dbData.eof Then
    Do While Not dbData.eof
      GetSaldo = GetSaldo + IIf((dbData!DK) = "D", -(dbData!Jumlah), (dbData!Jumlah))
      dbData.MoveNext
    Loop
  End If
End Function

'Private Function GetSaldo(ByVal cRekening As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As Double
'Dim dbSaldo As New ADODB.Recordset
'Dim dTgl As Date
'Dim dAkhirBulan As Date
'Dim nSaldoAwal As Double
'Dim cWhere As String
'Dim vaSaldo As New XArrayDB
'Dim n As Integer
'Dim nTemp As Double
'Dim cField As String
'Dim cTgl As String
'
'  GetSaldo = 0
'  dTgl = dTglAwal
'  dAkhirBulan = EOM(DateAdd("m", -1, dTgl))
'  'dAkhirBulan = EOM(dTglAwal)
'  cTgl = "Tgl <='" & Format(dAkhirBulan, "yyyy-mm-dd") & "'"
'  cField = " Sum(If(DK='D' and " & cTgl & " ,Jumlah,0)) as Debet, "
'  cField = cField & " Sum(If(DK='K' and " & cTgl & " ,Jumlah,0)) as Kredit"
'
'  Set dbSaldo = objData.Browse(GetDSN, "MutasiTabungan", cField, "Rekening", sisAssign, cRekening, , "Rekening,Tgl")
'  If Not dbSaldo.eof Then
'    nSaldoAwal = GetNull(dbSaldo!Kredit) - GetNull(dbSaldo!Debet)
'  End If
'
'  vaSaldo.ReDim 0, 0, 0, 2
'  vaSaldo(0, 0) = 0
'  vaSaldo(0, 1) = 0
'  vaSaldo(0, 2) = nSaldoAwal
'
'  dAkhirBulan = EOM(dTgl)
''  dAkhirBulan = dTglAkhir
'  cWhere = "And Tgl >='" & Format(dTgl, "yyyy-mm-dd") & "'"
'  cWhere = cWhere & "And Tgl <='" & Format(dAkhirBulan, "yyyy-mm-dd") & "'"
'  Set dbSaldo = objData.Browse(GetDSN, "MutasiTabungan", "Jumlah,DK", "Rekening", sisAssign, cRekening, cWhere, "Tgl")
'  If Not dbSaldo.eof Then
'    dbSaldo.MoveFirst
'    Do While Not dbSaldo.eof
'      vaSaldo.InsertRows vaSaldo.UpperBound(1) + 1
'      n = vaSaldo.UpperBound(1)
'      vaSaldo(n, 0) = IIf(GetNull(dbSaldo!DK) = "D", GetNull(dbSaldo!Jumlah), 0)
'      vaSaldo(n, 1) = IIf(GetNull(dbSaldo!DK) = "K", GetNull(dbSaldo!Jumlah), 0)
'      vaSaldo(n, 2) = vaSaldo(n - 1, 2) + vaSaldo(n, 1) - vaSaldo(n, 0)
'      dbSaldo.MoveNext
'    Loop
'  End If
'
'  nTemp = vaSaldo(0, 2)
'  For n = 1 To vaSaldo.UpperBound(1)
'    If vaSaldo(n, 2) < nTemp Then
'      nTemp = vaSaldo(n, 2)
'    Else
'      nTemp = nTemp
'    End If
'  Next
'  GetSaldo = nTemp
'End Function

Private Sub GetRpt()
Dim i As Integer
Dim n As Integer

On Error Resume Next

    vaSHU.ReDim 0, -1, 0, 5
    
    For i = vaArray.LowerBound(1) + 1 To vaArray.UpperBound(1)
      vaSHU.InsertRows vaSHU.UpperBound(1) + 1
      n = vaSHU.UpperBound(1)
      vaSHU(n, 0) = vaArray(i, 2)
      vaSHU(n, 1) = vaArray(i, 5)
      vaSHU(n, 2) = vaArray(i, 9)
'      vaSHU(n, 3) = vaArray(i, 13)
      vaSHU(n, 3) = vaArray(i, 11)
      vaSHU(n, 4) = vaArray(i, 15)
      vaSHU(n, 5) = vaArray(i, 16)
  
    Next i
    
    With FrmRPT
    .AddPageHeader "PEMBAGIAN SHU", tdbHalignCenter, , , , , 10
    .AddPageHeader "Periode " & Format(dTgl(0).Value, "dd/MM/yyyy") & "-" & Format(dTgl(1).Value, "dd/MM/yyyy"), tdbHalignCenter, , , True, , 10
    .AddPageHeader UCase(aCfg(msNama)), tdbHalignCenter, , , True, , 12
    .AddPageHeader "", , , , True
    .AddPageHeader "", , , , True
    
    .AddTableHeader "Nama"
    .AddTableHeader "SHU a Pinjaman", , , , 12
    .AddTableHeader "SHU a Sim Har", , , , 12
    .AddTableHeader "SHU a Sim Sukarela", , , , 12
    .AddTableHeader "SHU a Modal", , , , 12
    .AddTableHeader "Total SHU", , , , 12
    
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
        
        
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaSHU, True
  End With
End Sub
