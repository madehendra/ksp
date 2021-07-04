VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form RptLabaRugi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN LABA RUGI"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1200
      Left            =   0
      Top             =   0
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   2117
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
      Begin BiSANumberBoxProject.BiSANumberBox nLevel 
         Height          =   330
         Left            =   285
         TabIndex        =   4
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LEVEL"
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
      Begin BiSADateProject.BiSADate dAwal 
         Height          =   330
         Left            =   285
         TabIndex        =   0
         Top             =   195
         Width           =   3180
         _ExtentX        =   5609
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
         ForeColor       =   -2147483640
         Caption         =   "ANTARA TANGGAL"
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
      Begin BiSADateProject.BiSADate dAkhir 
         Height          =   330
         Left            =   3525
         TabIndex        =   1
         Top             =   195
         Width           =   1995
         _ExtentX        =   3519
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
         ForeColor       =   -2147483640
         Caption         =   "S.D"
         CaptionWidth    =   500
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
      Begin TrueDBReports60Ctl.TDBReports RptLaba 
         Height          =   570
         Left            =   4635
         TabIndex        =   5
         Top             =   540
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1005
         Caption         =   "Laba / Rugi"
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ErrorMsgCaption =   ""
         Filtered        =   0   'False
         DataMode        =   1
         DataMember      =   ""
         LinkSequence    =   1
         LinkOrder       =   0
         NameSubstitute  =   ""
         ConnectionString=   "DSN=Syariah"
         ConnectStringType=   3
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Syariah"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         CursorLocation  =   3
         ConnectionTimeout=   15
         CommandTimeout  =   30
         RecordSource    =   "Select 'Laporan Neraca' as Kode"
         CursorType      =   3
         CommandType     =   8
         MaxRecords      =   0
         LinkType        =   0
         Master          =   ""
         CallDataRead    =   0   'False
         ConvertNullToEmpty=   -1  'True
         DesignConnection=   -1  'True
         DesignTimeout   =   5
         UnitsOfMeasurement=   1
         Vedit_ShowGrid  =   -1  'True
         Vedit_SnapToGrid=   0   'False
         Vedit_GridUnitWidth=   160
         Vedit_GridUnitHeight=   160
         Vedit_ShowCellExpressions=   -1  'True
         Norm_rect_left  =   0
         Norm_rect_top   =   0
         Norm_rect_right =   0
         Norm_rect_bottom=   0
         Virgin          =   0   'False
         Fields.Count    =   6
         Fields(0).Name  =   "Golongan"
         Fields(0).DisplayName=   "Golongan"
         Fields(1).Name  =   "Kode"
         Fields(1).DisplayName=   "Kode"
         Fields(1).MaxLength=   20
         Fields(2).Name  =   "Keterangan"
         Fields(2).DisplayName=   "Keterangan"
         Fields(2).MaxLength=   50
         Fields(3).Name  =   "Awal"
         Fields(3).DisplayName=   "Awal"
         Fields(3).Type  =   5
         Fields(4).Name  =   "Mutasi"
         Fields(4).DisplayName=   "Mutasi"
         Fields(4).Type  =   5
         Fields(5).Name  =   "Akhir"
         Fields(5).DisplayName=   "Akhir"
         Sections.Count  =   6
         Sections(0).Name=   "ReportHeader"
         Sections(0).Condition=   "RecNo() = 0"
         Sections(0).StyleExp=   "tdb_RepHeader"
         Sections(0).Cells.Count=   3
         Sections(0).Cells(0).Name=   "ReportHeader"
         Sections(0).Cells(0).Exp=   """LAPORAN LABA / RUGI"""
         Sections(0).Cells(1).Name=   "Antara"
         Sections(0).Cells(1).Exp=   """~~Antara"""
         Sections(0).Cells(1).StyleExp=   "'tdb_RepPeriode'"
         Sections(0).Cells(1).NewLine=   -1  'True
         Sections(0).Cells(1).Width=   100
         Sections(0).Cells(1).CallExpression=   -1  'True
         Sections(0).Cells(2).Name=   "CELL_2"
         Sections(0).Cells(2).Exp=   """ """
         Sections(0).Cells(2).NewLine=   -1  'True
         Sections(1).Name=   "PageHeader"
         Sections(1).Type=   1
         Sections(1).StyleExp=   "tdb_PageHeader"
         Sections(1).Cells.Count=   1
         Sections(1).Cells(0).Name=   "PageNumber"
         Sections(1).Cells(0).Exp=   """Page "" & PageNo()"
         Sections(1).Cells(0).Placement=   2
         Sections(2).Name=   "DetailHeader"
         Sections(2).Type=   3
         Sections(2).StyleExp=   "tdb_TableHeader"
         Sections(2).Tabulator=   "Detail"
         Sections(2).Cells.Count=   5
         Sections(2).Cells(0).Name=   "Kode"
         Sections(2).Cells(0).Exp=   """Kode"""
         Sections(2).Cells(0).Width=   15
         Sections(2).Cells(1).Name=   "Keterangan"
         Sections(2).Cells(1).Exp=   """Keterangan"""
         Sections(2).Cells(1).Width=   50
         Sections(2).Cells(2).Name=   "Awal"
         Sections(2).Cells(2).Exp=   """1"""
         Sections(2).Cells(2).Width=   15
         Sections(2).Cells(2).CallExpression=   -1  'True
         Sections(2).Cells(3).Name=   "Mutasi"
         Sections(2).Cells(3).Exp=   """Mutasi"""
         Sections(2).Cells(3).Width=   15
         Sections(2).Cells(4).Name=   "Akhir"
         Sections(2).Cells(4).Exp=   """2"""
         Sections(2).Cells(4).CallExpression=   -1  'True
         Sections(3).Name=   "Detail"
         Sections(3).Type=   4
         Sections(3).Condition=   "Trim(Kode) <> """""
         Sections(3).StyleExp=   "'tdb_TableOddRow'"
         Sections(3).Cells.Count=   5
         Sections(3).Cells(0).Name=   "Kode"
         Sections(3).Cells(0).Exp=   "Kode"
         Sections(3).Cells(0).Width=   15
         Sections(3).Cells(1).Name=   "Keterangan"
         Sections(3).Cells(1).Exp=   "Keterangan"
         Sections(3).Cells(1).Width=   40
         Sections(3).Cells(2).Name=   "Awal"
         Sections(3).Cells(2).Exp=   "Awal"
         Sections(3).Cells(2).StyleExp=   "'tdb_Number'"
         Sections(3).Cells(2).Width=   15
         Sections(3).Cells(2).CallExpression=   -1  'True
         Sections(3).Cells(3).Name=   "Mutasi"
         Sections(3).Cells(3).Exp=   "Mutasi"
         Sections(3).Cells(3).StyleExp=   "'tdb_Number'"
         Sections(3).Cells(3).Width=   15
         Sections(3).Cells(3).CallExpression=   -1  'True
         Sections(3).Cells(4).Name=   "Akhir"
         Sections(3).Cells(4).Exp=   "Akhir"
         Sections(3).Cells(4).StyleExp=   "'tdb_Number'"
         Sections(3).Cells(4).Width=   15
         Sections(3).Cells(4).CallExpression=   -1  'True
         Sections(4).Name=   "TotalAktiva"
         Sections(4).Type=   4
         Sections(4).Condition=   "Trim(Kode) = """""
         Sections(4).StyleExp=   "'tdb_Total'"
         Sections(4).Tabulator=   "Detail"
         Sections(4).Cells.Count=   5
         Sections(4).Cells(0).Name=   "Kode"
         Sections(4).Cells(0).Exp=   """          "" & Keterangan"
         Sections(4).Cells(0).CellSpan=   2
         Sections(4).Cells(1).Name=   "Keterangan"
         Sections(4).Cells(2).Name=   "Awal"
         Sections(4).Cells(2).Exp=   "Awal"
         Sections(4).Cells(2).StyleExp=   "'tdb_NumberTotal'"
         Sections(4).Cells(2).CallExpression=   -1  'True
         Sections(4).Cells(3).Name=   "Mutasi"
         Sections(4).Cells(3).Exp=   "Mutasi"
         Sections(4).Cells(3).StyleExp=   "'tdb_NumberTotal'"
         Sections(4).Cells(3).CallExpression=   -1  'True
         Sections(4).Cells(4).Name=   "Akhir"
         Sections(4).Cells(4).Exp=   "Akhir"
         Sections(4).Cells(4).StyleExp=   "'tdb_NumberTotal'"
         Sections(4).Cells(4).CallExpression=   -1  'True
         Sections(5).Name=   "Pengesahan"
         Sections(5).Type=   5
         Sections(5).Condition=   "IsLastRec()"
         Sections(5).StyleExp=   "'tdb_Pengesahan'"
         Sections(5).Cells.Count=   3
         Sections(5).Cells(0).Name=   "CELL_0"
         Sections(5).Cells(0).Exp=   """Pembuat"""
         Sections(5).Cells(0).Width=   33
         Sections(5).Cells(1).Name=   "CELL_1"
         Sections(5).Cells(1).Exp=   """Pemeriksa"""
         Sections(5).Cells(1).Width=   33
         Sections(5).Cells(2).Name=   "CELL_2"
         Sections(5).Cells(2).Exp=   """Mengetahui"""
         Sections(5).Cells(2).Width=   33
         Styles.Count    =   25
         Styles(0).Name  =   "tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Charset=   0
         Styles(0).TextWrap=   0   'False
         Styles(1).Name  =   "tdb_PageHeader"
         Styles(1).ParentName=   "tdb_Base"
         Styles(1).Font_Charset=   0
         Styles(1).TextAlign=   2
         Styles(1).TextWrap=   0   'False
         Styles(1).fprops=   1
         Styles(2).Name  =   "tdb_PageFooter"
         Styles(2).ParentName=   "tdb_PageHeader"
         Styles(2).Font_Charset=   0
         Styles(2).TextWrap=   0   'False
         Styles(2).fprops=   0
         Styles(3).Name  =   "tdb_GroupHeaderBase"
         Styles(3).ParentName=   "tdb_Base"
         Styles(3).Font_Name=   "Arial"
         Styles(3).Font_Charset=   0
         Styles(3).TextWrap=   0   'False
         Styles(3).fprops=   2097152
         Styles(4).Name  =   "tdb_GroupHeader1"
         Styles(4).ParentName=   "tdb_GroupHeaderBase"
         Styles(4).Font_Size=   14
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).TextWrap=   0   'False
         Styles(4).fprops=   20971520
         Styles(5).Name  =   "tdb_GroupFooterBase"
         Styles(5).ParentName=   "tdb_Base"
         Styles(5).Font_Name=   "Arial"
         Styles(5).Font_Charset=   0
         Styles(5).TextAlign=   2
         Styles(5).TextWrap=   0   'False
         Styles(5).fprops=   2097153
         Styles(6).Name  =   "tdb_GroupFooter1"
         Styles(6).ParentName=   "tdb_GroupFooterBase"
         Styles(6).Font_Size=   14
         Styles(6).Font_Bold=   -1  'True
         Styles(6).Font_Charset=   0
         Styles(6).TextWrap=   0   'False
         Styles(6).fprops=   20971520
         Styles(7).Name  =   "tdb_GroupHeader2"
         Styles(7).ParentName=   "tdb_GroupHeaderBase"
         Styles(7).Font_Size=   14
         Styles(7).Font_Charset=   0
         Styles(7).TextWrap=   0   'False
         Styles(7).fprops=   4194304
         Styles(8).Name  =   "tdb_GroupFooter2"
         Styles(8).ParentName=   "tdb_GroupFooterBase"
         Styles(8).Font_Size=   14
         Styles(8).Font_Charset=   0
         Styles(8).TextWrap=   0   'False
         Styles(8).fprops=   4194304
         Styles(9).Name  =   "tdb_GroupHeader3"
         Styles(9).ParentName=   "tdb_GroupHeaderBase"
         Styles(9).Font_Size=   12
         Styles(9).Font_Bold=   -1  'True
         Styles(9).Font_Charset=   0
         Styles(9).TextWrap=   0   'False
         Styles(9).fprops=   20971520
         Styles(10).Name =   "tdb_GroupFooter3"
         Styles(10).ParentName=   "tdb_GroupFooterBase"
         Styles(10).Font_Size=   12
         Styles(10).Font_Bold=   -1  'True
         Styles(10).Font_Charset=   0
         Styles(10).TextWrap=   0   'False
         Styles(10).fprops=   20971520
         Styles(11).Name =   "tdb_GroupHeader4"
         Styles(11).ParentName=   "tdb_GroupHeaderBase"
         Styles(11).Font_Size=   12
         Styles(11).Font_Charset=   0
         Styles(11).TextWrap=   0   'False
         Styles(11).fprops=   4194304
         Styles(12).Name =   "tdb_GroupFooter4"
         Styles(12).ParentName=   "tdb_GroupFooterBase"
         Styles(12).Font_Size=   12
         Styles(12).Font_Charset=   0
         Styles(12).TextWrap=   0   'False
         Styles(12).fprops=   4194304
         Styles(13).Name =   "tdb_TableBase"
         Styles(13).ParentName=   "tdb_Base"
         Styles(13).Font_Name=   "Arial"
         Styles(13).Font_Charset=   0
         Styles(13).TextWrap=   0   'False
         Styles(13).BorderHT=   "tdb_ThinBlack"
         Styles(13).BorderHI=   "tdb_Invisible"
         Styles(13).BorderHB=   "tdb_ThinBlack"
         Styles(13).BorderVL=   "tdb_ThinBlack"
         Styles(13).BorderVI=   "tdb_ThinGray"
         Styles(13).BorderVR=   "tdb_ThinBlack"
         Styles(13).fprops=   4161536
         Styles(14).Name =   "tdb_TableOddRow"
         Styles(14).ParentName=   "tdb_TableBase"
         Styles(14).Font_Charset=   0
         Styles(14).TextWrap=   0   'False
         Styles(14).BorderHI=   "Inner"
         Styles(14).BorderVI=   "tdb_ThinBlack"
         Styles(14).fprops=   4784128
         Styles(15).Name =   "tdb_TableEvenRow"
         Styles(15).ParentName=   "tdb_TableOddRow"
         Styles(15).Font_Charset=   0
         Styles(15).TextWrap=   0   'False
         Styles(15).BackColor=   8454143
         Styles(15).NoFill=   0   'False
         Styles(15).fprops=   48
         Styles(16).Name =   "tdb_TableHeader"
         Styles(16).ParentName=   "tdb_TableBase"
         Styles(16).Font_Bold=   -1  'True
         Styles(16).Font_Charset=   0
         Styles(16).TextAlign=   1
         Styles(16).TextWrap=   0   'False
         Styles(16).BackColor=   15132390
         Styles(16).NoFill=   0   'False
         Styles(16).BorderHT=   "tdb_ThinBlack"
         Styles(16).BorderHI=   "tdb_ThinBlack"
         Styles(16).BorderHB=   "tdb_ThinBlack"
         Styles(16).BorderVL=   "tdb_ThinBlack"
         Styles(16).BorderVI=   "tdb_ThinBlack"
         Styles(16).BorderVR=   "tdb_ThinBlack"
         Styles(16).fprops=   23035961
         Styles(17).Name =   "tdb_TableFiller"
         Styles(17).ParentName=   "tdb_TableOddRow"
         Styles(17).Font_Charset=   0
         Styles(17).TextWrap=   0   'False
         Styles(17).MarginTop=   0
         Styles(17).MarginBottom=   0
         Styles(17).fprops=   20480
         Styles(18).Name =   "tdb_RepHeader"
         Styles(18).ParentName=   "tdb_Base"
         Styles(18).Font_Name=   "Arial"
         Styles(18).Font_Bold=   -1  'True
         Styles(18).Font_Charset=   0
         Styles(18).TextAlign=   1
         Styles(18).TextWrap=   0   'False
         Styles(18).fprops=   23068673
         Styles(19).Name =   "tdb_Total"
         Styles(19).ParentName=   "tdb_TableBase"
         Styles(19).Font_Bold=   -1  'True
         Styles(19).Font_Charset=   0
         Styles(19).TextWrap=   0   'False
         Styles(19).BackColor=   15132390
         Styles(19).NoFill=   0   'False
         Styles(19).BorderHT=   "tdb_ThinBlack"
         Styles(19).BorderHI=   "tdb_ThinBlack"
         Styles(19).BorderHB=   "tdb_ThinBlack"
         Styles(19).BorderVL=   "tdb_ThinBlack"
         Styles(19).BorderVI=   "tdb_ThinBlack"
         Styles(19).BorderVR=   "tdb_ThinBlack"
         Styles(19).fprops=   18841648
         Styles(20).Name =   "tdb_Number"
         Styles(20).ParentName=   "tdb_TableOddRow"
         Styles(20).Font_Charset=   0
         Styles(20).TextAlign=   2
         Styles(20).TextWrap=   0   'False
         Styles(20).fprops=   1
         Styles(21).Name =   "tdb_Number_Negative"
         Styles(21).ParentName=   "tdb_Number"
         Styles(21).Font_Charset=   0
         Styles(21).TextWrap=   0   'False
         Styles(21).ForeColor=   255
         Styles(21).fprops=   40
         Styles(22).Name =   "tdb_NumberTotal"
         Styles(22).ParentName=   "tdb_Total"
         Styles(22).Font_Charset=   0
         Styles(22).TextAlign=   2
         Styles(22).TextWrap=   0   'False
         Styles(22).fprops=   4194305
         Styles(23).Name =   "tdb_RepPeriode"
         Styles(23).ParentName=   "tdb_RepHeader"
         Styles(23).Font_Size=   9
         Styles(23).Font_Charset=   0
         Styles(23).TextWrap=   0   'False
         Styles(23).fprops=   4194304
         Styles(24).Name =   "tdb_Pengesahan"
         Styles(24).ParentName=   "tdb_Base"
         Styles(24).Font_Name=   "Arial"
         Styles(24).Font_Charset=   0
         Styles(24).TextAlign=   1
         Styles(24).TextWrap=   0   'False
         Styles(24).fprops=   2097153
         Lines.Count     =   4
         Lines(0).Name   =   "tdb_Invisible"
         Lines(1).Name   =   "tdb_ThinBlack"
         Lines(1).Thickness=   2
         Lines(2).Name   =   "tdb_ThinGray"
         Lines(2).Thickness=   2
         Lines(2).Color  =   8421504
         Lines(3).Name   =   "Inner"
         Lines(3).Thickness=   1
         Lines(3).Color  =   4210752
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginTop=   720
         Profiles(0).PrinterMarginBottom=   720
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "( 1-4 )"
         Height          =   240
         Left            =   2625
         TabIndex        =   6
         Top             =   645
         Width           =   525
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1185
      Width           =   6480
      _ExtentX        =   11430
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Height          =   435
         Left            =   5250
         TabIndex        =   2
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "RptLabaRugi.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4080
         TabIndex        =   3
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "RptLabaRugi.frx":00A6
      End
   End
End
Attribute VB_Name = "RptLabaRugi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaLaba As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim cSQL As String
Dim n As Long
Dim cRekening As String
Dim nJumlahAwal As Double
Dim nJumlahMutasi As Double
Dim nJumlahAkhir As Double
Dim nTotalAwal As Double
Dim nTotalMutasi As Double
Dim nTotalAkhir As Double

  
  cSQL = cSQL & " Select '1' as Golongan,'  ','Jumlah Pendapatan Operasional     ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select '3' as Golongan,'  ','Jumlah Biaya Operasional          ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select '4' as Golongan,'  ','Laba / Rugi Operasional           ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select '6' as Golongan,'  ','Jumlah Pendapatan Non Operasional ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select '8' as Golongan,'  ','Jumlah Biaya Non Operasional      ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select '9' as Golongan,'  ','Laba / Rugi Bersih                ' as Keterangan,0,0,0 "
  Set dbData = objData.SQL(GetDSN, cSQL)
  vaLaba.LoadRows dbData.GetRows(dbData.RecordCount)
  
  ' Ambil Data Rekening
  Set dbData = objData.Browse(GetDSN, "Rekening", "Kode,Keterangan", "Kode", sisGTEqual, "4", " and Kode <= '6'")
  If dbData.RecordCount > 0 Then
    dbData.MoveFirst
    Do While Not dbData.eof
      If Level(GetNull(dbData!Kode, "")) <= nLevel.Value Then
        vaLaba.InsertRows vaLaba.UpperBound(1) + 1
        n = vaLaba.UpperBound(1)
        
        vaLaba(n, 0) = GetGolongan(GetNull(dbData!Kode, ""))
        vaLaba(n, 1) = GetNull(dbData!Kode, "")
        vaLaba(n, 2) = (RekSpace(GetNull(dbData!Kode, ""), GetNull(dbData!Keterangan, "")))
        vaLaba(n, 3) = 0
        vaLaba(n, 4) = 0
        vaLaba(n, 5) = 0
      End If
      dbData.MoveNext
    Loop
  End If
  
  ' Ambil Data Saldo Awal
  Set dbData = objData.Browse(GetDSN, "SaldoRekening")
  If Not dbData.eof Then
    Do While Not dbData.eof
      If TypeRekening(GetNull(dbData!Rekening, "")) = SisBiaya Or TypeRekening(GetNull(dbData!Rekening, "")) = SisPendapatan Then
        cRekening = GetLevel(GetNull(dbData!Rekening, ""), nLevel.Value)
        n = vaLaba.Find(0, 1, cRekening)
        If n >= 0 Then
          vaLaba(n, 3) = vaLaba(n, 3) + GS(GetNull(dbData!Rekening, ""), GetNull(dbData!Awal))
          vaLaba(n, 5) = vaLaba(n, 3) + vaLaba(n, 4)
        End If
      End If
      dbData.MoveNext
    Loop
  End If
  
  ' Ambil Data Pada Buku Besar
  Set dbData = objData.Browse(GetDSN, "BukuBesar", "Rekening,Tgl,Debet,Kredit", "Posting", sisAssign, "0", " and Tgl <= '" & Format(dAkhir.Value, "yyyy-mm-dd") & "'")
  If Not dbData.eof Then
    Do While Not dbData.eof
      cRekening = GetLevel(GetNull(dbData!Rekening, ""), nLevel.Value)
      If TypeRekening(GetNull(dbData!Rekening, "")) = SisBiaya Or TypeRekening(GetNull(dbData!Rekening, "")) = SisPendapatan Then
        n = vaLaba.Find(0, 1, cRekening)
        If n >= 0 Then
          ' Untuk Sementera Mutasi tanggal 1 januari di anggap saldo awal supaya pada laporan Laba Rugi Benar
          If dbData!Tgl < dAwal.Value Or Format(GetNull(dbData!Tgl), "ddMM") = "0101" Then
            vaLaba(n, 3) = vaLaba(n, 3) + GS(dbData!Rekening, dbData!Debet) - GS(GetNull(dbData!Rekening, ""), GetNull(dbData!Kredit))
          Else
            vaLaba(n, 4) = vaLaba(n, 4) + GS(dbData!Rekening, dbData!Debet) - GS(GetNull(dbData!Rekening, ""), GetNull(dbData!Kredit))
          End If
          vaLaba(n, 5) = vaLaba(n, 3) + vaLaba(n, 4)
        End If
      End If
      dbData.MoveNext
    Loop
  End If
  
  

  vaLaba.QuickSort vaLaba.LowerBound(1), vaLaba.UpperBound(1), 0, XORDER_ASCEND, XTYPE_DEFAULT, 1, XORDER_ASCEND, XTYPE_DEFAULT
  For n = 0 To vaLaba.UpperBound(1)
    nJumlahAwal = nJumlahAwal + vaLaba(n, 3)
    nJumlahMutasi = nJumlahMutasi + vaLaba(n, 4)
    nJumlahAkhir = nJumlahAwal + nJumlahMutasi
    
    If Trim(vaLaba(n, 1)) = "" Then
      If vaLaba(n, 0) = "1" Or vaLaba(n, 0) = "3" Or vaLaba(n, 0) = "6" Or vaLaba(n, 0) = "8" Then
        vaLaba(n, 3) = nJumlahAwal
        vaLaba(n, 4) = nJumlahMutasi
        vaLaba(n, 5) = nJumlahAkhir
      Else
        vaLaba(n, 3) = nTotalAwal
        vaLaba(n, 4) = nTotalMutasi
        vaLaba(n, 5) = nTotalAkhir
      End If
      
      If vaLaba(n, 0) = "3" Or vaLaba(n, 0) = "8" Then
        ' Pendapatan - Biaya
        nTotalAwal = nTotalAwal - nJumlahAwal
        nTotalMutasi = nTotalMutasi - nJumlahMutasi
        nTotalAkhir = nTotalAkhir - nJumlahAkhir
      Else
        nTotalAwal = nTotalAwal + nJumlahAwal
        nTotalMutasi = nTotalMutasi + nJumlahMutasi
        nTotalAkhir = nTotalAkhir + nJumlahAkhir
      End If
      
      nJumlahAwal = 0
      nJumlahMutasi = 0
      nJumlahAkhir = 0
    End If
  Next
  
  Set RptLaba.Array = vaLaba
  RptLaba.Refresh
  With RptLaba
    .Profiles(0).PrinterMarginBottom = 720
    .Profiles(0).PrinterMarginLeft = 720
    .Profiles(0).PrinterMarginRight = 720
    .Profiles(0).PrinterMarginTop = 720
    .PrintPreview
  End With
End Sub

Private Function GetGolongan(ByVal cRekening As String) As String
  Select Case left(cRekening, 3)
    Case "4.1"
      GetGolongan = "0"
    Case "5.1"
      GetGolongan = "2"
    Case "5.2"
      GetGolongan = "2"
    Case "4.2"
      GetGolongan = "5"
    Case "5.3"
      GetGolongan = "7"
  End Select
End Function

Private Function GS(ByVal cRekening As String, ByVal nValue As Double) As Double
  GS = IIf(left(cRekening, 1) = "5", nValue, -nValue)
End Function

Private Sub Form_Load()
Dim n As Single
  CenterForm Me
  
  dAwal.Value = Date
  dAkhir.Value = Date
  nLevel.Value = 1
  TabIndex dAwal, n
  TabIndex dAkhir, n
  TabIndex nLevel, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub RptLaba_CellExpression(ByVal Section As Integer, ByVal Cell As Integer, Value As Variant)
Dim cChar As String
  If Section = 0 Then
    Value = "Periode : " & Format(dAwal.Value - 1, "dd mmmm yyyy") & " Dan " & Format(dAkhir.Value, "dd mmmm yyyy")
  ElseIf Section = 2 Then
    If Value = "1" Then
      Value = Format(dAwal.Value - 1, "dd-MM-yyyy")
    Else
      Value = Format(dAkhir.Value, "dd-MM-yyyy")
    End If
  Else
    Value = GetNull(Value, 0)
    cChar = IIf(Value < 0, "()", "  ")
    If Round(Value, 2) = 0 Then
      Value = ""
    Else
      Value = Format(Abs(GetNull(Value)), "###,###,###,###,##0.00")
      Value = left(cChar, 1) & Value & Right(cChar, 1)
    End If
  End If
End Sub



