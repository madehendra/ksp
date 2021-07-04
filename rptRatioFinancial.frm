VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form rptRatioFinancial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ratio Financial"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   2595
   Begin BiSAButtonProject.BiSAButton cmdOK 
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   135
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   688
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
   Begin TrueDBReports60Ctl.TDBReports RptNeraca 
      Height          =   570
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1005
      Caption         =   "Neraca"
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
      Sections.Count  =   1
      Sections(0).Name=   "SECTION_0"
      Styles.Count    =   25
      Styles(0).Name  =   "tdb_Base"
      Styles(0).ParentName=   ""
      Styles(0).Font_Size=   9.75
      Styles(0).Font_Charset=   0
      Styles(0).NoClipping=   -1  'True
      Styles(1).Name  =   "tdb_PageHeader"
      Styles(1).ParentName=   "tdb_Base"
      Styles(1).Font_Size=   9.75
      Styles(1).Font_Charset=   0
      Styles(1).TextAlign=   2
      Styles(1).NoClipping=   -1  'True
      Styles(1).fprops=   1
      Styles(2).Name  =   "tdb_PageFooter"
      Styles(2).ParentName=   "tdb_PageHeader"
      Styles(2).Font_Size=   9.75
      Styles(2).Font_Charset=   0
      Styles(2).NoClipping=   -1  'True
      Styles(2).fprops=   0
      Styles(3).Name  =   "tdb_GroupHeaderBase"
      Styles(3).ParentName=   "tdb_Base"
      Styles(3).Font_Name=   "Arial"
      Styles(3).Font_Size=   9.75
      Styles(3).Font_Charset=   0
      Styles(3).NoClipping=   -1  'True
      Styles(3).fprops=   2097152
      Styles(4).Name  =   "tdb_GroupHeader1"
      Styles(4).ParentName=   "tdb_GroupHeaderBase"
      Styles(4).Font_Size=   14
      Styles(4).Font_Bold=   -1  'True
      Styles(4).Font_Charset=   0
      Styles(4).NoClipping=   -1  'True
      Styles(4).fprops=   20971520
      Styles(5).Name  =   "tdb_GroupFooterBase"
      Styles(5).ParentName=   "tdb_Base"
      Styles(5).Font_Name=   "Arial"
      Styles(5).Font_Size=   9.75
      Styles(5).Font_Charset=   0
      Styles(5).TextAlign=   2
      Styles(5).NoClipping=   -1  'True
      Styles(5).fprops=   2097153
      Styles(6).Name  =   "tdb_GroupFooter1"
      Styles(6).ParentName=   "tdb_GroupFooterBase"
      Styles(6).Font_Size=   14
      Styles(6).Font_Bold=   -1  'True
      Styles(6).Font_Charset=   0
      Styles(6).NoClipping=   -1  'True
      Styles(6).fprops=   20971520
      Styles(7).Name  =   "tdb_GroupHeader2"
      Styles(7).ParentName=   "tdb_GroupHeaderBase"
      Styles(7).Font_Size=   14
      Styles(7).Font_Charset=   0
      Styles(7).NoClipping=   -1  'True
      Styles(7).fprops=   4194304
      Styles(8).Name  =   "tdb_GroupFooter2"
      Styles(8).ParentName=   "tdb_GroupFooterBase"
      Styles(8).Font_Size=   14
      Styles(8).Font_Charset=   0
      Styles(8).NoClipping=   -1  'True
      Styles(8).fprops=   4194304
      Styles(9).Name  =   "tdb_GroupHeader3"
      Styles(9).ParentName=   "tdb_GroupHeaderBase"
      Styles(9).Font_Size=   12
      Styles(9).Font_Bold=   -1  'True
      Styles(9).Font_Charset=   0
      Styles(9).NoClipping=   -1  'True
      Styles(9).fprops=   20971520
      Styles(10).Name =   "tdb_GroupFooter3"
      Styles(10).ParentName=   "tdb_GroupFooterBase"
      Styles(10).Font_Size=   12
      Styles(10).Font_Bold=   -1  'True
      Styles(10).Font_Charset=   0
      Styles(10).NoClipping=   -1  'True
      Styles(10).fprops=   20971520
      Styles(11).Name =   "tdb_GroupHeader4"
      Styles(11).ParentName=   "tdb_GroupHeaderBase"
      Styles(11).Font_Size=   12
      Styles(11).Font_Charset=   0
      Styles(11).NoClipping=   -1  'True
      Styles(11).fprops=   4194304
      Styles(12).Name =   "tdb_GroupFooter4"
      Styles(12).ParentName=   "tdb_GroupFooterBase"
      Styles(12).Font_Size=   12
      Styles(12).Font_Charset=   0
      Styles(12).NoClipping=   -1  'True
      Styles(12).fprops=   4194304
      Styles(13).Name =   "tdb_TableBase"
      Styles(13).ParentName=   "tdb_Base"
      Styles(13).Font_Name=   "Arial"
      Styles(13).Font_Size=   9.75
      Styles(13).Font_Charset=   0
      Styles(13).BorderHT=   "tdb_ThinBlack"
      Styles(13).BorderHI=   "tdb_Invisible"
      Styles(13).BorderHB=   "tdb_ThinBlack"
      Styles(13).BorderVL=   "tdb_ThinBlack"
      Styles(13).BorderVI=   "tdb_ThinGray"
      Styles(13).BorderVR=   "tdb_ThinBlack"
      Styles(13).NoClipping=   -1  'True
      Styles(13).fprops=   4161536
      Styles(14).Name =   "tdb_TableOddRow"
      Styles(14).ParentName=   "tdb_TableBase"
      Styles(14).Font_Charset=   0
      Styles(14).BorderHI=   "Inner"
      Styles(14).BorderVI=   "tdb_ThinBlack"
      Styles(14).NoClipping=   -1  'True
      Styles(14).fprops=   4784128
      Styles(15).Name =   "tdb_TableEvenRow"
      Styles(15).ParentName=   "tdb_TableOddRow"
      Styles(15).Font_Size=   9.75
      Styles(15).Font_Charset=   0
      Styles(15).BackColor=   8454143
      Styles(15).NoFill=   0   'False
      Styles(15).NoClipping=   -1  'True
      Styles(15).fprops=   48
      Styles(16).Name =   "tdb_TableHeader"
      Styles(16).ParentName=   "tdb_TableBase"
      Styles(16).Font_Bold=   -1  'True
      Styles(16).Font_Charset=   0
      Styles(16).TextAlign=   1
      Styles(16).BackColor=   15132390
      Styles(16).NoFill=   0   'False
      Styles(16).BorderHT=   "tdb_ThinBlack"
      Styles(16).BorderHI=   "tdb_ThinBlack"
      Styles(16).BorderHB=   "tdb_ThinBlack"
      Styles(16).BorderVL=   "tdb_ThinBlack"
      Styles(16).BorderVI=   "tdb_ThinBlack"
      Styles(16).BorderVR=   "tdb_ThinBlack"
      Styles(16).NoClipping=   -1  'True
      Styles(16).fprops=   23035961
      Styles(17).Name =   "tdb_TableFiller"
      Styles(17).ParentName=   "tdb_TableOddRow"
      Styles(17).Font_Size=   9.75
      Styles(17).Font_Charset=   0
      Styles(17).MarginTop=   0
      Styles(17).MarginBottom=   0
      Styles(17).NoClipping=   -1  'True
      Styles(17).fprops=   20480
      Styles(18).Name =   "tdb_RepHeader"
      Styles(18).ParentName=   "tdb_Base"
      Styles(18).Font_Name=   "Arial"
      Styles(18).Font_Bold=   -1  'True
      Styles(18).Font_Charset=   0
      Styles(18).TextAlign=   1
      Styles(18).NoClipping=   -1  'True
      Styles(18).fprops=   1096810497
      Styles(19).Name =   "tdb_Total"
      Styles(19).ParentName=   "tdb_TableBase"
      Styles(19).Font_Size=   9.75
      Styles(19).Font_Bold=   -1  'True
      Styles(19).Font_Charset=   0
      Styles(19).BackColor=   15132390
      Styles(19).NoFill=   0   'False
      Styles(19).BorderHT=   "tdb_ThinBlack"
      Styles(19).BorderHI=   "tdb_ThinBlack"
      Styles(19).BorderHB=   "tdb_ThinBlack"
      Styles(19).BorderVL=   "tdb_ThinBlack"
      Styles(19).BorderVI=   "tdb_ThinBlack"
      Styles(19).BorderVR=   "tdb_ThinBlack"
      Styles(19).NoClipping=   -1  'True
      Styles(19).fprops=   18841648
      Styles(20).Name =   "tdb_Number"
      Styles(20).ParentName=   "tdb_TableOddRow"
      Styles(20).Font_Size=   9.75
      Styles(20).Font_Charset=   0
      Styles(20).TextAlign=   2
      Styles(20).NoClipping=   -1  'True
      Styles(20).fprops=   1
      Styles(21).Name =   "tdb_Number_Negative"
      Styles(21).ParentName=   "tdb_Number"
      Styles(21).Font_Size=   9.75
      Styles(21).Font_Charset=   0
      Styles(21).ForeColor=   255
      Styles(21).NoClipping=   -1  'True
      Styles(21).fprops=   40
      Styles(22).Name =   "tdb_NumberTotal"
      Styles(22).ParentName=   "tdb_Total"
      Styles(22).Font_Charset=   0
      Styles(22).TextAlign=   2
      Styles(22).NoClipping=   -1  'True
      Styles(22).fprops=   4194305
      Styles(23).Name =   "tdb_RepPeriode"
      Styles(23).ParentName=   "tdb_RepHeader"
      Styles(23).Font_Size=   9
      Styles(23).Font_Charset=   0
      Styles(23).NoClipping=   -1  'True
      Styles(23).fprops=   4194304
      Styles(24).Name =   "tdb_Pengesahan"
      Styles(24).ParentName=   "tdb_Base"
      Styles(24).Font_Name=   "Arial"
      Styles(24).Font_Size=   9.75
      Styles(24).Font_Charset=   0
      Styles(24).TextAlign=   1
      Styles(24).NoClipping=   -1  'True
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
End
Attribute VB_Name = "rptRatioFinancial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cmdOK_Click()
  MsgBox Format(GetSolvabilitas, "###,###,###,##0.00")
End Sub

Private Function GetSolvabilitas() As Double
Dim aktiva As Double
Dim hutang As Double

  aktiva = 0
  hutang = 0
  
  Set dbData = objData.Browse(GetDSN, "saldorekening", "sum(awal) as totalaktiva", "left(rekening,1)", sisAssign, "1")
  If Not dbData.eof Then
    aktiva = GetNull(dbData!totalaktiva)
  End If
  
  
  Set dbData = objData.Browse(GetDSN, "bukubesar", "sum(debet-kredit) as totalaktiva", "left(rekening,1)", sisAssign, "1", " and year(tgl) = '2007'")
  If Not dbData.eof Then
    aktiva = aktiva + GetNull(dbData!totalaktiva)
  End If
  
  
  Set dbData = objData.Browse(GetDSN, "saldorekening", "sum(awal) as totalaktiva", "left(rekening,1)", sisAssign, "2")
  If Not dbData.eof Then
    hutang = GetNull(dbData!totalaktiva)
  End If
  
  Set dbData = objData.Browse(GetDSN, "bukubesar", "sum(kredit-debet) as totalhutang", "left(rekening,1)", sisAssign, "2", " and year(tgl) = '2007'")
  If Not dbData.eof Then
    hutang = -hutang + GetNull(dbData!totalhutang)
  End If
  
  GetSolvabilitas = Devide(aktiva, hutang) * 100
End Function
