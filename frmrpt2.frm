VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Begin VB.Form frmrpt2 
   Caption         =   "Form1"
   ClientHeight    =   615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   1740
   StartUpPosition =   3  'Windows Default
   Begin TrueDBReports60Ctl.TDBReports Rpt 
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1005
      Caption         =   "TDBReports1"
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
      ConnectionString=   ""
      ConnectStringType=   1
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      CursorLocation  =   2
      ConnectionTimeout=   15
      CommandTimeout  =   30
      RecordSource    =   ""
      CursorType      =   3
      CommandType     =   8
      MaxRecords      =   0
      LinkType        =   0
      Master          =   ""
      CallDataRead    =   0   'False
      ConvertNullToEmpty=   -1  'True
      DesignConnection=   -1  'True
      DesignTimeout   =   5
      UnitsOfMeasurement=   4
      Vedit_ShowGrid  =   -1  'True
      Vedit_SnapToGrid=   0   'False
      Vedit_GridUnitWidth=   2.82216666666667
      Vedit_GridUnitHeight=   2.82216666666667
      Vedit_ShowCellExpressions=   -1  'True
      Norm_rect_left  =   0
      Norm_rect_top   =   0
      Norm_rect_right =   0
      Norm_rect_bottom=   0
      Virgin          =   0   'False
      Profiles.Count  =   1
      Profiles(0).Name=   "PROFILE_0"
      Profiles(0).Active=   -1  'True
      Profiles(0).PreviewNoMinimize=   -1  'True
      Profiles(0).PreviewNoMaximize=   -1  'True
      Profiles(0).PreviewNoResize=   -1  'True
      Profiles(0).PreviewMaximized=   -1  'True
      Profiles(0).PrinterMarginLeft=   10
      Profiles(0).PrinterMarginTop=   10
      Profiles(0).PrinterMarginRight=   10
      Profiles(0).PrinterMarginBottom=   10
      Profiles(0).PrinterMargins_set=   -1  'True
      Profiles(0).PrinterPaperUserSize_set=   -1  'True
   End
End
Attribute VB_Name = "frmrpt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False