VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Begin VB.Form frmTDBR1 
   Caption         =   "Form1"
   ClientHeight    =   645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   645
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows Default
   Begin TrueDBReports60Ctl.TDBReports r1 
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
      Vedit_ShowGrid  =   0   'False
      Vedit_SnapToGrid=   0   'False
      Vedit_GridUnitWidth=   2.822
      Vedit_GridUnitHeight=   2.822
      Vedit_ShowCellExpressions=   -1  'True
      Norm_rect_left  =   0
      Norm_rect_top   =   0
      Norm_rect_right =   0
      Norm_rect_bottom=   0
      Virgin          =   0   'False
      Sections.Count  =   1
      Sections(0).Name=   "SECTION_0"
      Sections(0).Type=   4
      Sections(0).StyleExp=   "'STYLE_0'"
      Sections(0).AutoHeight=   0   'False
      Sections(0).Height=   100
      Sections(0).dtopts=   2
      Styles.Count    =   2
      Styles(0).Name  =   "STYLE_0"
      Styles(0).Font_Name=   "Arial"
      Styles(0).Font_Size=   9.75
      Styles(0).Font_Bold=   -1  'True
      Styles(0).Font_Charset=   0
      Styles(0).TextVAlign=   1
      Styles(0).MarginTop=   4
      Styles(0).MarginBottom=   4
      Styles(1).Name  =   "STYLE_1"
      Styles(1).ParentName=   "STYLE_0"
      Styles(1).Font_Name=   "Arial"
      Styles(1).Font_Size=   9.75
      Styles(1).Font_Charset=   0
      Styles(1).TextVAlign=   1
      Styles(1).MarginTop=   4
      Styles(1).MarginBottom=   4
      Styles(1).fprops=   16777216
      Profiles.Count  =   1
      Profiles(0).Name=   "Profile"
      Profiles(0).Active=   -1  'True
      Profiles(0).PreviewNoMinimize=   -1  'True
      Profiles(0).PreviewNoMaximize=   -1  'True
      Profiles(0).PreviewNoResize=   -1  'True
      Profiles(0).PreviewMaximized=   -1  'True
      Profiles(0).PreviewNoSaveLoad=   -1  'True
      Profiles(0).PrinterMarginLeft=   2
      Profiles(0).PrinterMarginTop=   2
      Profiles(0).PrinterMarginRight=   2
      Profiles(0).PrinterMarginBottom=   2
      Profiles(0).PrinterPaperSize=   256
      Profiles(0).PrinterMargins_set=   -1  'True
      Profiles(0).PrinterPaperSize_set=   -1  'True
      Profiles(0).PrinterPaperUserSize_set=   -1  'True
   End
End
Attribute VB_Name = "frmTDBR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer

Private Sub AddCell(ByVal PointKe As Integer, ByVal xPos As Integer, ByVal yPos As Integer, ByVal cValue As String, Optional ByVal nWidth As Double = 30)
  r1.Sections(0).Cells.Add PointKe
  With r1.Sections(0).Cells(PointKe)
    .Exp = "'" & cValue & "'"
    .Placement = tdbCellPlacementFree
    .WidthInPercent = False
    .Width = nWidth
    .x = xPos
    .Y = yPos
  End With
End Sub

Sub AddPoint(ByVal cValue As String, ByVal xPos As Integer, ByVal yPos As Integer, _
                     ByVal nWidth As Double, ByRef PointKe As Integer)
  
  AddCell PointKe, xPos, yPos, cValue, nWidth
  PointKe = PointKe + 1
End Sub

Sub SetMargin(ByVal nHeight As Double, ByVal nWidth As Double, Optional ByVal lFontBold As Boolean = True)
  ClearingReport
  With r1.Profiles(0)
    '.PrinterMargins_set = True
    '.PrinterMarginTop = nTop
    '.PrinterMarginLeft = nLeft
    
    .PrinterPaperSize_set = True
    .PrinterPaperHeight = nHeight
    .PrinterPaperWidth = nWidth
  End With
  'r1.Sections(0).Height = nHeight
  
  r1.Sections(0).Style.Font_Bold = lFontBold
End Sub

Sub PrintPreview(Optional ByVal lPreview As Boolean = True)
  If lPreview = True Then
    r1.PrintPreview
  Else
    r1.PrintData
  End If
End Sub

Private Sub ClearingReport()
Dim i As Double

  Do While r1.Parameters.Count > 0
    r1.Parameters.Remove 0
  Loop

  Do While r1.Sections(0).Cells.Count > 0
    r1.Sections(0).Cells.Remove 0
  Loop
End Sub
