VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rptTest 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   1365
      Top             =   7380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7245
      Left            =   15
      TabIndex        =   1
      Top             =   30
      Width           =   10065
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   8325
      TabIndex        =   0
      Top             =   7290
      Width           =   1785
   End
End
Attribute VB_Name = "rptTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oApp As New CRAXDRT.Application
Dim oRpt As New CRAXDRT.Report
Dim a As New CRAXDRT.ParameterValue
Dim ab As CrystalReport

Private Sub Command1_Click()
'  Set oRpt = oApp.OpenReport(App.Path & "\Report\test.rpt", 1)
'  MsgBox oRpt.GetNextRows(0, 1)
'  CRViewer1.ReportSource = oRpt
'  CRViewer1.ViewReport

  CrystalReport2.ReportFileName = App.Path & "\Report\test.rpt"
  CrystalParamater CrystalReport2, 0, "nama", "Madeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee gantengggggggggggggggggggggggggggggggg"
  CrystalParamater CrystalReport2, 1, "alamat", "Jl Pudak Gg 1 No 1"
  CrystalReport2.Action = 0
  
End Sub

Private Function CrystalParamater(Crys As CrystalReport, nIndex As Integer, cNamaField As String, cValue As String) As String
  Crys.ParameterFields(nIndex) = cNamaField & ";" & cValue & ";True"
End Function
