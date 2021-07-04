VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form trPostingAkhirHariPengendapan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POSTING AKHIR HARI PENGENDAPAN SIMPANAN POKOK DAN WAJIB"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10755
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   6435
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10755
      _cx             =   18971
      _cy             =   11351
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
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&SIMPAN"
         Height          =   375
         Left            =   4455
         TabIndex        =   4
         Top             =   60
         Width           =   1155
      End
      Begin VB.CommandButton cmdPosting 
         Caption         =   "&POSTING"
         Height          =   375
         Left            =   3255
         TabIndex        =   3
         Top             =   60
         Width           =   1170
      End
      Begin vbAcceleratorSGrid6.vbalGrid sGrid 
         Height          =   5745
         Left            =   45
         TabIndex        =   1
         Top             =   615
         Width           =   10665
         _ExtentX        =   18812
         _ExtentY        =   10134
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
      End
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   375
         Left            =   45
         TabIndex        =   2
         Top             =   60
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   661
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
         Caption         =   "TGL POSTING"
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
End
Attribute VB_Name = "trPostingAkhirHariPengendapan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset

Private Sub cmdPosting_Click()
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "postingakhirhari", "tgl", "tgl", sisAssign, Format(dTgl.Value, "yyyy-MM-dd"))
  If Not db.eof Then
    MsgBox "Tgl " & dTgl.Value & " sudah pernah diposting sebelumnya", vbInformation, "Informasi"
  End If
  Screen.MousePointer = vbHourglass
  DoLetPosting
  Screen.MousePointer = vbDefault
End Sub

Sub InitGrid(vbagrid As vbalGrid)
  With vbagrid
    .GridLines = False
    .AlternateRowBackColor = RGB(252, 252, 230)
    .RowMode = True
    .NoVerticalGridLines = True
    .GridLines = True
    .DrawFocusRectangle = False
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .BorderStyle = ecgBorderStyle3dThin
  End With
End Sub

Private Sub BuatKolom()
  With sGrid
    .AddColumn , "REKENING", ecgHdrTextALignCentre, , 100
    .AddColumn , "NAMA", ecgHdrTextALignCentre, , 200
    .AddColumn , "TGL VALUTA", ecgHdrTextALignCentre, , 100, , , , , "dd-MM-yyyy", , CCLSortDate
    .AddColumn , "RANGE ", ecgHdrTextALignCentre, , 100, , , , , "dd-MM-yyyy", , CCLSortDate
    .AddColumn , "SD RANGE ", ecgHdrTextALignCentre, , 100, , , , , "dd-MM-yyyy", , CCLSortDate
    .AddColumn , "JUMLAH", ecgHdrTextALignCentre, , 100, , , , , "###,###,##0.00", , CCLSortNumeric
    .AddColumn , "KODE", ecgHdrTextALignCentre, , 100, False
    .AddColumn , "GOLONGAN TABUNGAN", ecgHdrTextALignCentre, , 100, False
  End With
End Sub

Private Sub DoLetPosting()
Dim db As New ADODB.Recordset
Dim dStart As Date
Dim dEnd As Date
Dim i As Integer

  Set db = objData.Browse(GetDSN, "tabungan t", "t.*,r.Nama as nama", "t.close", sisDifference, "1", " and (t.golongantabungan = 'T1' or t.golongantabungan  = 'T2')", , Array("left join registernasabah r on r.kode = t.kode"))
  If Not db.eof Then
    i = 1
    With sGrid
      .Clear
      Do While Not db.eof
        If PostingAkhirHariPengendapan(objData, GetNull(db!Rekening), dTgl.Value) Then
          dStart = DateAdd("m", -2, dTgl.Value)
          dEnd = DateAdd("m", -1, dTgl.Value)
          'masukkan nilai ke kolom
          .CellDetails i, 1, GetNull(db!Rekening)
          .CellDetails i, 2, GetNull(db!nama)
          .CellDetails i, 3, GetNull(db!Tgl)
          .CellDetails i, 4, dStart
          .CellDetails i, 5, dEnd
          .CellDetails i, 6, GetSaldoTerendah(objData, GetNull(db!Rekening), dStart, dEnd, True), DT_RIGHT
          .CellDetails i, 7, GetNull(db!Kode)
          .CellDetails i, 8, GetNull(db!GolonganTabungan)
          i = i + 1
        End If
        db.MoveNext
      Loop
    End With
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim lRow As Integer

  objData.Update GetDSN, "postingakhirhari", "tgl = '" & Format(dTgl.Value) & "'", Array("tgl", "username", "datetime"), Array(Format(dTgl.Value, "yyyy-MM-dd"), GetRegistry(reg_UserName), SNow)
  With sGrid
    For lRow = 1 To .Rows
      objData.Update GetDSN, "simpananmengendap", "rekening = '" & .CellText(lRow, 1) & "' and tahun = '" & Year(dTgl.Value) & "' and bulan = '" & Month(dTgl.Value) & "'", Array("rekening", "tahun", "bulan", "jumlah", "kode", "golongantabungan"), Array(.CellText(lRow, 1), Year(dTgl.Value), Month(dTgl.Value), .CellText(lRow, 6), .CellText(lRow, 7), .CellText(lRow, 8))
      .RowVisible(lRow) = False
    Next lRow
    .Clear
  End With
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  sGrid.Clear
  BuatKolom
  InitGrid sGrid
  dTgl.Value = Date
  TabIndex dTgl, n
  TabIndex cmdPosting, n
  TabIndex cmdSimpan, n
End Sub

Private Sub KolomKlik(vbagrid As vbalGrid, lCol As Long)
Dim sTag As String
Dim iSortIndex As Long
      
   With vbagrid.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lCol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lCol
      End If
   
      ' Determine which sort order to apply:
      sTag = vbagrid.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      vbagrid.ColumnTag(lCol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = vbagrid.ColumnSortType(lCol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   vbagrid.Sort
   Screen.MousePointer = vbDefault
End Sub

Private Sub sGrid_ColumnClick(ByVal lCol As Long)
  KolomKlik sGrid, lCol
End Sub
