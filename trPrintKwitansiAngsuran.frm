VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form trPrintKwitansiAngsuran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Kwitansi Angsuran"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9930
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   6765
      Left            =   45
      TabIndex        =   0
      Top             =   165
      Width           =   9840
      _cx             =   17357
      _cy             =   11933
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   100
      ZoomMode        =   0
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
End
Attribute VB_Name = "trPrintKwitansiAngsuran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Public noOrder As String
Public nTotal As Double
Public nSubTotal As Double
Public nCash As Double
Public nChange As Double
Public nTax As String
Public nDiscount As String
Public lStatus As Double

Public cAnggota As String
Public cNamaAnggota As String
Public cNoValidasi As String
Public nAngsuranBunga As Double
Public nAngsuranPokok As Double
Public nAngsuranDenda As Double
Public nJumlahAngsuran As Double
Public nBakiDebet As Double
Public cRekening As String
Public cNamaPeminjam As String
Public dTgl As Date


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  CenterForm Me
  DoText
End Sub

Private Sub DoText()
Dim i%
Dim nTahun As Double
Dim nBulan As Double
Dim nHari As Double

    MousePointer = 11
    SetOriginalSettings
    vp.ZoomMode = zmPageWidth
    vp.StartDoc
      With vp
         vp = " "
        .FontSize = 12
        .TextColor = vbBlack
        .Text = vbTab & vbTab & vbTab & vbTab & aCfg(msNama) & vbCrLf
        .FontSize = 10
        .Text = vbTab & vbTab & vbTab & vbTab & aCfg(msAlamat) & vbCrLf
        .Text = vbTab & vbTab & vbTab & vbTab & aCfg(msTelepon) & vbCrLf
        .Text = "" & vbCrLf
        .FontSize = 9
        .FontName = "Tahoma"
        .Text = "No Validasi " & cNoValidasi & vbCrLf
        .Text = "Tgl " & Format(dTgl, "dd MM yyyy") & vbCrLf
        .Text = "" & vbCrLf
        .Text = "Anggota: [" & cAnggota & "]" & cNamaAnggota & vbCrLf
        .Text = "Rekening: [" & cRekening & "]"
        Garis
         .Text = "Angsuran pokok" & vbTab & ":" & Padl(Format(nAngsuranPokok, "###,###,##0.00"), 12, "*") & vbTab & vbTab & "Teller" & vbTab & vbTab & vbTab & "Peminjam" & vbCrLf
         .Text = "Angsuran Bunga" & vbTab & ":" & Padl(Format(nAngsuranBunga, "###,###,##0.00"), 12, "*") & vbCrLf
         .Text = "Denda(jika ada)" & vbTab & ":" & Padl(Format(nAngsuranDenda, "###,###,##0.00"), 12, "*") & vbCrLf
         .Text = "Total angsuran" & vbTab & ":" & Padl(Format(nJumlahAngsuran, "###,###,##0.00"), 12, "*") & vbTab & vbTab & "(" & GetRegistry(reg_FullName) & ")" & vbTab & vbTab & "(" & cNamaPeminjam & ")"
        Garis
        .Text = "Baki Debet" & vbTab & ":" & Padl(Format(nBakiDebet, "###,###,##0.00"), 12, "*")
      End With
    vp.EndDoc
    MousePointer = 0
End Sub

Private Sub Garis()
Dim a As Integer
  
  vp.Text = vbCrLf
  For a = 1 To 32
    vp.Text = "="
  Next a
  vp.Text = vbCrLf
End Sub

Private Sub SetOriginalSettings()
    With vp
        .PaperSize = pprUser
        
        .PaperHeight = 5000
        .PaperWidth = 10000

        .ToolTipText = ""
        ' font
        .FontName = "Arial"
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontSize = 11
        
        ' text
        .TextColor = 0
        .TextAngle = 0
        .TextAlign = taLeftMiddle
        
        'spacing
        .IndentLeft = 0
        .IndentFirst = 0
        .IndentRight = 0

        .MarginBottom = 0
        .MarginFooter = 0
        .MarginLeft = 500
        .MarginHeader = 0
        .MarginRight = 1440
        .MarginTop = 0
        

        'drawing
        .PenColor = vbBlack
        .PenStyle = psSolid
        .PenWidth = 1
        .BrushColor = &H8080FF
        .BrushStyle = bsSolid
    
        ' table
        .TableBorder = tbAll
    
        .X1 = 0
        .Y1 = 0
        .X2 = 0
        .Y2 = 0
    
    End With

End Sub


