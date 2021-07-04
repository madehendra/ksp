VERSION 5.00
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPenghapusanBungaTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penghapusan Bunga Tabungan"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6315
   Begin VB.TextBox cText 
      Height          =   1500
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "trPenghapusanBungaTabungan.frx":0000
      Top             =   495
      Width           =   5970
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Proses"
      Height          =   420
      Left            =   135
      TabIndex        =   1
      Top             =   2040
      Width           =   5970
   End
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2820
      _ExtentX        =   4974
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
      Caption         =   "Tgl"
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
Attribute VB_Name = "trPenghapusanBungaTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub cmdOK_Click()
  If MsgBox("Akan melakukan penghapusan bunga tabungan.. Akan dilanjutkan?", vbInformation + vbYesNo) = vbYes Then
    Set dbData = objData.Browse(GetDSN, "MutasiTabungan", , "KodeTransaksi", sisAssign, "03", " and tgl = '" & Format(dTgl.Value, "yyyy-MM-dd") & "'")
    If Not dbData.eof Then
      FrmPB.InitPB dbData.RecordCount
      Do While Not dbData.eof
        FrmPB.RunPB
        objData.Delete GetDSN, "BukuBesar", "Faktur", sisAssign, GetNull(dbData!Faktur)
        dbData.MoveNext
      Loop
      FrmPB.EndPB
      'hapus di mutasitabungan
      objData.Delete GetDSN, "MutasiTabungan", "KodeTransaksi", sisAssign, "03", "and tgl ='" & Format(dTgl.Value, "yyyy-MM-dd") & "'"
      MsgBox "Penghapusan mutasi bunga tabungan selesai..", vbInformation + vbOKOnly
    Else
      MsgBox "Data tidak ada.", vbInformation
    End If
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  CenterForm Me
  TabIndex dTgl, n
  TabIndex cmdOK, n
  GetInfo
End Sub

Private Sub GetInfo()
  cText = "Perhatian," & vbCrLf
  cText = cText & "Modul ini akan menghapus mutasi bunga tabungan yang pernah dilakukan "
  cText = cText & "pada tgl ybs." & vbCrLf
  cText = cText & "Setelah transaksi diproses, data yang sudah dihapus tidak bisa dikembalikan lagi (rollback)" & vbCrLf
  cText = cText & "Proses penghapusan akan mengapus data mutasi tabungan yang ada pada " & vbCrLf
  cText = cText & "Table MUTASITABUNGAN dan BUKUBESAR"
End Sub
