VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmPostingCikarSedana 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   705
      Left            =   780
      TabIndex        =   1
      Top             =   1110
      Width           =   2730
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   375
      Left            =   525
      TabIndex        =   0
      Top             =   210
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Label1"
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
End
Attribute VB_Name = "frmPostingCikarSedana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub BiSAButton1_Click()
Dim cSQL As String
  
  cSQL = "select d.rekening as rek,deb.rekening from debiturlama d"
  cSQL = cSQL & " left join debitur deb on deb.rekening = d.rekening"
  cSQL = cSQL & " where deb.rekening is null;"
  
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.eof Then
    Do While Not db.eof
      GetInsertDebitur GetNull(db!Rek)
      db.MoveNext
    Loop
  End If
End Sub

Private Sub GetInsertDebitur(ByVal cRek As String)
Dim dbData As New ADODB.Recordset

  Set dbData = objData.Browse(GetDSN, "debiturlama", , "rekening", sisAssign, cRek)
  If Not dbData.eof Then
    objData.Add GetDSN, "debitur", Array("posting", "statuspencairan", "faktur", "rekening", "nopengajuan", "wilayah", "kode", "golongankredit", "nospk", "sukubunga", "tgl", "plafond", "totalbunga", "lama", "ao", "status", "administrasi", "materai", "provisi", "notaris", "biayalainlain", "username", "datetime", "jatuhtempo", "wajibpokok", "statuspengajuan", "caraangsuran", "periodebungamenurun", "minimalperiode", "konpensasitelat", "dendatelatbayar", "caraperhitungan"), Array(dbData!posting, dbData!statuspencairan, dbData!Faktur, dbData!Rekening, dbData!NoPengajuan, dbData!Wilayah, dbData!Kode, dbData!GolonganKredit, dbData!NoSPK, dbData!SukuBunga, dbData!Tgl, dbData!plafond, dbData!totalBunga, dbData!Lama, dbData!AO, dbData!status, dbData!Administrasi, dbData!Materai, dbData!Provisi, dbData!Notaris, dbData!BiayaLainLain, _
    dbData!UserName, dbData!DateTime, dbData!JatuhTempo, dbData!wajibpokok, dbData!statuspengajuan, dbData!CaraAngsuran, dbData!PeriodeBungaMenurun, dbData!MinimalPeriode, dbData!KonpensasiTelat, dbData!DendaTelatBayar, dbData!caraperhitungan)
  End If
End Sub

Private Sub Command1_Click()
  Load trFormModal
  trFormModal.Show
  trFormModal.SetFocus
End Sub
