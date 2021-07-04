VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmPosting 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton frmPosting 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   555
      Width           =   1290
      _ExtentX        =   2275
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
Attribute VB_Name = "frmPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objData As New CodeSuiteLibrary.data
Dim db As New ADODB.Recordset

Private Sub frmPosting_Click()
Dim cSQL As String
Dim vaField
Dim vaValue

cSQL = "select deb.*,d.rekening as rek from debitur08 deb"
cSQL = cSQL & " left join debitur d on d.rekening = deb.rekening"
cSQL = cSQL & " Where d.Rekening Is Null"

Set db = objData.SQL(GetDSN, cSQL)
FrmPB.InitPB db.RecordCount
If Not db.eof Then
  Do While Not db.eof
    FrmPB.RunPB
    vaField = Array("Rekening", "Wilayah", "Kode", "GolonganKredit", _
                     "NoSPK", "SukuBunga", "Tgl", _
                     "Plafond", "Lama", "AO", "Administrasi", _
                     "Materai", "JatuhTempo", _
                     "TotalBunga", "NoPengajuan", _
                     "CaraAngsuran", "PeriodeBungaMenurun", _
                     "MinimalPeriode", "Provisi", "Notaris", "BiayaLainLain", _
                     "wajibpokok", "KonpensasiTelat", "DendaTelatBayar", "UserName", "DateTime", "caraperhitungan")
                     
    vaValue = Array(GetNull(db!Rekening), GetNull(db!Wilayah), GetNull(db!Kode), GetNull(db!GolonganKredit), _
                     GetNull(db!NoSPK), GetNull(db!SukuBunga), GetNull(db!Tgl), _
                     GetNull(db!plafond), GetNull(db!Lama), GetNull(db!AO), GetNull(db!Administrasi), _
                     GetNull(db!Materai), GetNull(db!JatuhTempo), _
                     GetNull(db!totalBunga), GetNull(db!NoPengajuan), _
                     GetNull(db!CaraAngsuran), GetNull(db!PeriodeBungaMenurun), _
                     GetNull(db!MinimalPeriode), GetNull(db!Provisi), GetNull(db!Notaris), GetNull(db!BiayaLainLain), _
                     GetNull(db!wajibpokok), GetNull(db!KonpensasiTelat), GetNull(db!DendaTelatBayar), GetNull(db!UserName), GetNull(db!DateTime), GetNull(db!caraperhitungan))
                  
    objData.Add GetDSN, "debitur", vaField, vaValue
    db.MoveNext
  Loop
  FrmPB.EndPB
End If
End Sub

Private Sub SimpanRealisasi(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String, ByVal cWilayah As String, _
                 ByVal cKode As String, ByVal cGolonganKredit As String, _
                 ByVal cNoSPK As String, ByVal nPersBunga As Double, _
                 ByVal dTgl As Date, ByVal nPlafond As Double, ByVal nLama As Integer, _
                 ByVal cAO As String, ByVal nAdministrasi As Double, ByVal nMaterai As Double, _
                 ByVal dJatuhTempo As Date, ByVal cNoPengaJuan As String, _
                 ByVal nTotalBunga As Double, ByVal cCaraPembayaran As String, ByVal nPeriod As Integer, _
                 ByVal nMinPeriode As Integer, ByVal nProv As Double, ByVal nNot As Double, _
                 ByVal nKonpensasi As Integer, ByVal nBiayalain As Double, ByVal nDendaKeterlamabatan As Double, ByVal wajibpokok As Double, ByVal cCaraPerhitungan As String)
Dim vaField
Dim vaValue
Dim n As Single

  vaField = Array("Rekening", "Wilayah", "Kode", "GolonganKredit", _
                  "NoSPK", "SukuBunga", "Tgl", _
                  "Plafond", "Lama", "AO", "Administrasi", _
                  "Materai", "JatuhTempo", _
                  "TotalBunga", "NoPengajuan", _
                  "CaraAngsuran", "PeriodeBungaMenurun", _
                  "MinimalPeriode", "Provisi", "Notaris", "BiayaLainLain", _
                  "wajibpokok", "KonpensasiTelat", "DendaTelatBayar", "UserName", "DateTime", "caraperhitungan")
                  
  vaValue = Array(cRekening, cWilayah, cKode, cGolonganKredit, _
                  cNoSPK, nPersBunga, dTgl, _
                  nPlafond, nLama, cAO, nAdministrasi, _
                  nMaterai, dJatuhTempo, _
                  nTotalBunga, cNoPengaJuan, _
                  cCaraPembayaran, nPeriod, _
                  nMinPeriode, nProv, nNot, nBiayalain, _
                  wajibpokok, nKonpensasi, nDendaKeterlamabatan, cusername, SNow, cCaraPerhitungan)
                  
  obj.Update GetDSN, "Debitur", "Rekening = '" & cRekening & "'", vaField, vaValue
  obj.Edit GetDSN, "PengajuanKredit", "Kode='" & cNoPengaJuan & "'", Array("StatusPengajuan"), Array("1")
End Sub


