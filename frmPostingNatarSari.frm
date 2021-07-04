VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmPostingNatarSari 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posting Natar Sari"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   420
      Left            =   255
      TabIndex        =   2
      Top             =   1185
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   741
      Caption         =   "Kredit"
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
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   450
      Left            =   240
      TabIndex        =   1
      Top             =   690
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   794
      Caption         =   "Tabungan"
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
   Begin BiSAButtonProject.BiSAButton cmdCreateKode 
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   767
      Caption         =   "Create Kode"
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
Attribute VB_Name = "frmPostingNatarSari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.data

Private Sub BiSAButton1_Click()
Dim cRekening As String
Dim vaField, vaValue
  Set dbData = objData.Browse(GetDSN, "sheet_tabungan")
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
      FrmPB.RunPB
      cRekening = GetNull(dbData!Rekening)
      
      'Hapus dulu di Mutasitabungan
      objData.Delete GetDSN, "MutasiTabungan", "Faktur", sisAssign, "SAT0000001", "And Rekening='" & cRekening & "'"
      
      'Simpan di MutasiTabungan
      vaField = Array("Faktur", "Tgl", "Rekening", "Jumlah", "DK", "Keterangan", "UserName", "DateTime")
      vaValue = Array("SAT0000001", "2009-09-08", cRekening, GetNull(dbData!saldo), "K", "SALDO AWAL TABUNGAN", cusername, SNow)
      objData.Add GetDSN, "MutasiTabungan", vaField, vaValue
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton2_Click()
Dim n As Integer
Dim vaField
Dim vaValue
Dim cGolonganKredit As String
Dim cFaktur As String
Dim dTgl As Date
Dim cRekeningPokok As String
Dim cRekeningBunga As String
Dim cRekeningDenda As String
Dim nSaldoAwalKredit As Double
Dim db As New ADODB.Recordset

  cFaktur = "SAK"
  dTgl = "2006-12-31"
  objData.Delete GetDSN, "Angsuran", "Faktur", sisAssign, cFaktur
  objData.Delete GetDSN, "saldoawalkredit", "Faktur", sisAssign, cFaktur
  nSaldoAwalKredit = 0
  Set dbData = objData.Browse(GetDSN, "sheet_kredit")
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
        FrmPB.RunPB
        vaField = Array("faktur", "rekening", "pokok", "bunga", "denda", "username", "datetime")
        vaValue = Array(cFaktur, GetNull(dbData!Rekening), GetNull(dbData!plafond - dbData!bakidebet), 0, 0, GetRegistry(reg_UserName), SNow)
        objData.Update GetDSN, "saldoawalkredit", "rekening = '" & GetNull(dbData!Rekening) & "'", vaField, vaValue
        objData.Edit GetDSN, "debitur", "rekening = '" & GetNull(dbData!Rekening) & "'", Array("simpananwajib", "plafond"), Array(GetNull(dbData!wajib), GetNull(dbData!plafond))
        cGolonganKredit = Mid(GetNull(dbData!Rekening), 4, 2)
        Set db = objData.Browse(GetDSN, "GolonganKredit", , "Kode", sisAssign, cGolonganKredit)
        If Not db.eof Then
          cRekeningPokok = GetNull(db!RekeningAngsuranPokok, "")
          cRekeningBunga = GetNull(db!rekeningangsuranbunga, "")
          cRekeningDenda = GetNull(db!rekeningdenda, "")
        End If
        vaField = Array("Faktur", "Tgl", "Rekening", "Pokok", "Bunga", "Denda", _
                        "Total", "DateTime", "UserName")
        vaValue = Array(cFaktur, "2009-09-08", GetNull(dbData!Rekening), GetNull(dbData!plafond - dbData!bakidebet), 0, 0, _
                       GetNull(dbData!plafond - dbData!bakidebet), SNow, GetRegistry(reg_UserName))
        objData.Add GetDSN, "Angsuran", vaField, vaValue
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  MsgBox "Data telah disimpan", , "Berhasil"
End Sub

Private Sub cmdCreateKode_Click()
  UpdateKredit
End Sub

Private Function GetKode(ByVal cDepan, ByVal cTengah, ByVal cAkhir) As String
  'sample
  '01.K1.000380.01
  GetKode = ""
  GetKode = "01." & cDepan & "." & Padl(cTengah, 6, "0") & "." & Padl(cAkhir, 2, "0")
End Function

Private Sub UpdateKredit()
 Set dbData = objData.Browse(GetDSN, "sheet_tabungan")
 If Not dbData.eof Then
  FrmPB.InitPB dbData.RecordCount
  Do While Not dbData.eof
    FrmPB.RunPB
    objData.Update GetDSN, "sheet_tabungan", "id = " & GetNull(dbData!id), Array("rekening"), Array(GetKode("T4", GetNull(dbData!Kode), GetNull(dbData!qty)))
    dbData.MoveNext
  Loop
 End If
 FrmPB.EndPB
End Sub
