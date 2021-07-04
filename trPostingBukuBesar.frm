VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPostingBukuBesar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POSTING BUKU BESAR"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5580
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   2820
      Left            =   -15
      Top             =   0
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   4974
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSADateProject.BiSADate BiSADate1 
         Height          =   330
         Left            =   1530
         TabIndex        =   9
         Top             =   2025
         Width           =   1440
         _ExtentX        =   2540
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
      Begin VB.CheckBox chkBukuBesar 
         Caption         =   "Angsuran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   225
         TabIndex        =   8
         Top             =   2100
         Width           =   1110
      End
      Begin VB.CheckBox chkBukuBesar 
         Caption         =   "Pencairan Pinjaman"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   225
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin BiSATextBoxProject.BiSABrowse cRekKasSaldoAwalSimpanan 
         Height          =   330
         Left            =   495
         TabIndex        =   5
         Top             =   900
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         Button          =   -1  'True
         Caption         =   "Rek Kas "
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
      Begin VB.CheckBox chkBukuBesar 
         Caption         =   "Saldo Awal Simpanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   4
         Top             =   615
         Width           =   1935
      End
      Begin VB.CheckBox chkBukuBesar 
         Caption         =   "Simpanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   3
         Top             =   390
         Width           =   1455
      End
      Begin VB.CheckBox chkBukuBesar 
         Caption         =   "Jurnal Umum"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "*Posting saldo awal pinjaman bisa dilakukan lewat menu saldo awal pinjaman"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   180
         TabIndex        =   7
         Top             =   1290
         Width           =   5220
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   570
      Left            =   -15
      Top             =   2820
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   1005
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdProses 
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "   Proses"
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
         Picture         =   "trPostingBukuBesar.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   4980
         TabIndex        =   2
         Top             =   90
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "trPostingBukuBesar.frx":059A
      End
   End
End
Attribute VB_Name = "trPostingBukuBesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As New CodeSuiteLibrary.data

Private Sub Proses()
  If chkBukuBesar(0).Value = 1 Then
    'POSTING JURNAL UMUM
    PostingJurnalUmum
  End If
  If chkBukuBesar(1).Value = 1 Then
    'POSTING SIMPANAN
    PostingSimpanan
  End If
  If chkBukuBesar(2).Value = 1 Then
    'POSTING SALDO AWAL SIMPANAN
    PostingSaldoAwalSimpanan
  End If
  If chkBukuBesar(3).Value = 1 Then
    'POSTING PENCAIRAN
    PostingPencairan
  End If
  If chkBukuBesar(4).Value = 1 Then
    PostingAngsuran
  End If
End Sub

Private Sub PostingPencairan()
Dim cSQL As String
Dim d As New ADODB.Recordset
Dim da As New ADODB.Recordset

  Set d = obj.Browse(GetDSN, "pencairankredit")
  If Not d.eof Then
    FrmPB.InitPB d.RecordCount
    Do While Not d.eof
      FrmPB.RunPB
      cSQL = "select p.faktur,p.rekening,b.rekening as rekeningjurnal,p.penarikan,b.debet,b.kredit from bukubesar b"
      cSQL = cSQL & " left join pencairankredit p on p.faktur = b.faktur"
      cSQL = cSQL & " Where b.status = 2 and b.faktur = '" & GetNull(d!Faktur) & "'"
      Set da = obj.SQL(GetDSN, cSQL)
      If Not da.eof Then
        Do While Not da.eof
          'Update table pencairan
          If GetNull(da!Debet) > 0 And GetNull(da!Kredit) <= 0 Then
            obj.Edit GetDSN, "pencairankredit", "faktur = '" & GetNull(da!Faktur) & "'", Array("total"), Array(GetNull(da!Debet))
          Else
            UpdatePencairan GetNull(da!Rekening), GetNull(da!Kredit), GetNull(da!rekeningjurnal), GetNull(da!Faktur)
          End If
          da.MoveNext
        Loop
      End If
      d.MoveNext
    Loop
  End If
  FrmPB.EndPB
  d.Close
  
  Set d = obj.Browse(GetDSN, "pencairankredit p", "p.rekening,p.faktur,p.tgl,r.nama,p.penarikan,p.total,p.administrasi,p.materai,p.Provisi,p.notaris,p.biayalain,p.username", "d.Tgl", sisGT, "2006-12-31", , , Array("left join debitur d on d.rekening = p.rekening", "left join registernasabah r on r.kode = d.kode"))
  If Not d.eof Then
    FrmPB.InitPB d.RecordCount
    Do While Not d.eof
      FrmPB.RunPB
      UpdRekPencairanKredit GetNull(d!Rekening), GetNull(d!Faktur), GetNull(d!Tgl), GetNull(d!nama), GetNull(d!penarikan), GetNull(d!Total), GetNull(d!Administrasi), GetNull(d!Materai), GetNull(d!Provisi), GetNull(d!Notaris), GetNull(d!biayalain), GetNull(d!UserName)
      d.MoveNext
    Loop
    FrmPB.EndPB
  End If
  d.Close
End Sub

Private Sub UpdRekPencairanKredit(ByVal Rek As String, ByVal cFaktur As String, ByVal dTgl As Date, ByVal cNamaPemilikRekening As String, ByVal nPlafondCair, ByVal nTotalPencairan, ByVal nAdministrasi, ByVal nMaterai, ByVal nProvisi, ByVal nNotaris, ByVal nBiayaLain, ByVal cUser As String)
Dim par As Single
Dim cRekeningKAS As String
Dim cRekeningAdministrasi As String
Dim cRekeningMaterai As String
Dim cRekeningProvisi As String
Dim cRekeningNotaris As String
Dim cRekeningBiayalain As String
Dim cRekeningKYD As String
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

  Set dbData = objData.Browse(GetDSN, "GolonganKredit", , "Kode", sisAssign, Mid(Rek, 4, 2))
  If Not dbData.eof Then
    cRekeningKYD = GetNull(dbData!Rekening, "")
    cRekeningAdministrasi = GetNull(dbData!rekeningadministrasi, "")
    cRekeningMaterai = GetNull(dbData!rekeningmaterai, "")
    cRekeningProvisi = GetNull(dbData!rekeningprovisi, "")
    cRekeningNotaris = GetNull(dbData!RekeningNotaris, "")
    cRekeningBiayalain = GetNull(dbData!RekeningBiayalainLain)
  End If
  
  par = vbTrigger.msRealisasiKredit
  objData.Delete GetDSN, "BukuBesar", "Status", sisAssign, par, "and Faktur = '" & cFaktur & "'"
  cRekeningKAS = GetKasTeller(cUser)
    UpdKodeTr objData, msRealisasiKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningKYD, "Pencairan Kredit an. " & cNamaPemilikRekening, nTotalPencairan, 0, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningKAS, "Pencairan Kredit an. " & cNamaPemilikRekening, 0, nPlafondCair, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningAdministrasi, "Adm. pencairan Kredit an. " & cNamaPemilikRekening, 0, nAdministrasi, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningMaterai, "Materai Pencairan Kredit an. " & cNamaPemilikRekening, 0, nMaterai, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningProvisi, "Provisi pencairan Kredit an. " & cNamaPemilikRekening, 0, nProvisi, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningNotaris, "Notaris Pencairan Kredit an. " & cNamaPemilikRekening, 0, nNotaris, "K", SNow
      UpdKodeTr objData, msRealisasiKredit, aCfg(msKodeCabang), cFaktur, dTgl, cRekeningBiayalain, "Biaya Lain Pencairan Kredit an. " & cNamaPemilikRekening, 0, nBiayaLain, "K", SNow
End Sub

Private Sub UpdatePencairan(ByVal Rek As String, ByVal nValue As Double, ByVal nRekJurnal As String, ByVal Faktur As String)
Dim d As New ADODB.Recordset

  Set d = obj.Browse(GetDSN, "golongankredit", "rekeningadministrasi,rekeningprovisi,rekeningmaterai", "kode", sisAssign, Mid(Rek, 4, 2))
  If Not d.eof Then
    Do While Not d.eof
      Select Case nRekJurnal
        Case GetNull(d!rekeningadministrasi)
          obj.Edit GetDSN, "pencairankredit", "faktur = '" & Faktur & "'", Array("administrasi"), Array(nValue)
        Case GetNull(d!rekeningprovisi)
          obj.Edit GetDSN, "pencairankredit", "faktur = '" & Faktur & "'", Array("provisi"), Array(nValue)
        Case GetNull(d!rekeningmaterai)
          obj.Edit GetDSN, "pencairankredit", "faktur = '" & Faktur & "'", Array("materai"), Array(nValue)
      End Select
      d.MoveNext
    Loop
  End If
End Sub

Private Sub PostingSimpanan()
Dim dba As New ADODB.Recordset
  
  'update rekening simpanan yg hilang terlebih dahulu
  Set dba = obj.Browse(GetDSN, "mutasitabungan")
  If Not dba.eof Then
    FrmPB.InitPB dba.RecordCount
    Do While Not dba.eof
      FrmPB.RunPB
      obj.Update GetDSN, "mutasitabungan", "rekening = '" & GetNull(dba!Rekening) & "' and faktur = '" & GetNull(dba!Faktur) & "'", Array("rekeningjurnal"), Array(GetRekKodeTransaksi(GetNull(dba!KodeTransaksi), GetNull(dba!UserName)))
      dba.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  obj.Delete GetDSN, "bukubesar", "left(faktur, 3)", sisDifference, "SAT", " and status = " & msTabungan
  Set dba = obj.Browse(GetDSN, "mutasitabungan", , "left(faktur, 3)", sisDifference, "SAT")
  If Not dba.eof Then
    FrmPB.InitPB dba.RecordCount
    Do While Not dba.eof
      FrmPB.RunPB
      'UpdMutasiTabungan obj, GetNull(dba!KodeTransaksi), GetNull(dba!faktur), GetNull(dba!Tgl), GetNull(dba!Rekening), GetNull(dba!jumlah), True, GetNull(dba!Keterangan), , GetNull(dba!DK), GetRekKodeTransaksi(GetNull(dba!KodeTransaksi), GetNull(dba!UserName))
      UpdRekTabungan obj, GetNull(dba!Faktur), False
      dba.MoveNext
    Loop
  End If
  FrmPB.EndPB
  
  dba.Close
End Sub

Private Function GetRekKodeTransaksi(KodeTransaksi As String, ByVal user As String) As String
Dim cRekeningJurnal As String
Dim db As New ADODB.Recordset

  cRekeningJurnal = ""
  Set db = obj.Browse(GetDSN, "kodetransaksi", , "kode", sisAssign, KodeTransaksi)
  If Not db.eof Then
    If GetNull(db!Kas) = "K" Then
      cRekeningJurnal = GetKasTeller(user)
    Else
      cRekeningJurnal = GetNull(db!Rekening)
    End If
  End If
  GetRekKodeTransaksi = cRekeningJurnal
End Function

Private Sub PostingSaldoAwalSimpanan()
Dim db As New ADODB.Recordset
  
  
  obj.SQL GetDSN, "delete from bukubesar where left(faktur,3) = 'SAT'"
  Set db = obj.Browse(GetDSN, "mutasitabungan", , "faktur", sisPrefix, "SAT")
  If Not db.eof Then
    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
      FrmPB.RunPB
      UpdKodeTr obj, msTabungan, aCfg(msKodeCabang), GetNull(db!Faktur), GetNull(db!Tgl), cRekKasSaldoAwalSimpanan.Text, "SALDO AWAL TABUNGAN", GetNull(db!Jumlah), 0, "K", SNow
        UpdKodeTr obj, msTabungan, aCfg(msKodeCabang), GetNull(db!Faktur), GetNull(db!Tgl), GetRekSimpanan(db!Rekening), "SALDO AWAL TABUNGAN", 0, GetNull(db!Jumlah), "K", SNow
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
End Sub

Private Function GetRekSimpanan(ByVal Rek As String) As String
Dim deb As New ADODB.Recordset

  GetRekSimpanan = ""
  Set deb = obj.Browse(GetDSN, "golongantabungan", , "Kode", sisAssign, Mid(Rek, 4, 2))
  If Not deb.eof Then
    GetRekSimpanan = GetNull(deb!Rekening, "")
  End If
End Function

Private Sub PostingJurnalUmum()
Dim db As New ADODB.Recordset

  obj.SQL GetDSN, "delete from bukubesar where status = " & vbTrigger.msJurnalLain
  Set db = obj.Browse(GetDSN, "jurnal")
  If Not db.eof Then
    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
      FrmPB.RunPB
      UpdKodeTr obj, msJurnalLain, aCfg(msKodeCabang), GetNull(db!Faktur), GetNull(db!Tgl), GetNull(db!Rekening), GetNull(db!Keterangan), GetNull(db!Debet), GetNull(db!Kredit)
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdProses_Click()
  If isValidProses Then
    Proses
  End If
End Sub

Private Function PostingAngsuran()
Dim dba As New ADODB.Recordset
Dim cGol As String
Dim cRekeningPokok As String
Dim cRekeningBunga As String
Dim cRekeningDenda As String
Dim cNamaDebitur As String

Set dba = obj.Browse(GetDSN, "angsuran", , "tgl", sisAssign, Format(BiSADate1.Value, "yyyy-MM-dd"))
If Not dba.eof Then
  FrmPB.InitPB dba.RecordCount
  Do While Not dba.eof
    FrmPB.RunPB
    cGol = Mid(GetNull(dba!Rekening), 4, 2)
    
    GetRekGolonganKredit cGol, cRekeningBunga, cRekeningPokok, cRekeningDenda
    cNamaDebitur = GetNamaDebitur(GetNull(dba!Rekening))
    
    obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msAngsuranKredit, "And Faktur='" & GetNull(dba!Faktur) & "'"
    UpdKodeTr obj, msAngsuranKredit, left(GetNull(dba!Rekening), 2), GetNull(dba!Faktur), Format(GetNull(dba!Tgl), "yyyy-MM-dd"), GetKasTeller(GetNull(dba!UserName)), "Angsuran Kredit an. " & cNamaDebitur, GetNull(dba!Total), 0, "K"
      UpdKodeTr obj, msAngsuranKredit, left(GetNull(dba!Rekening), 2), GetNull(dba!Faktur), Format(GetNull(dba!Tgl), "yyyy-MM-dd"), cRekeningPokok, "Angsuran Pokok Kredit an. " & cNamaDebitur, 0, GetNull(dba!pokok), "K"
      UpdKodeTr obj, msAngsuranKredit, left(GetNull(dba!Rekening), 2), GetNull(dba!Faktur), Format(GetNull(dba!Tgl), "yyyy-MM-dd"), cRekeningBunga, "Angsuran Bunga Kredit an. " & cNamaDebitur, 0, GetNull(dba!bunga), "K"
      UpdKodeTr obj, msAngsuranKredit, left(GetNull(dba!Rekening), 2), GetNull(dba!Faktur), Format(GetNull(dba!Tgl), "yyyy-MM-dd"), cRekeningDenda, "Denda Angsuran Kredit an. " & cNamaDebitur, 0, GetNull(dba!denda), "K"
    dba.MoveNext
  Loop
  FrmPB.EndPB
End If
End Function

Private Sub GetRekGolonganKredit(ByVal cGol As String, ByRef cRekeningBunga As String, ByRef cRekeningPokok As String, ByRef cRekeningDenda As String)
Dim a As New ADODB.Recordset
  
  Set a = obj.Browse(GetDSN, "golongankredit", , "kode", sisAssign, cGol)
  If Not a.eof Then
    Do While Not a.eof
      cRekeningBunga = GetNull(a!rekeningangsuranbunga)
      cRekeningPokok = GetNull(a!Rekening)
      cRekeningDenda = GetNull(a!rekeningdenda)
      a.MoveNext
    Loop
  End If
End Sub

Private Function GetNamaDebitur(ByVal Rek As String) As String
Dim d As New ADODB.Recordset

  GetNamaDebitur = ""
  Set d = obj.Browse(GetDSN, "registernasabah r", , "d.rekening", sisAssign, Rek, , , Array("left join debitur d on d.kode = r.kode"))
  If Not d.eof Then
    Do While Not d.eof
      GetNamaDebitur = GetNull(d!nama)
      d.MoveNext
    Loop
  End If
End Function

Private Function isValidProses() As Boolean
isValidProses = True
  
  If chkBukuBesar(2).Value = 1 Then
    If Trim(cRekKasSaldoAwalSimpanan.Text) = "" Then
      isValidProses = False
      MsgBox "Rek Saldo awal simpanan tidak boleh kosong"
      cRekKasSaldoAwalSimpanan.SetFocus
    End If
  End If
End Function

Private Sub cRekKasSaldoAwalSimpanan_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = obj.Browse(GetDSN, "rekening", "Kode,Keterangan", "Kode", sisPrefix, "1.")
  If Not db.eof Then
    cRekKasSaldoAwalSimpanan.Text = cRekKasSaldoAwalSimpanan.Browse(db)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  BiSADate1.Value = Date
  TabIndex chkBukuBesar(0), n
  TabIndex chkBukuBesar(1), n
  TabIndex chkBukuBesar(2), n
  TabIndex chkBukuBesar(3), n
  TabIndex chkBukuBesar(4), n
  TabIndex BiSADate1, n
  
  TabIndex cmdProses, n
  TabIndex cmdKeluar, n
End Sub
