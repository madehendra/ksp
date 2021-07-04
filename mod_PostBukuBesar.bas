Attribute VB_Name = "mod_PostBukuBesar"
Sub PostingJurnalUmum(ByVal obj As BiSAMyDLL.data)
Dim cCabang As String
Dim db As New ADODB.Recordset

  obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, msJurnalLain
  Set db = obj.Browse(GetDSN, "Jurnal j")
  If db.RecordCount > 0 Then
    FrmPB.InitPB db.RecordCount
    db.MoveFirst
    Do While Not db.eof
    FrmPB.RunPB
      cCabang = Mid(GetNull(db!Faktur), 4, 2)
      UpdKodeTr obj, msJurnalLain, cCabang, GetNull(db!Faktur), GetNull(db!Tgl), GetNull(db!Rekening), GetNull(db!Keterangan), GetNull(db!Debet), GetNull(db!Kredit)
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub
Sub PostingMutasiDeposito(ByVal obj As BiSAMyDLL.data)
Dim db As New ADODB.Recordset
Dim cabang As String

'  trPembukaan = 1 'D
'  trPencairanPokok = 2 'K
'  trPencairanBunga = 3 'K
'  trPinalti = 4 'D
'  trMaterai = 5 'D

  'PEMBUKAAN DEPOSITO
  Set db = obj.Browse(GetDSN, "mutasideposito m", "m.KodeMutasi,m.Faktur,m.Rekening,m.Tgl,r.Nama,d.NominalDeposito,m.Jumlah,m.DateTime,m.UserName,g.*", , , , , , Array("Left Join Deposito d on d.Rekening = m.Rekening", "Left Join RegisterNasabah r on r.Kode = d.Kode", "Left Join GolonganDeposito g on g.Kode = d.GolonganDeposito"))
  If Not db.eof Then
  
'    cRekeningDeposito = GetNull(dbData!RekeningAkuntansi, "")
'    cRekeningJT = GetNull(dbData!RekeningJatuhtempo, "")
'    cRekeningFinalty = GetNull(dbData!RekeningFinalty, "")
'    cRekeningTitipanBunga = GetNull(dbData!CadanganBunga, "")
'    cRekneningPajakBunga = GetNull(dbData!RekeningPajakBunga, "")
'    cRekeningMaterai = GetNull(dbData!RekeningMaterai, "")
  
    cabang = left(db!Rekening, 2)
    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
    obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, msDeposito, " and Faktur='" & GetNull(db!Faktur) & "'"
      FrmPB.RunPB
        If GetNull(db!KodeMutasi) = trDeposito.trPembukaan Then
          UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetKasTeller(obj, GetNull(db!UserName)), "Pembukaan Deposito a.n " & GetNull(db!Nama), GetNull(db!Jumlah), 0, , GetNull(db!DateTime)
              UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetNull(db!RekeningAkuntansi), "Pembukaan Deposito a.n " & GetNull(db!Nama), 0, GetNull(db!Jumlah), , GetNull(db!DateTime)
        ElseIf GetNull(db!KodeMutasi) = trDeposito.trPencairanPokok Then
          UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetNull(db!RekeningAkuntansi), "Pencairan Pokok Deposito a.n " & GetNull(db!Nama), GetNull(db!Jumlah), 0, "K", GetNull(db!DateTime)
            UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetKasTeller(objData, GetNull(db!UserName)), "Pencairan Pokok Deposito a.n " & GetNull(db!Nama), 0, GetNull(db!Jumlah) - GetFinalty(obj, GetNull(db!Faktur)) - GetMaterai(obj, GetNull(db!Faktur)), "K", GetNull(db!DateTime)
            UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetNull(dbData!RekeningFinalty, ""), "Finalty Pencairan Pokok Deposito a.n " & GetNull(db!Nama), 0, GetFinalty(obj, GetNull(db!Faktur)), "K", GetNull(db!DateTime)
            UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetNull(dbData!RekeningMaterai, ""), "Materai Pencairan Pokok Deposito a.n " & GetNull(db!Nama), 0, GetMaterai(obj, GetNull(db!Faktur)), "K", GetNull(db!DateTime)
        End If
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  'PECAIRAN BUNGA
  Set db = obj.Browse(GetDSN, "MutasiBungaDeposito b", "b.*,r.Nama,g.CadanganBunga,g.RekeningPajakBunga", , , , , , Array("Left Join Deposito d on d.Rekening = b.Rekening", "Left Join RegisterNasabah r on r.Kode = d.Kode", "Left Join GolonganDeposito g on g.Kode = d.GolonganDeposito"))
  If Not db.eof Then
    cabang = left(db!Rekening, 2)
    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
      obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, msDeposito, " and Faktur='" & GetNull(db!Faktur) & "'"
      FrmPB.RunPB
        UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetNull(db!CadanganBunga), "Pencairan Bunga Deposito a.n " & GetNull(db!Nama), GetNull(db!Jumlah), 0, , GetNull(db!DateTime)
          UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetKasTeller(obj, GetNull(db!UserName)), "Pencairan Bunga Deposito a.n " & GetNull(db!Nama), 0, GetNull(db!Jumlah) - GetNull(db!Pajak), , GetNull(db!DateTime)
          UpdKodeTr obj, msDeposito, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetNull(db!RekeningPajakBunga), "Pajak Bunga Deposito a.n " & GetNull(db!Nama), 0, GetNull(db!Pajak), , GetNull(db!DateTime)
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Function GetFinalty(ByVal obj As BiSAMyDLL.data, ByVal Faktur As String) As Double
Dim db As New ADODB.Recordset
  
  GetFinalty = 0
  Set db = obj.Browse(GetDSN, "MutasiDeposito", , "Faktur", sisAssign, Faktur, " and KodeMutasi='" & trDeposito.trPinalti & "'")
  If Not db.eof Then
    GetFinalty = GetNull(db!Jumlah)
  End If
End Function

Private Function GetMaterai(ByVal obj As BiSAMyDLL.data, ByVal Faktur As String) As Double
Dim db As New ADODB.Recordset
  
  GetMaterai = 0
  Set db = obj.Browse(GetDSN, "MutasiDeposito", , "Faktur", sisAssign, Faktur, " and KodeMutasi='" & trDeposito.trMaterai & "'")
  If Not db.eof Then
    GetMaterai = GetNull(db!Jumlah)
  End If
End Function

Sub PostingPencairanKredit(ByVal obj As BiSAMyDLL.data)
Dim db As New ADODB.Recordset
Dim par As Single
Dim cRekeningKAS As String
Dim cRekeningAdministrasi As String
Dim cRekeningMaterai As String
Dim cRekeningProvisi As String
Dim cRekeningNotaris As String
Dim cRekeningBiayalain As String
Dim cRekeningKYD As String
Dim cabang As String
Dim Faktur As String
Dim Tgl As Date
  
  Set db = obj.Browse(GetDSN, "PencairanKredit p", "p.Rekening,p.UserName,p.Faktur,p.Tgl,p.DateTime,p.Penarikan,g.Rekening as RekeningKredit,g.RekeningAdministrasi as RekeningAdministrasiKredit,g.RekeningMaterai,g.RekeningProvisi,g.RekeningNotaris,g.RekeningBiayalainLain,r.Nama,d.Plafond,d.Administrasi,d.Materai,d.Provisi,d.Notaris,d.BiayaLainLain", , , , , , Array("Left Join Debitur d on d.Rekening=p.Rekening", "Left Join GolonganKredit g on g.Kode=d.GolonganKredit", "Left join RegisterNasabah r on r.Kode = d.Kode"))
  If Not db.eof Then
  
    cabang = left(db!Rekening, 2)
    cRekeningKAS = GetKasTeller(obj, GetNull(db!UserName, ""))
    par = vbTrigger.msRealisasiKredit
    obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, par
    Faktur = GetNull(db!Faktur)
    Tgl = GetNull(db!Tgl)
    
    'informasi golongan kredit
    cRekeningKYD = GetNull(db!RekeningKredit, "")
    cRekeningAdministrasi = GetNull(db!RekeningAdministrasiKredit, "")
    cRekeningMaterai = GetNull(db!RekeningMaterai, "")
    cRekeningProvisi = GetNull(db!RekeningProvisi, "")
    cRekeningNotaris = GetNull(db!RekeningNotaris, "")
    cRekeningBiayalain = GetNull(db!RekeningBiayalainLain)

    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
      FrmPB.RunPB
      UpdKodeTr obj, msRealisasiKredit, cabang, Faktur, Tgl, cRekeningKYD, "Pencairan Kredit an. " & GetNull(db!Nama, ""), GetNull(db!Plafond), 0, "K", GetNull(db!DateTime)
        UpdKodeTr obj, msRealisasiKredit, cabang, Faktur, Tgl, cRekeningKAS, "Pencairan Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!Penarikan), "K", GetNull(db!DateTime)
        UpdKodeTr obj, msRealisasiKredit, cabang, Faktur, Tgl, cRekeningAdministrasi, "Adm. pencairan Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!Administrasi) / 100 * GetNull(db!Plafond), "K", GetNull(db!DateTime)
        UpdKodeTr obj, msRealisasiKredit, cabang, Faktur, Tgl, cRekeningMaterai, "Materai Pencairan Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!Materai), "K", GetNull(db!DateTime)
        UpdKodeTr obj, msRealisasiKredit, cabang, Faktur, Tgl, cRekeningProvisi, "Provisi pencairan Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!Provisi) / 100 * GetNull(db!Plafond), "K", GetNull(db!DateTime)
        UpdKodeTr obj, msRealisasiKredit, cabang, Faktur, Tgl, cRekeningNotaris, "Notaris Pencairan Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!Notaris), "K", GetNull(db!DateTime)
        UpdKodeTr obj, msRealisasiKredit, cabang, Faktur, Tgl, cRekeningBiayalain, "Biaya Lain Pencairan Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!BiayaLainLain), "K", GetNull(db!DateTime)
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Sub PostingAngsuranKredit(ByVal obj As BiSAMyDLL.data)
Dim cRekeningPokok As String
Dim cRekeningBunga As String
Dim cRekeningDenda As String
Dim db As New ADODB.Recordset
Dim cabang As String

  Set db = obj.Browse(GetDSN, "Angsuran a", "a.Rekening,a.Faktur,a.Tgl,a.Pokok,a.Bunga,a.Denda,a.Total,a.UserName,r.Nama,g.RekeningDenda,g.RekeningAngsuranPokok,g.RekeningAngsuranBunga", , , , , "a.Tgl,a.Rekening", Array("Left Join Debitur d on d.Rekening = a.Rekening", "Left Join GolonganKredit g on g.Kode = d.GolonganKredit", "Left Join RegisterNasabah r on r.Kode = d.Kode"))
  If Not db.eof Then
    FrmPB.InitPB db.RecordCount
    cabang = left(db!Rekening, 2)
    obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msAngsuranKredit
    Do While Not db.eof
      FrmPB.RunPB
      UpdKodeTr obj, msAngsuranKredit, cabang, GetNull(db!Faktur, ""), GetNull(db!Tgl), GetKasTeller(obj, GetNull(db!UserName)), "Angsuran Kredit an. " & GetNull(db!Nama, ""), GetNull(db!Total), 0, "K"
        UpdKodeTr obj, msAngsuranKredit, cabang, GetNull(db!Faktur, ""), GetNull(db!Tgl), GetNull(db!RekeningAngsuranPokok), "Angsuran Pokok Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!Pokok), "K"
        UpdKodeTr obj, msAngsuranKredit, cabang, GetNull(db!Faktur, ""), GetNull(db!Tgl), GetNull(db!RekeningAngsuranBunga), "Angsuran Bunga Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!Bunga), "K"
        UpdKodeTr obj, msAngsuranKredit, cabang, GetNull(db!Faktur, ""), GetNull(db!Tgl), GetNull(db!RekeningDenda), "Denda Angsuran Kredit an. " & GetNull(db!Nama, ""), 0, GetNull(db!Denda), "K"
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Sub PostingDiscountAngsuranHarian(ByVal obj As BiSAMyDLL.data)
Dim db As New ADODB.Recordset
Dim cabang As String
  
  Set db = obj.Browse(GetDSN, "PotonganAngsuran p", "p.faktur,p.Rekening,p.Tgl,p.JumlahPotongan,p.UserName,r.Nama", , , , , , Array("Left Join Debitur d on d.Rekening = p.Rekening", "Left Join RegisterNasabah r on r.Kode = d.Kode"))
  If Not db.eof Then
    cabang = left(db!Rekening, 2)
    FrmPB.InitPB db.RecordCount
    Do While Not db.eof
      FrmPB.RunPB
      obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msAngsuranKredit, " and Faktur='" & GetNull(db!Faktur) & "'"
      UpdKodeTr obj, msAngsuranKredit, cabang, GetNull(db!Faktur), GetNull(db!Tgl), GetKasTeller(obj, GetNull(db!UserName)), "Potongan Kredit Harian an. " & GetNull(db!Nama), GetNull(db!JumlahPotongan), 0, "K"
        UpdKodeTr obj, msAngsuranKredit, cabang, GetNull(db!Faktur), GetNull(db!Tgl), aCfg(msDiscountAngsuran), "Potongan Kredit Harian an. " & GetNull(db!Nama), 0, GetNull(db!JumlahPotongan), "K"
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub


Function PostingMutasiTabungan(ByVal obj As BiSAMyDLL.data)

Dim cDebet As String, cKredit As String
Dim cNamaKodeTransaksi As String
Dim cAtasNama As String
Dim dTgl As Date, cCabang As String
Dim cKas As String, cDK As String
Dim cKeterangan As String
Dim nJumlahDebet As Double, nJumlahKredit As Double
Dim vaJoint
Dim cSQL As String
Dim db As New ADODB.Recordset


  obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msTabungan
  cSQL = " "
  cSQL = cSQL & " Select m.*, t.Rekening as RekeningTabungan,t.Kode ,"
  cSQL = cSQL & " k.Kas,m.RekeningJurnal as RekeningKodeTransaksi,k.Keterangan as KeteranganKodeTransaksi,"
  cSQL = cSQL & " g.Rekening as RekeningPerkiraanTabungan,r.Nama as NamaNasabah,g.RekeningBunga"
  cSQL = cSQL & " From MutasiTabungan m"
  cSQL = cSQL & " Left Join Tabungan t on m.Rekening = t.Rekening"
  cSQL = cSQL & " Left Join GolonganTabungan g on g.Kode = t.GolonganTabungan"
  cSQL = cSQL & " Left Join KodeTransaksi k on k.Kode = m.KodeTransaksi"
  cSQL = cSQL & " Left Join RegisterNasabah r on r.Kode = t.Kode"
  
  Set db = obj.SQL(GetDSN, cSQL)
  FrmPB.InitPB db.RecordCount
  Do While Not db.eof
    FrmPB.RunPB
    cCabang = left(db!Rekening, 2)
    dTgl = db!Tgl
    cDK = db!DK
    cKas = GetNull(db!Kas, "")
    nJumlahDebet = db!Jumlah
    nJumlahKredit = db!Jumlah
    cAtasNama = GetNull(db!NamaNasabah, "")
    cNamaKodeTransaksi = GetNull(db!KeteranganKodeTransaksi, "")
    If cDK = "D" Then
      cDebet = GetNull(db!RekeningPerkiraanTabungan, "")
      cKredit = GetNull(db!RekeningKodeTransaksi)
    Else
      cDebet = GetNull(db!RekeningKodeTransaksi, "")
      cKredit = GetNull(db!RekeningPerkiraanTabungan, "")
    End If
    
    If db!KodeTransaksi = aCfg(msKodeBagiHasil) And Trim(db!RekeningBunga) <> "" Then
      cDebet = db!RekeningBunga
    End If
    
    UpdKodeTr obj, msTabungan, cCabang, GetNull(db!Faktur), _
              dTgl, cDebet, GetNull(db!Keterangan), nJumlahDebet, , cKas
    UpdKodeTr obj, msTabungan, cCabang, GetNull(db!Faktur), dTgl, cKredit, GetNull(db!Keterangan), , nJumlahKredit, cKas
    
    db.MoveNext
  Loop
  FrmPB.EndPB
End Function
