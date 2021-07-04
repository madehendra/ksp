Attribute VB_Name = "Trigger"
Option Explicit
Dim dbData As New ADODB.Recordset

Public Enum vbTrigger
  msTabungan = 1
  msRealisasiKredit = 2
  msAngsuranKredit = 3
  msTitipanAngsuran = 4
  msKasKeluar = 5
  msKasMasuk = 6
  msDeposito = 7
  msRekeningRRP = 8
  msJurnalLain = 9
  msPenyusutanAktiva = 10
  msAmortisasiProvisi = 11
  msPenambahanPlafond = 12
  msAdministrasiTabungan = 13
End Enum

Public Enum trDeposito
  trPembukaan = 1
  trPencairanPokok = 2
  trPencairanBunga = 3
  trPinalti = 4
  trMaterai = 5
End Enum

Sub UpdKodeTr(ByVal obj As CodeSuiteLibrary.data, ByVal par As vbTrigger, ByVal cCabang As String, ByVal cFaktur As String, _
              ByVal dTgl As Date, ByVal cRekening As String, _
              Optional ByVal cKeterangan As String = "", Optional ByVal nDebet As Double = 0, _
              Optional ByVal nKredit As Double = 0, Optional cKas As String = "K", _
              Optional ByVal cNow As String = "")
              
Dim vaField, vaValue
  cNow = IIf(cNow = "", SNow, cNow)
  If nDebet <> 0 Or nKredit <> 0 Then
    vaField = Array("Cabang", "Faktur", "Tgl", "Rekening", "Keterangan", "Debet", "Kredit", _
                    "Kas", "UserName", "Status", "DateTime")
    vaValue = Array(cCabang, cFaktur, dTgl, cRekening, cKeterangan, nDebet, nKredit, _
                    cKas, cusername, par, cNow)
    obj.Add GetDSN, "BukuBesar", vaField, vaValue
  End If
End Sub

Sub DelKodeTr(ByVal obj As CodeSuiteLibrary.data, ByVal par As vbTrigger, ByVal cCabang As String, ByVal cFaktur As String)
  obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, Trim(Str(par)), " And Faktur = '" & cFaktur & "'"
End Sub

Function UpdRekTabungan(ByVal obj As CodeSuiteLibrary.data, ByVal cFaktur As String, Optional ByVal lDelBukuBesar As Boolean = True)
Dim cDebet As String, cKredit As String
Dim cNamaKodeTransaksi As String
Dim cAtasNama As String
Dim dTgl As Date, cCabang As String
Dim cKas As String, cDK As String
Dim cKeterangan As String
Dim nJumlahDebet As Double, nJumlahKredit As Double
Dim vaJoint
Dim cSQL As String

  If lDelBukuBesar = True Then
    obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msTabungan, " and Faktur = '" & cFaktur & "'"
  End If
  cSQL = " "
  cSQL = cSQL & " Select m.*, t.Rekening as RekeningTabungan,t.Kode ,"
  cSQL = cSQL & " k.Kas,m.RekeningJurnal as RekeningKodeTransaksi,k.Keterangan as KeteranganKodeTransaksi,"
  cSQL = cSQL & " g.Rekening as RekeningPerkiraanTabungan,r.Nama as NamaNasabah,g.RekeningBunga"
  cSQL = cSQL & " From MutasiTabungan m"
  cSQL = cSQL & " Left Join Tabungan t on m.Rekening = t.Rekening"
  cSQL = cSQL & " Left Join GolonganTabungan g on g.Kode = t.GolonganTabungan"
  cSQL = cSQL & " Left Join KodeTransaksi k on k.Kode = m.KodeTransaksi"
  cSQL = cSQL & " Left Join RegisterNasabah r on r.Kode = t.Kode"
  cSQL = cSQL & " Where m.Faktur='" & cFaktur & "'"
  Set dbData = obj.SQL(GetDSN, cSQL)
  Do While Not dbData.eof
    cCabang = left(dbData!Rekening, 2)
    dTgl = dbData!Tgl
    cDK = dbData!DK
    cKas = GetNull(dbData!Kas, "")
    nJumlahDebet = dbData!Jumlah
    nJumlahKredit = dbData!Jumlah
    cAtasNama = GetNull(dbData!NamaNasabah, "")
    cNamaKodeTransaksi = GetNull(dbData!KeteranganKodeTransaksi, "")
    If cDK = "D" Then
      cDebet = GetNull(dbData!RekeningPerkiraanTabungan, "")
      cKredit = GetNull(dbData!RekeningKodeTransaksi)
    Else
      cDebet = GetNull(dbData!RekeningKodeTransaksi, "")
      cKredit = GetNull(dbData!RekeningPerkiraanTabungan, "")
    End If
    
    If dbData!KodeTransaksi = aCfg(mskodebagihasil) And Trim(dbData!Rekeningbunga) <> "" Then
      cDebet = dbData!Rekeningbunga
    End If
    
    UpdKodeTr obj, msTabungan, cCabang, cFaktur, _
              dTgl, cDebet, GetNull(dbData!Keterangan), nJumlahDebet, , cKas
    UpdKodeTr obj, msTabungan, cCabang, cFaktur, dTgl, cKredit, GetNull(dbData!Keterangan), , nJumlahKredit, cKas
    
    dbData.MoveNext
  Loop
End Function

Function UpdRekDeposito(ByVal obj As CodeSuiteLibrary.data, ByVal cFaktur As String)
Dim cNamaKodeMutasi, cAtasNama As String
Dim dTgl As Date
Dim cCabang, cKeterangan As String
Dim cDebet As String, cKredit As String
Dim cKas As String, cDK As String
Dim nJumlahDebet, nJumlahKredit As Double
Dim vaJoint

  
  obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, vbTrigger.msDeposito, " and Faktur = '" & cFaktur & "'"
  vaJoint = Array("Left Join Deposito d on d.Rekening = m.Rekening", _
                  "Left Join GolonganDeposito g on g.Kode = d.GolonganDeposito", _
                  "Left Join KodeMutasiDeposito k on k.KodeMutasi = m.KodeMutasi", _
                  "Left Join RegisterNasabah r on r.Kode = d.Kode")
  Set dbData = obj.Browse(GetDSN, "MutasiDeposito m", "m.*,d.Rekening as RekeningDeposito,k.DK,k.Kas,k.Rekening as RekeningKodeMutasi,k.Keterangan as NamaKodeMutasi,k.Rekening as RekPerDeposito,r.Nama as NamaNasabah", "m.Faktur", sisAssign, cFaktur, , , vaJoint)
  Do While Not dbData.eof
    cCabang = left(dbData!Rekening, 2)
    dTgl = dbData!Tgl
    cDK = dbData!DK
    cKas = dbData!Kas
    nJumlahDebet = dbData!Jumlah
    nJumlahKredit = dbData!Jumlah
    cAtasNama = dbData!NamaNasabah
    cNamaKodeMutasi = dbData!NamaKodeMutasi
    If cDK = "D" Then
      cDebet = dbData!RekPerDeposito
      cKredit = dbData!RekeningKodeMutasi
    Else
      cDebet = dbData!RekeningKodeMutasi
      cKredit = dbData!RekPerDeposito
    End If
    UpdKodeTr obj, msDeposito, cCabang, cFaktur, _
              dTgl, cDebet, cNamaKodeMutasi & " an. " & cAtasNama, nJumlahDebet, , cKas
    UpdKodeTr obj, msDeposito, cCabang, cFaktur, dTgl, cKredit, cNamaKodeMutasi & " an. " & cAtasNama, , nJumlahKredit, cKas
    dbData.MoveNext
  Loop
End Function

Function UpdRekRealisasiPembiayaan(ByVal obj As CodeSuiteLibrary.data, ByVal cFaktur As String, Optional ByVal cNow As String = "")
Dim vaJoint
Dim cField As String
Dim par As Single
Dim nNetto As Double
Dim cRekeningKAS As String
Dim nPlafond As Double

  cNow = IIf(cNow = "", SNow, cNow)
  par = vbTrigger.msRealisasiKredit
  obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, par, "and Faktur = '" & cFaktur & "'"
  
  cField = "d.StatusPencairan,d.CaraPencairan,d.Rekening,d.Tgl,d.CaraPerhitungan,d.PlafondPRK,"
  cField = cField & "d.Plafond,d.Administrasi,d.Materai,d.Notaris,d.Asuransi,d.TotalBunga,d.Provisi,"
  cField = cField & "g.Rekening as RekeningRealisasi,g.RekeningAdministrasi,"
  cField = cField & "g.RekeningMaterai,g.RekeningNotaris,g.RekeningAsuransi,g.RekeningProvisi,"
  cField = cField & "c.Nama as NamaDebitur "
  vaJoint = Array("Left Join GolonganKredit g on d.GolonganKredit = g.Kode", _
                  "Left Join RegisterNasabah c on c.Kode = d.Kode")
  Set dbData = obj.Browse(GetDSN, "Debitur d", cField, "d.Faktur", sisAssign, cFaktur, , , vaJoint)
  If dbData.RecordCount > 0 Then
    If dbData!statuspencairan = "1" Then
      cRekeningKAS = IIf(dbData!CaraPencairan = "1", cKasTeller, aCfg(msKodePemindahBukuan))
      nNetto = GetNull(dbData!plafond) - GetNull(dbData!Administrasi) - GetNull(dbData!Materai)
      nPlafond = dbData!plafond
        UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, dbData!RekeningRealisasi, "Realisasi/Pencairan Kredit an. " & dbData!NamaDebitur, nPlafond, , "K", cNow
          UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, cRekeningKAS, "Realisasi/Pencairan Kredit an. " & dbData!NamaDebitur, , nPlafond, "K", cNow
        UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, cRekeningKAS, "Administrasi Realisasi Kredit an. " & dbData!NamaDebitur, dbData!Administrasi, , "K", cNow
          UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, dbData!rekeningadministrasi, "Administrasi Realisasi Kredit an. " & dbData!NamaDebitur, , dbData!Administrasi, "K", cNow
        UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, cRekeningKAS, "Notaris Realisasi Kredit an. " & dbData!NamaDebitur, dbData!Notaris, , "K", cNow
          UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, dbData!RekeningNotaris, "Notaris Realisasi Kredit an. " & dbData!NamaDebitur, , dbData!Notaris, "K", cNow
        UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, cRekeningKAS, "Materai Realisasi Kredit an. " & dbData!NamaDebitur, dbData!Materai, , "K", cNow
          UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, dbData!rekeningmaterai, "Materai Realisasi Kredit an. " & dbData!NamaDebitur, , dbData!Materai, "K", cNow
        UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, cRekeningKAS, "Asuransi Realisasi Kredit an. " & dbData!NamaDebitur, dbData!Asuransi, , "K", cNow
          UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, dbData!Rekeningasuransi, "Asuransi Realisasi Kredit an. " & dbData!NamaDebitur, , dbData!Asuransi, "K", cNow
        UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, cRekeningKAS, "Provisi Realisasi Kredit an. " & dbData!NamaDebitur, dbData!Provisi, , "K", cNow
          UpdKodeTr obj, msRealisasiKredit, left(dbData!Rekening, 2), cFaktur, dbData!Tgl, dbData!rekeningprovisi, "Provisi Realisasi Kredit an. " & dbData!NamaDebitur, , dbData!Provisi, "K", cNow
    End If
  End If
End Function

Function UpdMutasiTabungan(ByVal obj As CodeSuiteLibrary.data, ByVal cKodeTransaksi As String, _
                           ByVal cFaktur As String, ByVal dTgl As Date, ByVal cRekening As String, _
                           ByVal nTotal As Double, Optional ByVal lWithDeleteExist As Boolean = False, _
                           Optional ByVal cKeterangan As String = "", _
                           Optional ByVal lUpdateToBukuBesar As Boolean = True, _
                           Optional ByVal cManualDK As String, Optional ByVal cRekeningJurnal As String, _
                           Optional ByVal cNow As String = "")
Dim vaField, vaValue
Dim cDK As String
Dim cKas As String

  cNow = IIf(cNow = "", SNow, cNow)
  If lWithDeleteExist Then
    Set dbData = obj.Browse(GetDSN, "MutasiTabungan", "Rekening", "Faktur", sisAssign, cFaktur)
    obj.Delete GetDSN, "MutasiTabungan", "Faktur", sisAssign, cFaktur
  End If
  
  If nTotal > 0 Then
    Set dbData = obj.Browse(GetDSN, "KodeTransaksi", , "Kode", sisAssign, cKodeTransaksi)
    If dbData.RecordCount > 0 Then
      cDK = GetNull(dbData!DK, "")
      cKas = GetNull(dbData!Kas, "")
      'cRekeningJurnal = IIf(GetNull(dbData!DK, "") = "M", cRekeningJurnal, dbData!Rekening)
      cRekeningJurnal = GetNull(dbData!Rekening, "")
    End If
    vaField = Array("Faktur", "Tgl", "KodeTransaksi", "Rekening", _
                    "Jumlah", "Keterangan", "DK", _
                    "RekeningJurnal", _
                    "UserName", "DateTime")
    vaValue = Array(cFaktur, dTgl, cKodeTransaksi, cRekening, _
                    nTotal, cKeterangan, IIf(cDK = "M", cManualDK, cDK), _
                    IIf(cKas = "K", cKasTeller, cRekeningJurnal), _
                    cusername, cNow)
    obj.Add GetDSN, "MutasiTabungan", vaField, vaValue
        
    If lUpdateToBukuBesar Then
      UpdRekTabungan obj, cFaktur
    End If
  End If
End Function

Private Function UpdSaldoTabungan(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String)
Dim vaField, vaValue
Dim nDebet As Double
Dim nKredit As Double
Dim cSQL As String

  cSQL = cSQL & "Select 'D' as Jenis,Sum(Jumlah) as Mutasi From MutasiTabungan "
  cSQL = cSQL & "Where Posting = ' ' and Rekening = '" & cRekening & "' and DK = 'D'"
  cSQL = cSQL & " Union "
  cSQL = cSQL & "Select 'K' as Jenis,Sum(Jumlah) as Mutasi From MutasiTabungan "
  cSQL = cSQL & "Where Posting = ' ' and Rekening = '" & cRekening & "' and DK = 'K'"
  
  Set dbData = obj.SQL(GetDSN, cSQL)
  If dbData.RecordCount > 0 Then
    dbData.MoveFirst
    Do While Not dbData.eof
      If dbData!Jenis = "D" Then
        nDebet = GetNull(dbData!Mutasi, 0)
      Else
        nKredit = GetNull(dbData!Mutasi, 0)
      End If
      dbData.MoveNext
    Loop
    obj.Edit GetDSN, "Tabungan", "Rekening = '" & cRekening & "'", _
                     Array("Debet", "Kredit", "Akhir"), _
                     Array(nDebet, nKredit, "&Awal+Kredit-Debet")
  End If
End Function

Sub DeleteRealisasi(ByVal obj As CodeSuiteLibrary.data, ByVal cFaktur As String, _
                    Optional ByVal lWithDebitur As Boolean = True)
                    
  If lWithDebitur Then
    obj.Delete GetDSN, "Debitur", "Faktur", sisAssign, cFaktur
  End If
'  UpdRekRealisasiPembiayaan obj, cFaktur
End Sub

Sub UpdRekJurnal(ByVal obj As CodeSuiteLibrary.data, ByVal cFaktur As String)
Dim cCabang As String

  cCabang = Mid(cFaktur, 4, 2)
  DelKodeTr obj, msJurnalLain, cCabang, cFaktur
  Set dbData = obj.Browse(GetDSN, "Jurnal j", , "j.Faktur", sisAssign, cFaktur)
  If dbData.RecordCount > 0 Then
    dbData.MoveFirst
    Do While Not dbData.eof
      UpdKodeTr obj, msJurnalLain, cCabang, cFaktur, dbData!Tgl, dbData!Rekening, dbData!Keterangan, dbData!Debet, dbData!Kredit
      dbData.MoveNext
    Loop
  End If
End Sub


Function GetSaldoTabungan1(ByVal objData As CodeSuiteLibrary.data, ByVal cRekening As String) As Double
Dim cWhere As String
Dim n As Double
Dim nSaldoAwal As Double
Dim nSaldoAkhir As Double

  nSaldoAwal = 0
  nSaldoAkhir = 0
  
  cWhere = " t.Rekening = '" & cRekening & "' "
  cWhere = cWhere & " and t.Tgl <= '" & Format(Date, "yyyy-mm-dd") & "'"
  Set dbData = objData.Browse(GetDSN, "Tabungan t", "t.AwalTahun", , , , cWhere, "t.GolonganTabungan,t.Rekening", _
               Array("Left Join RegisterNasabah r on t.Kode = r.Kode", _
                     "Left Join GolonganTabungan g on t.GolonganTabungan = g.Kode"))
  
  If dbData.RecordCount > 0 Then
    nSaldoAwal = dbData!AwalTahun
    
    cWhere = " Posting = 0 and " & cWhere
    Set dbData = objData.Browse(GetDSN, "MutasiTabungan t", "t.Rekening,t.DK,t.Jumlah", , , , cWhere)
    If dbData.RecordCount > 0 Then
      dbData.MoveFirst
      Do While Not dbData.eof
         nSaldoAkhir = nSaldoAkhir + Round(IIf((dbData!DK) = "D", -(dbData!Jumlah), (dbData!Jumlah)), 2)
        dbData.MoveNext
      Loop
    End If
    GetSaldoTabungan1 = nSaldoAkhir
  End If
End Function

Function GetNamaHari(ByVal nDayOfDate As Integer) As String
Dim vaN
Dim vaNamaHari
Dim i As Integer
Dim nHariKe As Integer

    nHariKe = nDayOfDate Mod 7
    vaNamaHari = Array("Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu")
    vaN = Array(0, 1, 2, 3, 4, 5, 6)
    For i = 0 To 6
      If vaN(i) = nHariKe Then
        GetNamaHari = vaNamaHari(i)
      End If
    Next
End Function

Sub UpdUrutFaktur(obj As CodeSuiteLibrary.data, cFaktur As String)
  If Trim(cFaktur) <> "" Then
    obj.Update GetDSN, "UrutFaktur", "Faktur ='" & cFaktur & "'", Array("Faktur"), Array(cFaktur)
  End If
End Sub

'======================
Function GetBungaRegulerPublic(ByVal nSisaPokok As Double, ByVal nBunga As Double) As Double
  GetBungaRegulerPublic = nSisaPokok * (nBunga / 100)
  GetBungaRegulerPublic = Mod50(GetBungaRegulerPublic)
End Function

Function GetTunggakanPokokHarianPublic(ByVal obj As CodeSuiteLibrary.data, ByVal nKe As Integer, ByVal cRekening As String, ByVal TglTransaksi As Date)
Dim xArray As New XArrayDB
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim n As Integer
Dim nAngsBln As Double
Dim nTotalAngsur As Double
Dim nTotalAngsur1 As Double
Dim nSisa As Double
Dim cWhere As String
Dim nSum As Double
Dim db As New ADODB.Recordset
Dim db1 As New ADODB.Recordset


  Set db = obj.Browse(GetDSN, "debitur d", "d.*", "d.Rekening", sisAssign, cRekening)
  If nKe > 1 Then
    nAngsBln = Round(GetNull(db!plafond) / GetNull(db!Lama), 2)
    dTglAkhir = DateAdd("m", nKe - 1, GetNull(db!Tgl))
    nTotalAngsur = 0
    Set db1 = obj.Browse(GetDSN, "Angsuran", "Sum(Pokok) as AngsuranPokok", "Rekening", sisAssign, cRekening, "And Tgl <= '" & Format(dTglAkhir, "yyyy-mm-dd") & "'")
    If Not db1.eof Then
      nTotalAngsur = GetNull(db1!angsuranpokok)
    End If
    nSum = nAngsBln * (nKe - 1)
    nSisa = nSum - nTotalAngsur

    If nSisa > 0 Then
      dTglAwal = DateAdd("d", 1, dTglAkhir)
      Set db1 = obj.Browse(GetDSN, "Angsuran", "Sum(Pokok) as AngsuranPokok", "Rekening", sisAssign, cRekening, "And Tgl > '" & Format(dTglAkhir, "yyyy-mm-dd") & "' And Tgl <= '" & Format(TglTransaksi, "yyyy-mm-dd") & "' Group By rekening")
      If Not db1.eof Then
        nTotalAngsur1 = GetNull(db1!angsuranpokok)
      End If
      
      If (nTotalAngsur + nTotalAngsur1) >= (nSisa + nTotalAngsur) Then
        GetTunggakanPokokHarianPublic = 0
      Else
        GetTunggakanPokokHarianPublic = (nSum + nTotalAngsur1) - (nTotalAngsur1 + nTotalAngsur)
      End If
    End If
  End If
End Function

Function GetTunggakanBungaHarianPublic(ByVal obj As CodeSuiteLibrary.data, ByVal nKe As Integer, ByVal cRekening As String, ByVal TglTransaksi As Date)
Dim xArray As New XArrayDB
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim n As Integer
Dim nTotalAngsur As Double
Dim nTotalAngsur1 As Double
Dim nSisa As Double
Dim cWhere As String
Dim nSum As Double
Dim nTotBunga As Double
Dim nBD As Double
Dim nBungaPerHari As Double
Dim nBunga1 As Double
Dim nBunga2 As Double
Dim db As New ADODB.Recordset
Dim db1 As New ADODB.Recordset

  Set db = obj.Browse(GetDSN, "Debitur d", "d.*", "d.Rekening", sisAssign, cRekening)
  If nKe > 1 Then
    nTotBunga = 0
    nBunga1 = 0
    nBunga2 = 0
    For n = 1 To nKe - 1
      nBD = GetNull(db!plafond)
      If n > 1 Then
        'nBD = GetBD(n, cRekening)
      End If
      nTotBunga = nTotBunga + GetBungaRegulerPublic(nBD, GetNull(db!SukuBunga) / 12)
      nBungaPerHari = nTotBunga / GetNull(db!PeriodeBungaMenurun)
      nBunga1 = nBungaPerHari * GetNull(db!MinimalPeriode)
      nBunga2 = nBunga2 + nBunga1
    Next
    
    dTglAkhir = DateAdd("m", nKe - 1, GetNull(db!Tgl))
    nTotalAngsur = 0
    
    Set db1 = obj.Browse(GetDSN, "Angsuran", "Sum(Bunga) as AngsuranBunga", "Rekening", sisAssign, cRekening, "And Tgl <= '" & Format(dTglAkhir, "yyyy-mm-dd") & "'")
    If Not db1.eof Then
      nTotalAngsur = GetNull(db1!AngsuranBunga)
    End If
    nSum = nBunga2
    nSisa = nSum - nTotalAngsur
    
    If nSisa > 0 Then
      dTglAwal = DateAdd("d", 1, dTglAkhir)
      Set db1 = obj.Browse(GetDSN, "Angsuran", "Sum(Bunga) as AngsuranBunga", "Rekening", sisAssign, cRekening, "And Tgl > '" & Format(dTglAkhir, "yyyy-mm-dd") & "' And Tgl <= '" & Format(TglTransaksi, "yyyy-mm-dd") & "'")
      If Not db1.eof Then
        nTotalAngsur1 = GetNull(db1!AngsuranBunga)
      End If
      If (nTotalAngsur + nTotalAngsur1) >= (nSisa + nTotalAngsur) Then
        GetTunggakanBungaHarianPublic = 0
      Else
        GetTunggakanBungaHarianPublic = (nSum + nTotalAngsur1) - (nTotalAngsur + nTotalAngsur1)
      End If
    End If
  End If
End Function

Sub GetAngsuranPokokBungaKredit(ByVal obj As CodeSuiteLibrary.data, ByVal Rekening As String, ByVal TglTransaksi As Date, ByRef nAngBunga As Double, ByRef nAngPokok As Double)
Dim n As Single
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim nSukuBungaPerBulan As Double
Dim xArray As New XArrayDB
Dim dTanggalAwal As Date
Dim dTanggalAkhir As Date
Dim nBD As Double
Dim cRek As String
Dim nBulanKe As Integer
Dim nAngsPokokPerBulan As Double
Dim nPK As Double
Dim db As New ADODB.Recordset
Dim db1 As New ADODB.Recordset
Dim nTunggakanBunga As Double
Dim nTunggakanPokok As Double

  Set db = obj.Browse(GetDSN, "Debitur d", "d.*", "d.Rekening", sisAssign, Rekening)

  xArray.ReDim 0, GetNull(db!Lama), 0, 1
  dTglAwal = (DateAdd("d", 1, GetNull(db!Tgl)))
  dTglAkhir = (DateAdd("m", 1, GetNull(db!Tgl)))
  
  For n = 1 To GetNull(db!Lama)
    xArray(n, 0) = dTglAwal
    xArray(n, 1) = dTglAkhir
    dTglAwal = (DateAdd("d", 1, dTglAkhir))
    dTglAkhir = (DateAdd("m", 1, dTglAwal))
  Next
  
  For n = 1 To xArray.UpperBound(1)
    If Between(TglTransaksi, xArray(n, 0), xArray(n, 1)) Then
      dTanggalAkhir = DateAdd("m", -1, xArray(n, 1)) - 1
      nBulanKe = n
      Exit For
    End If
  Next

  If nBulanKe <= 1 Then
    nBD = GetNull(db!plafond)
  Else
    cRek = Rekening
    Set db1 = obj.Browse(GetDSN, "Angsuran", "Sum(Pokok) as Pokok", "Rekening", sisAssign, cRek, "And tgl <='" & Format(dTanggalAkhir, "yyyy-mm-dd") & "' Group By Rekening", "Tgl")
    If Not db1.eof Then
      nBD = GetNull(db!plafond) - GetNull(db1!pokok)
    End If
  End If
  
  nAngsPokokPerBulan = Round(GetNull(db!plafond) / GetNull(db!Lama), 2)
  If nAngsPokokPerBulan > nBD Then
    nPK = nBD
  Else
    nPK = GetNull(db!plafond)
  End If
  
  nSukuBungaPerBulan = Round(GetNull(db!SukuBunga) / 12, 2)
  nAngPokok = Round((nPK / GetNull(db!Lama)) / GetNull(db!PeriodeBungaMenurun))
  nAngBunga = GetBungaRegulerPublic(nBD, nSukuBungaPerBulan)
  nAngBunga = Round(nAngBunga / GetNull(db!PeriodeBungaMenurun))
  nTunggakanPokok = GetTunggakanPokokHarianPublic(obj, nBulanKe, cRek, TglTransaksi)
  nTunggakanBunga = GetTunggakanBungaHarianPublic(obj, nBulanKe, cRek, TglTransaksi)
End Sub

Function IsDalamPeriode(ByVal Tgl As Date, ByVal dTglReal As Date, ByVal dDueDate As Date, ByVal Lama As Integer) As Boolean
Dim n As Single
Dim dTglAwal As Date
Dim dTglAkhir As Date
Dim xArray As New XArrayDB
Dim dTanggalAwal As Date
Dim dTanggalAkhir As Date
'Dim nBulanKe As Integer
Dim nAngsPokokPerBulan As Double
Dim db As New ADODB.Recordset

  IsDalamPeriode = False
  xArray.ReDim 0, Lama, 0, 1
  dTglAwal = (DateAdd("d", 1, dTglReal))
  dTglAkhir = (DateAdd("m", 1, dTglReal))
  Lama = DateDiff("m", dTglAwal, dTglAkhir)
  For n = 1 To Lama
    xArray(n, 0) = dTglAwal
    xArray(n, 1) = dTglAkhir
    dTglAwal = (DateAdd("d", 1, dTglAkhir))
    dTglAkhir = (DateAdd("m", 1, dTglAwal))
  Next
  
  For n = 1 To xArray.UpperBound(1)
    If Between(Tgl, xArray(n, 0), xArray(n, 1)) Then
      dTanggalAkhir = DateAdd("m", -1, xArray(n, 1)) - 1
'      'Lihat apakah dalam bulan ini sudah pernah posting bunga?
'      Set db = obj.Browse(GetDSN, "MutasiBungaDeposito", , "Rekening", sisAssign, Rekening, " and month(tgl)=" & Month(Tgl) & "")
'      If Not db.eof Then
'      Else
'
'      End If
'      'nBulanKe = n
        IsDalamPeriode = True
      Exit For
    End If
  Next
End Function
