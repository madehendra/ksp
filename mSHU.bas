Attribute VB_Name = "mSHU"
Option Explicit

Function GetRugiLabaSHU(ByVal obj As CodeSuiteLibrary.data, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset

  GetRugiLabaSHU = 0
  cSQL = "select sum(kredit-debet) as labarugi  from bukubesar where (left(rekening,1) = '5' or left(rekening,1) = '4') and tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.eof Then
    GetRugiLabaSHU = GetNull(db!labarugi)
  End If
End Function

Function GetRatioBahasKredit(ByVal obj As CodeSuiteLibrary.data, ByVal cDebitur As String, ByVal dAwal As Date, ByVal dAkhir As Date, ByVal nTotalKreditMengendap As Double)
Dim db As New ADODB.Recordset
Dim nRatio As Double

  nRatio = Devide(GetKreditMengendap(obj, cDebitur, dAwal, dAkhir), nTotalKreditMengendap) * 100
End Function

Function GetKreditMengendap(ByVal obj As CodeSuiteLibrary.data, ByVal cDebitur As String, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
Dim db As New ADODB.Recordset
Dim nTotalKredit As Double

  nTotalKredit = 0
  Set db = obj.Browse(GetDSN, "debitur", "rekening,kode,plafond", "kode", sisAssign, cDebitur, " and tgl >='" & Format(dAwal, "yyyy-MM-dd") & "' and tgl <= '" & Format(dAkhir, "yyyy-MM-dd") & "'")
  If Not db.eof Then
    Do While Not db.eof
      If isValidMengendap(obj, GetNull(db!Rekening)) Then
        nTotalKredit = nTotalKredit + GetNull(db!plafond)
      End If
      db.MoveNext
    Loop
  End If
  GetKreditMengendap = nTotalKredit
End Function

Private Function isValidMengendap(ByVal obj As CodeSuiteLibrary.data, ByVal cRek As String) As Boolean
Dim db As New ADODB.Recordset

  isValidMengendap = False
  Set db = obj.Browse(GetDSN, "angsuran", "count(faktur) as angsuran", "rekening", sisAssign, cRek)
  If Not db.eof Then
    If GetNull(db!angsuran) >= 2 Then
      isValidMengendap = True
    End If
  End If
End Function

Function GetTotalKreditMengendap(ByVal obj As CodeSuiteLibrary.data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
Dim db As New ADODB.Recordset
Dim nTotalKredit As Double

  nTotalKredit = 0
  Set db = obj.Browse(GetDSN, "debitur", , , , , " 1=1 and tgl >='" & Format(dAwal, "yyyy-MM-dd") & "' and tgl <= '" & Format(dAkhir, "yyyy-MM-dd") & "'")
  If Not db.eof Then
    Do While Not db.eof
      If isValidMengendap(obj, GetNull(db!Rekening)) Then
        nTotalKredit = nTotalKredit + GetNull(db!plafond)
      End If
      db.MoveNext
    Loop
  End If
  GetTotalKreditMengendap = nTotalKredit
End Function

Function GetTotalDepositoMengendap(ByVal obj As CodeSuiteLibrary.data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
Dim db As New ADODB.Recordset

  GetTotalDepositoMengendap = 0
  Set db = obj.Browse(GetDSN, "deposito", "sum(nominaldeposito) as nominaldeposito", "tgl", sisGTEqual, Format(DateAdd("m", -1, dAwal), "yyyy-MM-dd"), " and tgl <='" & Format(DateAdd("m", -1, dAkhir), "yyyy-MM-dd") & "'")
  If Not db.eof Then
    GetTotalDepositoMengendap = GetNull(db!nominaldeposito)
  End If
  
  Set db = obj.Browse(GetDSN, "deposito", "sum(nominaldeposito) as nominaldeposito", "lastperpanjangan", sisGTEqual, Format(DateAdd("m", -1, dAwal), "yyyy-MM-dd"), " and lastperpanjangan <='" & Format(DateAdd("m", -1, dAkhir), "yyyy-MM-dd") & "'")
  If Not db.eof Then
    GetTotalDepositoMengendap = GetTotalDepositoMengendap + GetNull(db!nominaldeposito)
  End If
End Function

Function GetTotalDepositoMengendapAnggota(ByVal obj As CodeSuiteLibrary.data, ByVal dAwal As Date, ByVal dAkhir As Date, ByVal cRegister As String) As Double
Dim db As New ADODB.Recordset

  GetTotalDepositoMengendapAnggota = 0
  Set db = obj.Browse(GetDSN, "deposito d", "sum(d.nominaldeposito) as nominaldeposito", "d.tgl", sisGTEqual, Format(DateAdd("m", -1, dAwal), "yyyy-MM-dd"), " and d.tgl <='" & Format(DateAdd("m", -1, dAkhir), "yyyy-MM-dd") & "' and d.kode = '" & cRegister & "'", , Array("left join registernasabah r on r.kode = d.kode"))
  If Not db.eof Then
    GetTotalDepositoMengendapAnggota = GetNull(db!nominaldeposito)
  End If
  
  Set db = obj.Browse(GetDSN, "deposito d", "sum(d.nominaldeposito) as nominaldeposito", "d.lastperpanjangan", sisGTEqual, Format(DateAdd("m", -1, dAwal), "yyyy-MM-dd"), " and d.lastperpanjangan <='" & Format(DateAdd("m", -1, dAkhir), "yyyy-MM-dd") & "' and d.kode = '" & cRegister & "'", , Array("left join registernasabah r on r.kode = d.kode"))
  If Not db.eof Then
    GetTotalDepositoMengendapAnggota = GetTotalDepositoMengendapAnggota + GetNull(db!nominaldeposito)
  End If
End Function

Sub PostingSaldoTerendah(obj As CodeSuiteLibrary.data, dTgl As Date)
Dim db As New ADODB.Recordset
Dim dTglValuta As Date
Dim dTglClose As Date
Dim n As Integer
Dim dNow As Date
Dim dNext As Date
Dim nMonth As Integer
Dim nYear As Integer
Dim PB As New FrmPB
Dim lValuta As Boolean

  
  Set db = obj.SQL(GetDSN, "select rekening,tgl,close,tglpenutupan,kode,golongantabungan from tabungan where golongantabungan = 'T1' or golongantabungan = 'T2'")
  '01.T1.000001.01
'  Set db = obj.SQL(GetDSN, "select rekening,tgl,close,tglpenutupan,kode,golongantabungan from tabungan where rekening = '01.T1.000001.01'")
  If Not db.eof Then
    obj.SQL GetDSN, "delete from simpananmengendap"
    PB.InitPB db.RecordCount
    db.MoveFirst
    Do While Not db.eof
      lValuta = True
      PB.RunPB
      dTglValuta = GetNull(db!Tgl)
      If GetNull(db!Close) <> 1 Then
        dTglClose = dTgl
      Else
        dTglClose = GetNull(db!TglPenutupan)
      End If
      n = DateDiff("m", dTglValuta, dTglClose)
      dNow = dTglValuta
      For n = 1 To n
        dNext = DateAdd("m", 1, dNow)
        If dTglClose > dNext Then
          nMonth = Month(dNext)
          nYear = Year(dNext)
          'jalankan rutin untuk mendapatkan saldo minimum bulan dNext
          'setelah mendapatkan saldo minimum, simpan pada tabel untuk bulan nMonth
          obj.Update GetDSN, "simpananmengendap", "rekening = '" & db!Rekening & "' and tahun = '" & nYear & "' and bulan = '" & nMonth & "'", Array("rekening", "tahun", "bulan", "jumlah", "kode", "golongantabungan"), Array(GetNull(db!Rekening), nYear, nMonth, GetSaldoTerendah(obj, GetNull(db!Rekening), dNow, dNext, lValuta), GetNull(db!Kode), GetNull(db!GolonganTabungan))
          lValuta = Not lValuta
          dNow = dNext
        Else
          Exit For
        End If
      Next n
      db.MoveNext
    Loop
    PB.EndPB
  End If
  
End Sub

Function GetSaldoTerendah(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date, ByVal lValuta As Boolean) As Double
Dim dba As New ADODB.Recordset
Dim nAwal As Double
Dim nSaldoTerendah As Double

  nAwal = 0
  Set dba = obj.Browse(GetDSN, "mutasitabungan", "rekening,jumlah,dk", "rekening", sisAssign, cRekening, " and tgl < '" & Format(dTglAwal, "yyyy-MM-dd") & "'", "tgl,faktur")
  If Not dba.eof Then
    dba.MoveFirst
    Do While Not dba.eof
      nAwal = nAwal + IIf(dba!DK = "K", GetNull(dba!Jumlah), -GetNull(dba!Jumlah))
      dba.MoveNext
    Loop
  End If
  
  nSaldoTerendah = nAwal
  Set dba = obj.Browse(GetDSN, "mutasitabungan", "rekening,jumlah,dk", "rekening", sisAssign, cRekening, " and tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' and tgl < '" & Format(dTglAkhir, "yyyy-MM-dd") & "'", "tgl,faktur")
  If Not dba.eof Then
    dba.MoveFirst
    Do While Not dba.eof
      nAwal = nAwal + IIf(dba!DK = "K", GetNull(dba!Jumlah), -GetNull(dba!Jumlah))
      If nAwal < nSaldoTerendah Then
        nSaldoTerendah = nAwal
      End If
      dba.MoveNext
    Loop
  End If
  dba.Close
  GetSaldoTerendah = nSaldoTerendah
End Function

Function PostingAkhirHariPengendapan(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String, ByVal dTgl As Date) As Boolean
Dim db As New ADODB.Recordset

  PostingAkhirHariPengendapan = False
  Set db = obj.Browse(GetDSN, "tabungan", , "close", sisDifference, "1", " and Rekening = '" & cRekening & "'")
  If Not db.eof Then
    Do While Not db.eof
      If Day(GetNull(db!Tgl)) = Day(dTgl) Then
        'lakukan rutin pengendapan
        PostingAkhirHariPengendapan = True
      End If
      'jika tgl sekarang adalah akhir bulan, maka cek tabungan yang tgl valutanya lebih dari tgl akhir bulan
      If Day(dTgl) = Day(EOM(dTgl)) Then
        'ambil sisa tanggal
        If Day(GetNull(db!Tgl)) > Day(dTgl) Then
          'lakukan rutin pengendapan
          PostingAkhirHariPengendapan = True
        End If
      End If
      db.MoveNext
    Loop
  End If
End Function
