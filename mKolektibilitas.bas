Attribute VB_Name = "mKolektibilitas"
Option Explicit

Function GetKolek(ByVal obj As CodeSuiteLibrary.data, ByVal cRek As String, ByVal dTgl As Date, ByVal nWajibPokok As Double) As String
Dim n As Integer
Dim db As New ADODB.Recordset
  
'Klasifikasi 1 = LANCAR
'Apabila tidak terdapat tunggakan pokok dan bunga

'Klasifikasi 2 = DALAM PERHATIAN KHUSUS
'Apabila terdapat tunggakan pokok atau bunga sampai dengan 3 bulan

'Klasifikasi 3 = KURANG LANCAR
'Apabila terdapat tunggakan pokok atau bunga sampai dengan 6 bulan

'Klasifikasi 4 = DIRAGUKAN
'Apabila terdapat tunggakan pokok atau bunga sampai dengan 9 bulan

'Klasifikasi 5 = MACET
'Apabila terdapat tunggakan pokok atau bunga diatas 9 bulan

'n = GetLate(obj, cRek, dTgl)
'If n > 9 Then
'  GetKolek = "5. MACET"
'ElseIf n <= 9 And n > 6 Then
'  GetKolek = "4. DIRAGUKAN"
'ElseIf n <= 6 And n > 3 Then
'  GetKolek = "3. KURANG LANCAR"
'ElseIf n <= 3 And n > 0 Then
'  GetKolek = "2. DPK"
'Else
'  GetKolek = "1. LANCAR"
'End If


'Mitra Abadi

'n = GetLate(obj, cRek, dTgl, nWajibPokok)
'If n > 5 Then
'  GetKolek = "4. MACET"
'ElseIf n <= 6 And n > 3 Then
'  GetKolek = "3. DIRAGUKAN"
'ElseIf n <= 3 And n > 0 Then
'  GetKolek = "2. KURANG LANCAR"
'Else
'  GetKolek = "1. LANCAR"
'End If

n = GetTelatBayarBunga(obj, cRek)
'kurang lancar
'diragukan
'macet

If n >= 12 Then
  GetKolek = "Macet"
ElseIf n < 12 And n >= 6 Then
  GetKolek = "Diragukan"
ElseIf n < 6 And n >= 3 Then
  GetKolek = "Kurang Lancar"
ElseIf n < 3 Then
  GetKolek = "lancar"
End If
End Function

Sub GetDiff(dtgl1 As Date, dTgl2 As Date, ByRef nM As Integer, ByRef nD As Integer)
Dim dTmp As Date

  nM = 0
  nD = 0
  Do While dtgl1 <= dTgl2
    If DateDiff("d", dtgl1, dTgl2) <= 31 Then
      nD = DateDiff("d", dtgl1, dTgl2)
      Exit Sub
    Else
      nM = nM + 1
      dTmp = DateAdd("m", 1, dtgl1)
      dtgl1 = dTmp
    End If
  Loop
End Sub

Private Function GetLate(ByVal obj As CodeSuiteLibrary.data, ByVal cRek As String, ByVal dTgl As Date, ByVal nWajibPokok As Double) As Double
  GetLate = GetTunggakanPokok(obj, cRek, nWajibPokok) + GetTelatBulan(obj, cRek, dTgl)
End Function

Private Function GetTunggakanPokok(ByVal obj As CodeSuiteLibrary.data, ByVal cRek As String, ByVal nWajibPokok As Double) As Double
Dim db As New ADODB.Recordset
Dim nJumlahAngsuran As Integer
'Dim nWajibPokok As Double
Dim nJumlahAngsuranPokok As Double
Dim n As Integer

'  nWajibPokok = nWajibPokok
  GetTunggakanPokok = 0
  Set db = obj.Browse(GetDSN, "angsuran", "count(faktur) as jumlahangsuran", "rekening", sisAssign, cRek)
  If Not db.eof Then
    nJumlahAngsuran = GetNull(db!jumlahangsuran, 0)
  End If
  
'  Set db = obj.Browse(GetDSN, "debitur", "wajibpokok", "rekening", sisAssign, cRek)
'  If Not db.eof Then
'    nWajibPokok = GetNull(db!wajibpokok)
'  End If
  
  Set db = obj.Browse(GetDSN, "angsuran", "sum(pokok) as angsuranpokok", "rekening", sisAssign, cRek)
  If Not db.eof Then
    nJumlahAngsuranPokok = GetNull(db!angsuranpokok, 0)
  End If
  
  If nWajibPokok > 0 Then
    n = DevideMod(nJumlahAngsuranPokok, nWajibPokok)
    GetTunggakanPokok = nJumlahAngsuran - n
  End If
  
End Function

Private Function GetTelatBayarBunga(ByVal obj As CodeSuiteLibrary.data, ByVal cRek As String) As Single
Dim dbData As New ADODB.Recordset
Dim nSelisihBulan As Single
Dim nJumlahXBunga As Single

  GetTelatBayarBunga = 0
  Set dbData = obj.Browse(GetDSN, "debitur", "tgl,lama", "rekening", sisAssign, cRek)
  If Not dbData.eof Then
    nSelisihBulan = DateDiff("m", Format(GetNull(dbData!Tgl), "yyyy-MM-dd"), Format(Date, "yyyy-MM-dd"))
  End If
  
  Set dbData = obj.Browse(GetDSN, "angsuran", "count(bunga) as jumlah", "rekening", sisAssign, cRek, " and bunga <> 0")
  If Not dbData.eof Then
    nJumlahXBunga = GetNull(dbData!Jumlah)
  End If

  GetTelatBayarBunga = nSelisihBulan - nJumlahXBunga
End Function

Private Function GetTelatBulan(ByVal obj As CodeSuiteLibrary.data, ByVal cRek As String, ByVal dTgl As Date) As Double
Dim db As New ADODB.Recordset
Dim nJumlahAngsuran As Integer
Dim n As Integer

  GetTelatBulan = 0
  Set db = obj.Browse(GetDSN, "angsuran", "count(faktur) as jumlahangsuran", "rekening", sisAssign, cRek)
  If Not db.eof Then
    nJumlahAngsuran = GetNull(db!jumlahangsuran, 0)
  End If
  
  Set db = obj.Browse(GetDSN, "debitur", "tgl", "rekening", sisAssign, cRek)
  If Not db.eof Then
    n = DateDiff("m", Format(GetNull(db!Tgl), "yyyy-MM-dd"), Format(EOM(dTgl), "yyyy-MM-dd"))
  End If
  If nJumlahAngsuran < n - 1 Then
    GetTelatBulan = (n - 1) - nJumlahAngsuran
  End If
End Function
