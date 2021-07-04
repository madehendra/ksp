Attribute VB_Name = "modPerhitunganFlat"
Option Explicit

Function fGetPeriode(ByVal obj As CodeSuiteLibrary.data, ByVal cRekeningKredit As String, ByVal dTgl As Date, ByRef nPeriode As Integer, ByRef nLate As Integer) As Integer
Dim db As New ADODB.Recordset
Dim n As Integer
Dim dAwal As Date
Dim dAkhir As Date

'Fungsi ini untuk menentukan posisi periode pembayaran angsuran nTgl dari cRekeningKredit

  Set db = obj.Browse(GetDSN, "debitur", "rekening,lama,konpensasitelat,tgl", "rekening", sisAssign, cRekeningKredit)
  If Not db.eof Then
    dAwal = DateAdd("d", 1, GetNull(db!Tgl))
    dAkhir = DateAdd("d", GetNull(db!KonpensasiTelat), DateAdd("m", 1, dAwal))
    For n = 1 To GetNull(db!Lama)
      If dTgl >= dAwal And dTgl <= dAkhir Then
        nPeriode = n - 1
        nLate = DateDiff("d", dAwal, dTgl)
        fGetPeriode = n
        db.Close
        Exit Function
      End If
      dAwal = DateAdd("d", 1, dAkhir)
      dAkhir = DateAdd("m", 1, DateAdd("d", -1, dAwal))
    Next n
    nPeriode = GetNull(db!Lama)
    nLate = DateDiff("d", dAkhir, dTgl)
    fGetPeriode = nPeriode + 1
  End If
  db.Close
End Function


Sub fGetBungaPokokPeriodeKe(ByVal obj As CodeSuiteLibrary.data, ByVal cRekeningKredit As String, ByVal nPeriodeKe As Integer, ByRef nPokokAngsuran As Double, ByRef nBungaAngsuran As Double)
Dim db As New ADODB.Recordset

  nPokokAngsuran = 0
  nBungaAngsuran = 0
  Set db = obj.Browse(GetDSN, "debitur", , "rekening", sisAssign, cRekeningKredit)
  If Not db.eof Then
    nPokokAngsuran = nPeriodeKe * (GetNull(db!plafond) / GetNull(db!Lama))
    nBungaAngsuran = nPeriodeKe * (GetNull(db!plafond) * GetNull(db!SukuBunga) / 12 / 100)
  End If
  db.Close
End Sub

Sub fGetBungaPokok(ByVal obj As CodeSuiteLibrary.data, ByVal cRekeningKredit As String, ByRef nPokokAngsuran As Double, ByRef nBungaAngsuran As Double)
Dim db As New ADODB.Recordset

  nPokokAngsuran = 0
  nBungaAngsuran = 0
  Set db = obj.Browse(GetDSN, "angsuran", "sum(pokok) as pokok, sum(bunga) as bunga", "rekening", sisAssign, cRekeningKredit)
  If Not db.eof Then
    nPokokAngsuran = GetNull(db!pokok)
    nBungaAngsuran = GetNull(db!bunga)
  End If
  db.Close
End Sub
