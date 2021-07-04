Attribute VB_Name = "Func"
Option Explicit

Public nUserLevel As Single
Public cUserID As String
Public cusername As String
Public cFullName As String
Public cKasTeller As String
Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset

Public Enum SisPos
  Normal = 0
  Add = 1
  Edit = 2
  Delete = 3
End Enum

Public Enum SisCfg
  msKodeKas = 0
  msNama = 1
  msAlamat = 2
  msTelepon = 3
  msFax = 4
  msKota = 5
  msEmail = 6
  msProvinsi = 7
  msKodeCabang = 8
  msLebarGolongan = 9
  mskodebagihasil = 10
  msKodePajakBagiHasil = 11
  msKodePemindahBukuan = 12
  msPicturePath = 13
  msBiayaAdministrasiPenutupan = 14
  msKodeTutupBuku = 15
  msKodeSetoranTunai = 16
  msKodePenarikanTunai = 17
  msKodePenarikanPemindahBukuan = 18
  msKodePenyetoranPemindahbukuan = 19
  msKodePembulatankas = 20
  msKodeAdministrasi = 21
  msKodelaba = 22
  msTopValidasiTabungan = 23
  msLeftValidasiTabungan = 24
  msTopBilyetDeposito = 25
  msLeftBilyetDeposito = 26
  msVersion = 27
  msNamaDirut = 28
  msKodeKasInduk = 29
  msPostingBungaTabungan = 30
  msDefaultTeller = 31
  msLockTeller = 32
  msKodeTransaksiPB = 33
  msSerialKey = 34
  
    
  msNamaPembuatNeraca = 80
  msJabatanPembuatNeraca = 81
  msNamaPemeriksaNeraca = 82
  msJabatanPemeriksaNeraca = 83
  msOptTampilkanFootNoteNeraca = 84
  
  '=----- SHU
  msSHUJasaUsaha = 90
  msSHUModal = 91
  msSHUPinjaman = 92
  msSHUDeposito = 93
  msSHUSimpanan = 94
  msSHUKodeGolonganSimpananHarian = 95
End Enum

Public Enum SisTypeRekening
  SisAktiva = 1
  SisHutang = 2
  SisModal = 3
  SisPendapatan = 4
  SisBiaya = 5
  SisAdministratif = 6
End Enum

Public Enum SisJenisProduk
  Tabungan = 1
  Deposito = 2
  Kredit = 3
End Enum

Public Enum SisJenisPerhitungan
  SisFlat = 1
  SisReguler = 2
  SisAnuitas = 3
End Enum

Public Enum sisSimpan
  sisUpdate = 1
  sisAdd = 2
  sisEdit = 3
End Enum

Public Enum SisRekeningNasabah
  rek_Tabungan = 0
  rek_Kredit = 1
  rek_Deposito = 2
  rek_RegisterNasabah = 3
End Enum

Public Enum SisFaktur
  fkt_MutasiTabungan = 0
  fkt_Angsuran = 1
  fkt_Relisasi = 2
  fkt_Jurnal = 3
  fkt_Deposito = 4
  fkt_Titipan = 5
End Enum

Function GetLastFaktur(ByVal nPar As SisFaktur, ByVal dTgl As Date, Optional ByVal lUpdate As Boolean = False) As String
Dim vaFaktur
Dim db As New ADODB.Recordset
Dim obj As New CodeSuiteLibrary.data
Dim cChar As String
Dim cNomor As String
Dim nCount As Double
Dim cKode As String

  cNomor = 1
  vaFaktur = Array("TB", "AG", "R0", "JR", "DP", "TT")
  cChar = vaFaktur(nPar)
  cKode = cChar & aCfg(msKodeCabang) & Format(dTgl, "yyyyMMdd")
  If lUpdate Then
    
    obj.Add GetDSN, "NomorFaktur", Array("Kode"), Array(cKode)
    Set db = obj.SQL(GetDSN, "Select Last_Insert_id() as Total")
    
    ' Untuk Menghemat Ukuran Table Hapus jika Nomor ID < Nomor yang aktif
    If db.RecordCount > 0 Then
      obj.Delete GetDSN, "NomorFaktur", "Kode", sisAssign, cKode, " and id < " & GetNull(db!Total)
    End If
    
    
    nCount = 0
  Else
    Set db = obj.Browse(GetDSN, "NomorFaktur", "Max(ID) as Total", "Kode", sisAssign, cKode)
    nCount = 1
  End If
  If db.RecordCount > 0 Then
    cNomor = GetNull(db!Total, 0) + nCount
  End If

  cNomor = cChar & aCfg(msKodeCabang) & Format(dTgl, "yyyyMMdd") & Padl(Trim(cNomor), 8, "0")
  GetLastFaktur = cNomor
End Function

Function TypeRekening(ByVal cRekening) As SisTypeRekening
  cRekening = left(cRekening, 1)
  TypeRekening = Val(cRekening)
End Function

Function RekSpace(cRekening, cKeterangan) As String
Dim n As Single, nDot As Single
  For n = 1 To Len(cRekening)
    If Mid(cRekening, n, 1) = "." Then
      nDot = nDot + 1
    End If
  Next
  If nDot >= 1 Then
    RekSpace = Space((nDot - 1) * 4) & cKeterangan
  End If
End Function

Function GetInduk(Optional cRekening As String = "", Optional lPad As Boolean = True)
Dim lStop As Boolean, nOldLen As Byte
  lStop = False
  nOldLen = Len(cRekening)
  Do While Not lStop
    cRekening = Trim(cRekening)
    If Right(cRekening, 1) = "." Then
      cRekening = left(cRekening, Len(cRekening) - 1)
    Else
      lStop = True
    End If
  Loop
  If lPad Then
    GetInduk = Pad(cRekening, nOldLen)
  Else
    GetInduk = cRekening
  End If
End Function

Function GetDetail(cRekening, Optional ByVal cFormat As String = "999.99.99.999", _
                   Optional ByVal nLeft As Single = 3) As String
Dim cPict As String
  cPict = Replace(cFormat, "9", " ")
  cRekening = Mid(cRekening, nLeft)
  cRekening = Trim(cRekening)
  GetDetail = cRekening & Right(cPict, Len(cPict) - Len(cRekening))
End Function


Function InitConnection(Optional pAuto As Boolean = False)
'  GetNotifikasiAdd "Melakukan Koneksi Ke Server"
  If Not SetAuto() Then
   ' On Error Resume Next
    Set GetDSN = New ADODB.Connection
    'GetDSN.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=True;Data Source=" & GetRegistry(reg_DSN)
    GetDSN.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=True;Data Source=" & GetRegistry(reg_DSN)
    GetDSN.CursorLocation = adUseClient
'     GetDSN.CursorLocation = adUseServer
'     GetDSN
    GetDSN.Open

  End If
  
  If pAuto Then
    SetAuto pAuto
  End If
'  GetNotifikasiRemove
End Function

Function SetAuto(Optional lAuto As Variant = Null) As Boolean
Static l As Boolean
  SetAuto = l
  If Not IsNull(lAuto) Then
    l = lAuto
  End If
End Function

Sub InitCfg()
'Dim cTipe As String
'On Error GoTo salah:
'
'  Set dbData = objData.Browse(GetDSN, "config")
'  If dbData.RecordCount > 0 Then
'    Do While Not dbData.eof
'      cTipe = IIf(dbData!Tipe = "D", "[D]", "[C]")
'      SaveSetting "USPD", App.EXEName, "Cfg" & dbData!jenis, cTipe & dbData!Keterangan
'      dbData.MoveNext
'    Loop
'  End If
'  Exit Sub
'salah:
'  If err.Number = 3704 Then
'    MsgBox "FATAL ERROR, Program tidak bisa terkoneksi dengan database!!", vbCritical, "FATAL ERROR"
'    End
'  End If
  
  
Dim cTipe As String
On Error GoTo salah

  Set dbData = objData.Browse(GetDSN, "config")
  If dbData.RecordCount > 0 Then
    Do While Not dbData.eof
      cTipe = IIf(GetNull(dbData!Tipe) = "D", "[D]", "[C]")
      SaveSetting "madehendra-ksp", App.EXEName, "Cfg" & GetNull(dbData!Jenis), cTipe & GetNull(dbData!Keterangan)
      dbData.MoveNext
    Loop
  End If
  
salah:
If err.Number = 3704 Then
  MsgBox "Invalid database!!", vbExclamation
  End
End If
End Sub

Function aCfg(ByVal par As SisCfg, Optional cDefault As Variant = "") As Variant
Dim vRetval As Variant
Dim cTipe As String
Dim cValue As String

  vRetval = GetSetting("USPD", App.EXEName, "Cfg" & par, "[C]" & cDefault)
  cTipe = left(vRetval, 3)
  cValue = Mid(vRetval, 4)
  Select Case cTipe
    Case "[D]"
      aCfg = DateSerial(left(cValue, 4), Mid(cValue, 5, 2), Mid(cValue, 7, 2))
    Case Else
      aCfg = cValue
  End Select
End Function

Function UpdCfg(par As SisCfg, Keterangan, Optional ByVal obj As Variant = Null)
Dim cType As String
  If IsNull(obj) Then
    Set obj = New CodeSuiteLibrary.data
  End If

  cType = "C"
  If VarType(Keterangan) = vbDate Then
    Keterangan = Format(Keterangan, "yyyymmdd")
    cType = "D"
  End If

  SaveSetting "USPD", App.EXEName, "Cfg" & par, "[" & cType & "]" & Keterangan
  obj.Update GetDSN, "config", "jenis = '" & par & "'", Array("jenis", "Keterangan", "tipe"), Array(par, Keterangan, cType)
End Function

Function CenterForm(bForm As Form, Optional ByVal lZeroTopLeft As Boolean = False)
  If lZeroTopLeft Then
    bForm.left = 0
    bForm.Top = 0
  Else
    bForm.left = (Screen.Width / 2) - (bForm.Width / 2) - 100
    bForm.Top = (Screen.Height / 2) - (bForm.Height / 2) - 750
   End If
  bForm.Icon = aMainmenu.Icon
End Function

Function GetLevel(ByVal cRekening As String, ByVal nLevel As Single)
Dim nLeft As Single
  Select Case nLevel
    Case 1
      nLeft = 5
    Case 2
      nLeft = 8
    Case 3
      nLeft = 11
    Case 4
      nLeft = 15
  End Select
  GetLevel = left(cRekening, nLeft)
End Function

Function MaxLevel() As Single
Dim n As Single, nCount As Single, cRekening As String

  cRekening = Trim("999.99.99.999")
  nCount = 1
  For n = 1 To Len(cRekening)
    If Mid(cRekening, n, 1) = "." Then
      nCount = nCount + 1
    End If
  Next
  MaxLevel = nCount
End Function

Function Level(cRekening As String) As Single
Dim n As Single, nCount As Single
  cRekening = Mid(Trim(cRekening), 3)
  nCount = 1
  For n = 1 To Len(cRekening)
    If Mid(cRekening, n, 1) = "." Then
      nCount = nCount + 1
    End If
  Next
  Level = nCount
End Function

Function GetNull(Value, Optional Default As Variant = 0)
  GetNull = IIf(IsNull(Value) Or IsEmpty(Value), Default, Value)
End Function

Function GetFormLevel(cFormName, nLevel, Optional cmdAdd As Object, Optional cmdEdit As Object, Optional cmdDelete As Object) As String
Dim cStatus As String

  Set dbData = objData.Browse(GetDSN, "FormLevel", "Status", "Nama", sisAssign, cFormName, " and UserLevel = " & nLevel)
  If dbData.RecordCount > 0 Then
    cStatus = dbData!status
  Else
    cStatus = "111"
  End If
  cStatus = IIf(nUserLevel = 0, "111", cStatus)
  GetFormLevel = cStatus

  On Error Resume Next
  cmdAdd.Enabled = Val(left(cStatus, 1))
  cmdEdit.Enabled = Val(Mid(cStatus, 2, 1))
  cmdDelete.Enabled = Val(Mid(cStatus, 3, 1))
End Function

Function TabIndex(obj As Object, n As Single)
  obj.TabIndex = n
  n = n + 1
End Function

Function Mod25(ByVal nJumlah As Double)
Dim cJumlah As String

  nJumlah = Round(nJumlah)
  cJumlah = Right(Str(nJumlah), 2)
  Mod25 = nJumlah - Val(cJumlah)
  If Val(cJumlah) = 25 Or Val(cJumlah) = 0 Then
    Mod25 = nJumlah
  ElseIf Val(cJumlah) < 25 Then
    Mod25 = Mod25 + 25
  ElseIf Val(cJumlah) > 25 Then
    Mod25 = Mod25 + 50
  End If
End Function

Function Mod50(ByVal nJumlah As Double)
Dim cJumlah As String

  nJumlah = Round(nJumlah)
  cJumlah = Right(Str(nJumlah), 2)
  Mod50 = nJumlah - Val(cJumlah)
  If Val(cJumlah) = 50 Or Val(cJumlah) = 0 Then
    Mod50 = nJumlah
  ElseIf Val(cJumlah) < 50 Then
    Mod50 = Mod50 + 50
  ElseIf Val(cJumlah) > 50 Then
    Mod50 = Mod50 + 100
  End If
End Function

Function Mod1000(ByVal nJumlah As Double) As Double
Dim nMod As Double
Dim cJumlah As String
Dim lNegative As Boolean
Dim a As Double

  a = 1000
  lNegative = nJumlah < 0
  nJumlah = Round(Abs(nJumlah), 0)
  nMod = Int(Devide(nJumlah, a))
  nMod = nJumlah - (nMod * a)
  If nMod >= 500 Then
    nJumlah = nJumlah + 500
  End If
  cJumlah = nJumlah
  cJumlah = Padl(cJumlah, Len(cJumlah) + 3, "0")
  cJumlah = left(cJumlah, Len(cJumlah) - 3)
  Mod1000 = Val(cJumlah)
  If lNegative Then
    Mod1000 = -Mod1000
  End If
End Function

'Untuk Mendapatkan Baki Debet
Function GetBakiDebet(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String, ByVal nPlafond As Double, ByVal dTgl As Date) As Double
Dim dbBK As New ADODB.Recordset

  Set dbBK = obj.Browse(GetDSN, "Angsuran", "Tgl,Sum(Pokok) as Pokok", "Rekening", sisAssign, cRekening, " And Tgl <= '" & Format(dTgl, "yyyy-mm-dd") & "' Group by Rekening ")
  GetBakiDebet = nPlafond
  If Not dbBK.eof Then
    GetBakiDebet = nPlafond - GetNull(dbBK!pokok)
  End If
End Function

Function GetBilyetDeposito(ByVal obj As CodeSuiteLibrary.data, ByVal Cabang As String) As String
Dim cBilyet As String
Dim cNoBilyet As String
Dim dbBilyet As New ADODB.Recordset

  Set dbBilyet = obj.Browse(GetDSN, "NomorBilyet", "Max(NomorBilyet) NomorBilyet", "NomorBilyet", sisPrefix, Cabang)
  cBilyet = "1"
  If Not dbBilyet.eof Then
    cBilyet = Str(Val(Right(GetNull(dbBilyet!NomorBilyet), 8)) + 1)
  End If
  cNoBilyet = Cabang & Padl(Trim(cBilyet), 8, "0")
  GetBilyetDeposito = cNoBilyet
End Function

Sub SetButton(cmdSimpan As Object, cmdKeluar As Object, cmdAdd As Object, _
              cmdEdit As Object, cmdHapus As Object, nPos, lPar As Boolean, _
              Optional cmdAktivasi As Object)
  On Error Resume Next
  cmdSimpan.Enabled = lPar
  cmdAdd.Enabled = Not lPar
  cmdEdit.Enabled = Not lPar
  cmdHapus.Enabled = Not lPar
  cmdAktivasi.Visible = nUserLevel = 0
  If lPar Then
    Set cmdKeluar.Picture = aMainmenu.pcCancel.Picture
    cmdKeluar.Caption = "      &Cancel "
  Else
    Set cmdKeluar.Picture = aMainmenu.pcExit.Picture
    cmdKeluar.Caption = "      &Exit"
    
    Select Case nPos
      Case 1
        cmdAdd.SetFocus
      Case 2
        cmdEdit.SetFocus
      Case 3
        cmdHapus.SetFocus
    End Select
    nPos = 0
  End If
  
  If Not lPar Then
    GetFormLevel cmdSimpan.Parent.name, nUserLevel, cmdAdd, cmdEdit, cmdHapus
  End If
End Sub

Function GetPicture(ByVal cPath As String) As String
  On Error GoTo salah
  If Dir(cPath) <> "" Then
    GetPicture = cPath
  Else
    GetPicture = ""
  End If
  Exit Function
salah:
  GetPicture = ""
End Function

Function KasTeller() As Boolean
  KasTeller = Trim(cKasTeller) <> ""
  If Trim(cKasTeller) = "" Then
    MsgBox "Kode Kas Teller Tidak Ada, Anda Tidak Bisa Menjalakan Modul Ini" + vbCrLf + _
           "Hubungi Suppervisor Anda Untuk Melakukan Setup Kas Teller", vbExclamation + vbOKCancel
  End If
End Function

Function Min(n, a) As Double
  Min = IIf(n < a, n, a)
End Function

Function Max(n, a)
  Max = IIf(n < a, a, n)
End Function

Sub GetMinMax(ByVal cTable As String, vaValue, Optional ByVal cField As String = "Kode")
  Set dbData = objData.Browse(GetDSN, cTable, "Min(" & cField & ") as Min, Max(" & cField & ") as Max")
  vaValue(0).Text = ""
  vaValue(1).Text = ""
  If Not dbData.eof Then
    vaValue(0).Text = GetNull(dbData!Min, "")
    vaValue(1).Text = GetNull(dbData!Max, "")
  End If
End Sub

Function Devide(ByVal a As Double, ByVal b As Double) As Double
  If a = 0 Or b = 0 Then
    Devide = 0
  Else
    Devide = a / b
  End If
End Function

Function DevideMod(ByVal a As Double, ByVal b As Double) As Double
  If a = 0 Or b = 0 Then
    DevideMod = 0
  Else
    DevideMod = a \ b
  End If
End Function

Function GetFrekuensi(ByVal cNamaTabel As String, ByVal cCabang As String, ByVal cJenisProduk As SisJenisProduk, ByVal cGolonganProduk As String, ByVal cNoUrutRegister As String)
Dim cPrefix As String
Dim cNomorFrekuensi As String
Dim cNol As String

  cPrefix = cCabang & "." & cGolonganProduk & "." & cNoUrutRegister
  cNomorFrekuensi = "1"
  Select Case cJenisProduk
'    Case Is = 1, 2, 3
'      Set dbData = objData.Browse(GetDSN, cNamaTabel, "Rekening", "Rekening", sisPrefix, cPrefix)
'      If Not dbData.eof Then
'        cNomorFrekuensi = dbData.RecordCount + 1
'      End If
'    Case Is = 4, 5, 6
'      Set dbData = objData.Browse(GetDSN, cNamaTabel, "Rekening", "Rekening", sisPrefix, cPrefix)
'      If Not dbData.eof Then
'        cNomorFrekuensi = dbData.RecordCount + 1
'      End If
'    Case Is = 7, 8, 9
'      Set dbData = objData.Browse(GetDSN, cNamaTabel, "Rekening", "Rekening", sisPrefix, cPrefix)
'      If Not dbData.eof Then
'        cNomorFrekuensi = Str(dbData.RecordCount + 1)
'      End If

    Case Is = 1, 2, 3
      Set dbData = objData.Browse(GetDSN, cNamaTabel, "max(Rekening) as frekuensi", "Rekening", sisPrefix, cPrefix)
      If Not dbData.eof Then
        cNomorFrekuensi = Right(GetNull(dbData!frekuensi), 2) + 1
      End If
    Case Is = 4, 5, 6
      Set dbData = objData.Browse(GetDSN, cNamaTabel, "Rekening", "Rekening", sisPrefix, cPrefix)
      If Not dbData.eof Then
        cNomorFrekuensi = Right(GetNull(dbData!frekuensi), 2) + 1
      End If
    Case Is = 7, 8, 9
      Set dbData = objData.Browse(GetDSN, cNamaTabel, "Rekening", "Rekening", sisPrefix, cPrefix)
      If Not dbData.eof Then
        cNomorFrekuensi = Right(GetNull(dbData!frekuensi), 2) + 1
      End If

  End Select
  GetFrekuensi = IIf(Len(cNomorFrekuensi) = 1, "0" & cNomorFrekuensi, cNomorFrekuensi)
End Function

Function GetFrekuensiOld(ByVal cNamaTabel As String, ByVal cCabang As String, ByVal cJenisProduk As SisJenisProduk, ByVal cGolonganProduk As String, ByVal cNoUrutRegister As String)
Dim cPrefix As String
Dim cNomorFrekuensi As String
Dim cNol As String

  cPrefix = cCabang & "." & cGolonganProduk & "." & cNoUrutRegister
  cNomorFrekuensi = "1"
  Select Case cJenisProduk
    Case Is = 1, 2, 3
      Set dbData = objData.Browse(GetDSN, cNamaTabel, "Rekening", "Rekening", sisPrefix, cPrefix)
      If Not dbData.eof Then
        cNomorFrekuensi = dbData.RecordCount + 1
      End If
    Case Is = 4, 5, 6
      Set dbData = objData.Browse(GetDSN, cNamaTabel, "Rekening", "Rekening", sisPrefix, cPrefix)
      If Not dbData.eof Then
        cNomorFrekuensi = dbData.RecordCount + 1
      End If
    Case Is = 7, 8, 9
      Set dbData = objData.Browse(GetDSN, cNamaTabel, "Rekening", "Rekening", sisPrefix, cPrefix)
      If Not dbData.eof Then
        cNomorFrekuensi = Str(dbData.RecordCount + 1)
      End If
  End Select
  GetFrekuensiOld = IIf(Len(cNomorFrekuensi) = 1, "0" & cNomorFrekuensi, cNomorFrekuensi)
End Function

Function SetNomorRekening(ByVal cCabang As String, ByVal cGolongan As String, ByVal cUrut As String, ByVal cFrekuensi As String)
  SetNomorRekening = cCabang & "." & cGolongan & "." & cUrut & "." & cFrekuensi
End Function

Sub GetGambar(IMagePoto, ImageTTD, cPathPhoto As String, cPathTTD As String)
On Error Resume Next
  IMagePoto.Picture = LoadPicture(GetPicture(cPathPhoto))
  ImageTTD.Picture = LoadPicture(GetPicture(cPathTTD))
End Sub

Function GetSaldoTab(ByVal obj As CodeSuiteLibrary.data, ByVal cRekening As String, ByVal dAkhir As Date) As Double
Dim cTgl As String
Dim nJumlah As Double
Dim cWhere As String
Dim dbSaldo As New ADODB.Recordset
  
  cWhere = " Tgl <= '" & Format(dAkhir, "yyyy-MM-dd") & "' and Rekening = '" & cRekening & "'"
  nJumlah = 0
  Set dbSaldo = obj.Browse(GetDSN, "MutasiTabungan", "DK,Jumlah", , , , cWhere)
  If Not dbSaldo.eof Then
    Do While Not dbSaldo.eof
      nJumlah = nJumlah + IIf(dbSaldo!DK = "K", dbSaldo!Jumlah, -dbSaldo!Jumlah)
      dbSaldo.MoveNext
    Loop
  End If
  GetSaldoTab = nJumlah
End Function

Function GetPajak(ByVal obj As CodeSuiteLibrary.data, ByVal nSaldoTabungan As Double, ByVal nBunga As Double, ByVal cGolonganTabungan As String) As Double
  Set dbData = obj.Browse(GetDSN, "Golongantabungan", "SaldoMinimumKenaPajak,PajakBunga", "Kode", sisAssign, cGolonganTabungan)
  If dbData.RecordCount > 0 Then
    If nSaldoTabungan >= GetNull(dbData!SaldoMinimumKenaPajak) Then
      GetPajak = Round(nBunga * GetNull(dbData!pajakbunga) / 100, 2)
    End If
  End If
  GetPajak = Round(GetPajak, 2)
End Function

Function GetKasTeller(ByVal cusername As String) As String
Dim dbKas As New ADODB.Recordset
Dim ob As New CodeSuiteLibrary.data

  GetKasTeller = ""
  Set dbKas = ob.Browse(GetDSN, "Username", "kasTeller", "userName", sisAssign, cusername)
  If Not dbKas.eof Then
    GetKasTeller = GetNull(dbKas!KasTeller, "")
  End If
End Function

Function IsInPeriod(ByVal dTgl As Date) As Boolean
Dim db As New ADODB.Recordset
Dim obj As New CodeSuiteLibrary.data
Dim dAwal As Date
Dim dAkhir As Date
Dim lNull As Boolean

  lNull = False
  IsInPeriod = True
  Set db = obj.Browse(GetDSN, "Periode", "Min(Awal) as Awal,Max(Akhir) as Akhir", "Status", sisAssign, "0")
  If db.RecordCount > 0 Then
    If IsNull(db!Awal) Or IsNull(db!akhir) Then
      lNull = True
    End If
  Else
    lNull = True
  End If
  
  If lNull Then
    MsgBox "Tanggal Periode Akuntansi Belum di Setup, Transaksi Tidak Bisa Dilanjutkan" & Chr(13) & "Lakukan Setup Periode terlebih dahulu", vbExclamation
    IsInPeriod = False
    Exit Function
  Else
    dAwal = GetNull(db!Awal)
    dAkhir = GetNull(db!akhir)
  End If
  
  If Not (dTgl >= dAwal And dTgl <= dAkhir) Then
    MsgBox "Periode Transaksi sudah tutup, Transaksi Tidak bisa dilanjutkan", vbExclamation
    IsInPeriod = False
  End If
End Function

Function GetRptMaster(ByVal cTabel As String, cJudulLaporan As String, ByVal cNamaForm As String) As XArrayDB
Dim vaArray As New XArrayDB
Dim cJudul As String

  Set dbData = objData.Browse(GetDSN, cTabel, "Kode,Keterangan", , , , , "Kode")
  If Not dbData.eof Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  End If
  With FrmRPT
    cJudul = aCfg(msNama, "")
    .AddPageHeader cJudul, tdbHalignCenter, , , , dbArial, 12, True
    .AddPageHeader cJudulLaporan, tdbHalignCenter, , , True, dbArial, 12, True

    .AddTableHeader "Kode", , , , 7, , , , , , , , , , , , , 5
    .AddTableHeader "Keterangan"

    .AddTableBody
    .AddTableBody

    .Preview vaArray, True
  End With
End Function

Function GetSukuBungaDeposito(ByVal nPlafond As Double, ByVal nSukuBunga As Double) As Double
  GetSukuBungaDeposito = Round((nSukuBunga / 12) / 100 * nPlafond)
End Function

Function GetLastNomorRegister(ByVal cCabang As String, Optional ByVal cFaktur As String, Optional ByVal lUpdate As Boolean = False) As String
Dim db As New ADODB.Recordset
Dim obj As New CodeSuiteLibrary.data
Dim cChar As String
Dim cNomor As String
Dim nCount As Double
Dim cKode As String

  cNomor = 1
  cChar = cCabang
  cKode = cChar
  
  If Trim(cFaktur = "") Then
    If lUpdate Then
      
      obj.Add GetDSN, "nomorregister", Array("Kode"), Array(cKode)
      Set db = obj.SQL(GetDSN, "Select Last_Insert_id() as Total")
      
      ' Untuk Menghemat Ukuran Table Hapus jika Nomor ID < Nomor yang aktif
      If db.RecordCount > 0 Then
        obj.Delete GetDSN, "nomorregister", "Kode", sisAssign, cKode, " and id < " & GetNull(db!Total)
      End If
      
      nCount = 0
    Else
      Set db = obj.Browse(GetDSN, "nomorregister", "Max(ID) as Total", "Kode", sisAssign, cKode)
      nCount = 1
    End If
    If db.RecordCount > 0 Then
      cNomor = GetNull(db!Total, 0) + nCount
    End If
  Else
    cNomor = Trim(cFaktur)
  End If
  cNomor = Padl(cNomor, 6, "0")
  GetLastNomorRegister = cNomor
End Function
