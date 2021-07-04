Attribute VB_Name = "System"
Option Explicit
Public GetDSN As New ADODB.Connection


Public Enum vbRegistry
  reg_DSN = 0
  reg_UserLevel = 1
  reg_UserID = 2
  reg_UserName = 3
  reg_FullName = 4
  reg_Database = 5
  reg_IP = 6
  reg_ServerUID = 7
  reg_ServerPWD = 8
  reg_Wallpaper = 9
  reg_DefaultProduct = 10
  reg_lock = 11
End Enum

Public Enum SisFormatType
  Sis_yyyy_MM_dd = 0
  Sis_dd_MM_yyyy = 1
  Sis_BilRpPict2 = 2
  Sis_BilRpPict = 3
End Enum

Function SisFormat(ByVal Value, bFormat As SisFormatType, Optional ByVal cNegatifSeparator As String = "") As String
Dim vaFormat
  
  vaFormat = Array("yyyy-MM-dd", "dd-MM-yyyy", "###,###,###,###,###,##0.00", "###,###,###,###,###,##0", "dd-MM-YY")
  If Len(cNegatifSeparator) > 0 Then
    If Value >= 0 Then
      cNegatifSeparator = Space(Len(cNegatifSeparator))
    End If

    SisFormat = left(cNegatifSeparator, 1) & Format(Abs(Value), vaFormat(bFormat)) & Right(cNegatifSeparator, 1)
  Else
    SisFormat = Format(Value, vaFormat(bFormat))
  End If
End Function

Function GetDSNOld() As String
  GetDSN = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & GetRegistry(reg_DSN)
End Function

Function CheckData(cData, cMsg) As Boolean
  CheckData = True
  If Len(Trim(cData)) = 0 Or cData = 0 Then
    CheckData = False
    MsgBox (cMsg), , "ERROR"
  End If
End Function

Function Padl(Optional ByVal cCharacter As String = "", Optional ByVal nLen As Byte = 0, Optional ByVal cChar = " ") As String
Dim n As Byte, x As String
  x = ""
  If Len(cCharacter) < nLen Then
    For n = 1 To nLen - Len(cCharacter)
      x = cChar & x
    Next
    Padl = x & cCharacter
  Else
    Padl = Mid(cCharacter, 1, nLen)
  End If
End Function

Function Padr(Optional cCharacter As String = "", Optional nLen As Byte = 0, Optional cChar = " ") As String
Dim n As Byte, x As String
  cCharacter = left(cCharacter, nLen)
  x = ""
  If Len(cCharacter) < nLen Then
    For n = 1 To nLen - Len(cCharacter)
      x = cChar & x
    Next
    Padr = cCharacter + x
  Else
    Padr = Mid(cCharacter, 1, nLen)
  End If
End Function

Function Pad(cCharacter As String, Optional nLen As Byte = 0, Optional cChar = " ") As String
  Pad = Padr(cCharacter, nLen, cChar)
End Function

Function BOM(ByVal dDate As Date) As Date
  BOM = dDate - Day(dDate) + 1
End Function

Function BOY(ByVal dDate As Date) As Date
  BOY = DateSerial(Year(dDate), "01", "01")
End Function

Function EOM(ByVal dDate As Date) As Date
Dim n As Byte, OldDate As Date
  OldDate = dDate
  Do While Month(dDate) = Month(OldDate)
    dDate = dDate + 1
  Loop
  EOM = dDate - 1
End Function

Function GetMonth(nMonth As Single) As String
Dim vaMonth
  vaMonth = Array("Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember")
  GetMonth = vaMonth(nMonth - 1)
End Function

Function RAT(cChar, cString As String) As Single
Dim n As Single
  For n = Len(cString) To 1 Step -1
    If Mid(cString, n, 1) = cChar Then
      RAT = n
      Exit Function
    End If
  Next
End Function

Function CreateDSNOlkd(cDSN As String, ByVal cIPServer As String, Optional cDatabase As String = "Syariah", Optional cUser As String = "root", Optional cPwd As String = "", Optional cPort As String = "3307")

Dim cKey As String, x As Long, buffer As String * 255

  ' Ambil Posisi Directory System Window
  x = GetSystemDirectory(buffer, 255)
  buffer = left(buffer, x)

  ' Register DSN
  SetStringValue "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\ODBC Data Sources", cDSN, "MySQL ODBC 3.51 Driver"
  
  ' Configurasi DSN
  cKey = "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\" & cDSN
  
  DeleteKey cKey
  CreateKey cKey
  SetStringValue cKey, "Database", cDatabase
  SetStringValue cKey, "Description", ""
  'myodbc5a.dll
  SetStringValue cKey, "Driver", "myodbc5a.dll"
  'SetStringValue cKey, "Driver", Trim(Buffer) & "\myodbc3.dll"
  SetStringValue cKey, "Option", "3"
  SetStringValue cKey, "Password", cPwd
  SetStringValue cKey, "Port", cPort
  SetStringValue cKey, "Server", cIPServer
  SetStringValue cKey, "Stmt", ""
  SetStringValue cKey, "User", cUser
End Function


Function CreateDSNlama(cDSN As String, ByVal cIPServer As String, Optional cDatabase As String = "Syariah", Optional cUser As String = "root", Optional cPwd As String = "", Optional cPort As String = "", Optional cMYODBCPATH As String = "", Optional cMYODBCFile As String = "")

Dim cKey As String, x As Long, buffer As String * 255

  ' Ambil Posisi Directory System Window
  x = GetSystemDirectory(buffer, 255)
  buffer = left(buffer, x)

  ' Register DSN
  SetStringValue "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\ODBC Data Sources", cDSN, "MySQL ODBC Driver"
  
  ' Configurasi DSN
  cKey = "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\" & cDSN
  
  DeleteKey cKey
  CreateKey cKey
  SetStringValue cKey, "Database", cDatabase
  SetStringValue cKey, "Description", ""

  SetStringValue cKey, "Driver", cMYODBCFile
  SetStringValue cKey, "Option", "3"
  SetStringValue cKey, "Password", "FullMoon"
  SetStringValue cKey, "Port", cPort
  SetStringValue cKey, "Server", cIPServer
  SetStringValue cKey, "Stmt", ""
  SetStringValue cKey, "User", "kode"
  SetStringValue cKey, "Uid", cUser
  
End Function

Function CreateDSN(cDSN As String, ByVal cIPServer As String, Optional cDatabase As String = "Syariah", Optional cUser As String = "root", Optional cPwd As String = "", Optional cPort As String = "", Optional cMYODBCPATH As String = "", Optional cMYODBCFile As String = "")
Dim cKey As String, x As Long, buffer As String * 255

  ' Ambil Posisi Directory System Window
  x = GetSystemDirectory(buffer, 255)
  buffer = left(buffer, x)

  ' Register DSN
  SetStringValue "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\ODBC Data Sources", cDSN, "MySQL ODBC Driver"
  
  ' Configurasi DSN
  cKey = "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\" & cDSN
  
  DeleteKey cKey
  CreateKey cKey
  SetStringValue cKey, "Database", cDatabase
  SetStringValue cKey, "Description", ""
  
  
'  If Left(GetOsVersion, 1) > 5 Then
'    'tidak support
''    MsgBox ("Maaf OS tidak support")
'    SetStringValue cKey, "Driver", "C:\Program Files\MySQL\Connector ODBC 5.2\myodbc5.dll"
'  Else
'    'support
'    SetStringValue cKey, "Driver", Trim(buffer) & "\myodbc3.dll"
'
'  End If
  
'  SetStringValue cKey, "Driver", "C:\Program Files\MySQL\Connector ODBC 5.2\myodbc5a.dll"
  
'  SetStringValue cKey, "Driver", "C:\Program Files\MariaDB\MariaDB ODBC Driver\maodbc.dll"
'  SetStringValue cKey, "Driver", Trim(buffer) & "\myodbc5a.dll"
'  SetStringValue cKey, "Driver", App.Path & "\myodbc5a.dll"
'  SetStringValue cKey, "Driver", cMYODBCPATH & "\myodbc3.dll"
' sebelumnya tolong di set path dari path installasi myodbc nya

  SetStringValue cKey, "Driver", cMYODBCFile
  SetStringValue cKey, "Option", "3"
  SetStringValue cKey, "Password", cPwd
  SetStringValue cKey, "Port", cPort
  SetStringValue cKey, "Server", cIPServer
  SetStringValue cKey, "Stmt", ""
  SetStringValue cKey, "User", cUser
  SetStringValue cKey, "Uid", cUser
  
  'SetStringValue "HKEY_CURRENT_USER\Control Panel\International", "sShortDate", "dd-MM-yyyy"
End Function


Function Replicate(cString As String, nCount) As String
Dim n, cRetval As String
  For n = 1 To nCount
    cRetval = cRetval & cString
  Next
  Replicate = cRetval
End Function

Function GetOpt(Opt) As String
Dim n As Single, i As Single, lChar As Boolean
  For n = 0 To Opt.Count - 1
    If Opt(n).Value Then
      With Opt(n)
        For i = 1 To Len(.Caption)
          If lChar Then
            GetOpt = UCase(Mid(.Caption, i, 1))
            Exit Function
          End If
          If Mid(.Caption, i, 1) = "&" Then
            lChar = True
          End If
        Next
      End With
    End If
  Next
End Function

Sub SetOpt(Opt, cChar As String)
Dim n As Single, i As Single, lChar As Boolean
  Opt(0).Value = True
  For n = 0 To Opt.Count - 1
    With Opt(n)
      For i = 1 To Len(.Caption)
        If lChar Then
          lChar = False
          If UCase(Mid(.Caption, i, 1)) = UCase(cChar) Then
            Opt(n).Value = True
            Exit Sub
          End If
        End If
        
        If Mid(.Caption, i, 1) = "&" Then
          lChar = True
        End If
      Next
    End With
  Next
End Sub

Sub InitGrid(TDB As TDBGrid)
Dim nSplit As Integer
Dim nCol As Integer
Dim nBack As Double
Dim nFore As Double
Dim nBack1 As Double

  nBack = &HFFC0C0     'vbButtonFace
  nBack1 = vbButtonFace
  nFore = &H0&        'vbButtonText
  
  TDB.CaptionStyle.BackColor = nBack
  TDB.CaptionStyle.ForeColor = nFore
  TDB.CaptionStyle.Font.Bold = True
  TDB.DeadAreaBackColor = nBack1
    
  For nSplit = 0 To TDB.Splits.Count - 1
    TDB.Splits(nSplit).CaptionStyle.BackColor = nBack
    TDB.Splits(nSplit).CaptionStyle.ForeColor = nFore
    TDB.Splits(nSplit).CaptionStyle.Font.Bold = True
    
    TDB.Splits(nSplit).HeadBackColor = nBack
    TDB.Splits(nSplit).HeadForeColor = nFore
    TDB.Splits(nSplit).HeadFont.Bold = True
    
    TDB.Splits(nSplit).SelectedStyle.BackColor = vbHighlight
    TDB.Splits(nSplit).SelectedStyle.ForeColor = &H8000000E
    TDB.Splits(nSplit).MarqueeStyle = dbgHighlightCell
    
    For nCol = 0 To TDB.Splits(nSplit).Columns.Count - 1
      TDB.Splits(nSplit).Columns(nCol).HeadBackColor = nBack
      TDB.Splits(nSplit).Columns(nCol).HeadForeColor = nFore
      TDB.Splits(nSplit).Columns(nCol).HeadFont.Bold = True
      
      TDB.Splits(nSplit).Columns(nCol).FooterBackColor = nBack
      TDB.Splits(nSplit).Columns(nCol).FooterForeColor = nFore
    Next
  Next
  
  For nCol = 0 To TDB.Columns.Count - 1
    TDB.Columns(nCol).HeadBackColor = nBack
    TDB.Columns(nCol).HeadForeColor = nFore
    TDB.Columns(nCol).HeadFont.Bold = True
  Next
  TDB.HeadFont.Bold = True
End Sub

Function SNow() As String
  SNow = Format(Now, "yyyy-mm-dd hh:mm:ss")
End Function

Function Between(ByVal Value, ByVal Lower, ByVal Upper) As Boolean
  Between = Value >= Lower And Value <= Upper
End Function

Function GetData(ByVal cData As String, ByVal cKey As String, ByVal cDefault As String) As String
Dim n As Double
  cData = LCase(cData)
  cKey = LCase(cKey)
  GetData = cDefault
  n = InStr(1, cData, cKey)
  If n <> 0 Then
    cData = Replace(cData, cKey, "")
    GetData = cData
  End If
End Function

Sub GetIPNumber(ByRef cIPNumber As String, ByRef cDatabase As String, ByRef cDSN As String, ByRef cPort As String, ByRef cKey As String, Optional ByRef cModePelunasanPiutang As String = "")

Dim cFile As String
Dim n As Double
Dim cData As String

  cFile = App.Path & "\config.ini"
  If Dir(cFile) <> "" Then
    Open cFile For Input Shared As #1
    Do While Not eof(1)
      Line Input #1, cData
      cData = Replace(cData, " ", "")
      
      cIPNumber = GetData(cData, "IP=", cIPNumber)
      cDatabase = GetData(cData, "DATABASE=", cDatabase)
      cPort = GetData(cData, "PORT=", cPort)
      cDSN = GetData(cData, "DSN=", cDSN)
      cKey = GetData(cData, "KEY=", cKey)
      cModePelunasanPiutang = GetData(cData, "LUNASPIUTANG=", cModePelunasanPiutang)
    Loop
    Close #1
  End If
  
  'simpan pada registry
  SaveRegistry reg_DSN, cDSN
  SaveRegistry reg_Database, cDatabase
  SaveRegistry reg_IP, cIPNumber
'  SaveRegistry reg_ModePelunasanPiutang, cModePelunasanPiutang
  
  If Trim(cIPNumber) = "" Then
    cIPNumber = "LocalHost"
  End If
  If Trim(cDatabase) = "" Then
    cDatabase = "RENT"
  End If
  If Trim(cDSN) = "" Then
    cDSN = "RENT"
  End If
End Sub

Function SaveRegistry(par As vbRegistry, cValue)
  SaveSetting "USPD", App.EXEName, "Reg" & par, cValue
End Function

Function GetRegistry(par As vbRegistry, Optional cDefault = "")
  GetRegistry = GetSetting("USPD", App.EXEName, "Reg" & par, "")
End Function

Function GetAppDescription() As String
  GetAppDescription = App.Title & "." & App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Sub CetakValidasiTabungan(ByVal cFaktur As String, ByVal dTanggal As Date, ByVal dDateTime As String, ByVal cNomorRekening As String, ByVal cNamaNasabah As String, _
                                 ByVal cKodeTransaksi As String, ByVal cNamaKodeTransaksi As String, ByVal cDK As String, ByVal nJumlahMutasi As Double)
Dim NoRekening As String
Dim DK As String
Dim nTopMargin As Integer
Dim n As Integer
Dim nLeftMargin As Integer

  If cDK = "D" Then
    DK = "DB Rp    " & Format(nJumlahMutasi, "##,###,###,##0.00")
  Else
    DK = "CR Rp    " & Format(nJumlahMutasi, "##,###,###,##0.00")
  End If
  
  nTopMargin = aCfg(msTopValidasiTabungan)
  If aCfg(msLeftValidasiTabungan) = 0 Or aCfg(msTopValidasiTabungan) = 0 Then
    MsgBox "Setting margin untuk buku tabungan belum di setup!!", vbInformation
  Else
'    With aMainmenu.IO1
'      .Open "LPT1:", ""
'      .WriteString Chr(27) & Chr(15) & vbCrLf
'      For n = 1 To nTopMargin
'        .WriteString vbCrLf
'      Next
'      .WriteString Space(aCfg(msLeftValidasiTabungan)) & cFaktur & "   " & dTanggal & "   " & dDateTime & "   " & cNomorRekening & "   " & cNamaNasabah & vbCrLf
'      .WriteString Space(aCfg(msLeftValidasiTabungan)) & cusername & "   " & cKodeTransaksi & "   " & cNamaKodeTransaksi & "   " & DK & vbCrLf
'      .WriteString vbFormFeed
'      .Close
'    End With
  End If
End Sub
