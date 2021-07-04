Attribute VB_Name = "Version"
'Option Explicit
'
'Dim objData As New BiSAMyDLL.data
'Dim dbData As New ADODB.Recordset
'
'Function CheckVersion()
'Dim nVersion As Double
'Dim nOldVersion As Double
'Dim nRec As Double
'
'  On Error Resume Next
'  nVersion = 200405021 '2004-05-02 1
'  nOldVersion = aCfg(msVersion, 0)
'
''  If nOldVersion < 20031218 Then
''    objdata.SQL GetDSN, "Update totsl set total = subtotal - discount1 - discount2"
''  End If
'  If nOldVersion < 200405021 Then
'    objData.SQL GetDSN, "ALTER TABLE `debitur` ADD `PlafondPRK` DOUBLE(16,2)  DEFAULT '0' AFTER `TabunganWajib`"
'    objData.SQL GetDSN, "ALTER TABLE `golongankredit` ADD `RekeningPiutangJasa` CHAR(20) NOT NULL AFTER `RekeningAngsuranBunga`"
'    objData.SQL GetDSN, "ALTER TABLE `angsuran` ADD `StatusPendapatan` CHAR(1)  DEFAULT '0' AFTER `Kas`"
'    objData.SQL GetDSN, "CREATE TABLE `Denda` (`ID` DOUBLE AUTO_INCREMENT, `Tgl` DATE DEFAULT '0000-00-00', `Status` CHAR (1) DEFAULT '0', PRIMARY KEY(`ID`)) "
'    objData.SQL GetDSN, "ALTER TABLE `denda` ADD `Rekening` CHAR(15)  AFTER `ID`"
'    objData.SQL GetDSN, "ALTER TABLE `denda` ADD `TglPosting` DATE DEFAULT '0000-00-00' AFTER `Status`"
'    objData.SQL GetDSN, "CREATE TABLE `mutasidenda` (`ID` DOUBLE AUTO_INCREMENT, `Rekening` CHAR (15), `Tgl` DATE DEFAULT '0000-00-00', `Debet` DOUBLE (16,2) DEFAULT '0', `Kredit` DOUBLE (16,2) DEFAULT '0', PRIMARY KEY(`ID`)) "
'  End If
'
'  nOldVersion = Max(nVersion, nOldVersion)
'  UpdCfg msVersion, nVersion
'End Function
