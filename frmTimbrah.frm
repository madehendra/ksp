VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmTimbrah 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   525
      Left            =   4410
      TabIndex        =   5
      Top             =   2295
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   926
      Caption         =   "Label1"
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
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   750
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      Caption         =   "debitur"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   720
      Left            =   420
      TabIndex        =   3
      Top             =   4410
      Width           =   2190
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1080
      Left            =   450
      TabIndex        =   2
      Top             =   2580
      Width           =   2595
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   780
      Left            =   240
      TabIndex        =   1
      Top             =   1170
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   540
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   2445
   End
End
Attribute VB_Name = "frmTimbrah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub BiSAButton1_Click()
Dim cSQL As String
Dim vaField
Dim vaValue

vaField = Array("Rekening", "Wilayah", "Kode", "GolonganKredit", _
              "NoSPK", "SukuBunga", "Tgl", _
              "Plafond", "Lama", "AO", "Administrasi", _
              "Materai", "JatuhTempo", _
              "TotalBunga", "NoPengajuan", _
              "CaraAngsuran", "PeriodeBungaMenurun", _
              "MinimalPeriode", "Provisi", "Notaris", "BiayaLainLain", _
              "wajibpokok", "KonpensasiTelat", "DendaTelatBayar", "UserName", "DateTime", "caraperhitungan")

cSQL = " select deb.* from debitur15 deb"
cSQL = cSQL & " left join debitur d on d.rekening = deb.rekening"
cSQL = cSQL & " where d.rekening is null;"
Set dbData = objData.SQL(GetDSN, cSQL)
If Not dbData.eof Then
  Do While Not dbData.eof
                  
    vaValue = Array(GetNull(dbData!Rekening), GetNull(dbData!Wilayah), GetNull(dbData!Kode), GetNull(dbData!GolonganKredit), _
                  GetNull(dbData!NoSPK), GetNull(dbData!SukuBunga), GetNull(dbData!Tgl), _
                  GetNull(dbData!plafond), GetNull(dbData!Lama), GetNull(dbData!AO), GetNull(dbData!Administrasi), _
                  GetNull(dbData!Materai), GetNull(dbData!JatuhTempo), _
                  GetNull(dbData!totalBunga), GetNull(dbData!NoPengajuan), _
                  GetNull(dbData!CaraAngsuran), GetNull(dbData!PeriodeBungaMenurun), _
                  GetNull(dbData!MinimalPeriode), GetNull(dbData!Provisi), GetNull(dbData!Notaris), GetNull(dbData!BiayaLainLain), _
                  GetNull(dbData!wajibpokok), GetNull(dbData!KonpensasiTelat), GetNull(dbData!DendaTelatBayar), GetNull(dbData!UserName), GetNull(dbData!DateTime), GetNull(dbData!caraperhitungan))
      objData.Add GetDSN, "debitur", vaField, vaValue
    dbData.MoveNext
  Loop
  MsgBox "Selesai"
End If
End Sub

Private Sub BiSAButton2_Click()
Dim cSQL As String
Dim vaField
Dim vaValue

vaField = Array("Rekening", "Wilayah", "Kode", "GolonganKredit", _
              "NoSPK", "SukuBunga", "Tgl", _
              "Plafond", "Lama", "AO", "Administrasi", _
              "Materai", "JatuhTempo", _
              "TotalBunga", "NoPengajuan", _
              "CaraAngsuran", "PeriodeBungaMenurun", _
              "MinimalPeriode", "Provisi", "Notaris", "BiayaLainLain", _
              "wajibpokok", "KonpensasiTelat", "DendaTelatBayar", "UserName", "DateTime", "caraperhitungan")

cSQL = "select * from debitur15 where rekening = '01.K2.001201.01'"

Set dbData = objData.SQL(GetDSN, cSQL)
If Not dbData.eof Then
  Do While Not dbData.eof
                  
    vaValue = Array(GetNull(dbData!Rekening), GetNull(dbData!Wilayah), GetNull(dbData!Kode), GetNull(dbData!GolonganKredit), _
                  GetNull(dbData!NoSPK), GetNull(dbData!SukuBunga), GetNull(dbData!Tgl), _
                  GetNull(dbData!plafond), GetNull(dbData!Lama), GetNull(dbData!AO), GetNull(dbData!Administrasi), _
                  GetNull(dbData!Materai), GetNull(dbData!JatuhTempo), _
                  GetNull(dbData!totalBunga), GetNull(dbData!NoPengajuan), _
                  GetNull(dbData!CaraAngsuran), GetNull(dbData!PeriodeBungaMenurun), _
                  GetNull(dbData!MinimalPeriode), GetNull(dbData!Provisi), GetNull(dbData!Notaris), GetNull(dbData!BiayaLainLain), _
                  GetNull(dbData!wajibpokok), GetNull(dbData!KonpensasiTelat), GetNull(dbData!DendaTelatBayar), GetNull(dbData!UserName), GetNull(dbData!DateTime), GetNull(dbData!caraperhitungan))
      objData.Add GetDSN, "debitur", vaField, vaValue
    dbData.MoveNext
  Loop
  MsgBox "Selesai"
End If
End Sub

Private Sub Command1_Click()
Dim cSQL As String
cSQL = "select d.rekening,d.nopengajuan from debitur d"
cSQL = cSQL & " left join saldoawalkredit s on s.rekening = d.rekening"
cSQL = cSQL & " Where s.Rekening Is Null"

  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    Do While Not dbData.eof
      'hapus di table pengajuan
      objData.Delete GetDSN, "pengajuankredit", "kode", sisAssign, GetNull(dbData!NoPengajuan)
      'hapus di table realisasi
      objData.Delete GetDSN, "pencairankredit", "rekening", sisAssign, GetNull(dbData!Rekening)
      'hapus di table debitur
      objData.Delete GetDSN, "debitur", "rekening", sisAssign, GetNull(dbData!Rekening)
      dbData.MoveNext
    Loop
  End If
  MsgBox "Uh. selesai :)"
End Sub

Private Sub Command2_Click()
Dim cSQL As String
  cSQL = ""
  cSQL = "select s.rekening,d.plafond,s.pokok from saldoawalkredit s"
  cSQL = cSQL & " left join debitur d on d.rekening = s.rekening"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    Do While Not dbData.eof
      'update di table saldoawalkredit
      'update di table angsuran
      objData.Update GetDSN, "saldoawalkredit", "rekening = '" & GetNull(dbData!Rekening) & "'", Array("pokok"), Array(GetNull(dbData!plafond) - GetNull(dbData!pokok))
      objData.Update GetDSN, "angsuran", "rekening = '" & GetNull(dbData!Rekening) & "'", Array("pokok"), Array(GetNull(dbData!plafond) - GetNull(dbData!pokok))
      dbData.MoveNext
    Loop
  End If
  MsgBox "Selesai"
End Sub

Private Sub Command3_Click()
Dim cSQL As String

  cSQL = ""
  cSQL = "select s.rekening,d.plafond,s.pokok from saldoawalkredit s"
  cSQL = cSQL & " left join debitur d on d.rekening = s.rekening"
  cSQL = cSQL & " where s.pokok = 0"
  
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    Do While Not dbData.eof
      'delete di angsuran dan saldoawalkredit
      objData.Delete GetDSN, "saldoawalkredit", "rekening", sisAssign, GetNull(dbData!Rekening)
      dbData.MoveNext
    Loop
  End If
  MsgBox "Uh Selesai Juga"
  
End Sub

Private Sub Command4_Click()
Dim cSQL As String

  cSQL = ""
  cSQL = "select a.rekening from agunan a"
  cSQL = cSQL & " left join debitur d on d.rekening = a.rekening"
  cSQL = cSQL & " where d.rekening is null"
  
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    Do While Not dbData.eof
      'delete di angsuran dan saldoawalkredit
      objData.Delete GetDSN, "agunan", "rekening", sisAssign, GetNull(dbData!Rekening)
      dbData.MoveNext
    Loop
  End If
  MsgBox "Uh Selesai Juga"
End Sub
