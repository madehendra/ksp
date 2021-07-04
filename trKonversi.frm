VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form trKonversi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konversi Data"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3840
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   570
      TabIndex        =   1
      Top             =   1320
      Width           =   1650
   End
   Begin BiSAButtonProject.BiSAButton OK 
      Height          =   465
      Left            =   915
      TabIndex        =   0
      Top             =   585
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   820
      Caption         =   "Konversi Anggota"
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
End
Attribute VB_Name = "trKonversi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaField
Dim vaValue

Private Sub Command1_Click()
  GetUpdateKode
End Sub

Private Sub OK_Click()
  Set dbData = objData.Browse(GetDSN, "anggota")
  vaField = Array("Kode", "TglRegister", "Nama", "Kelamin", "GolonganDarah", _
                "TempatLahir", "TglLahir", "StatusPerkawinan", _
                "KTP", "Agama", "Pekerjaan", "Wilayah", "Alamat", "Telepon", _
                "AlamatKantor", "KodePosKantor", "TeleponKantor", "FaxKantor", _
                "Path", "Path1", "TglKtp", "NPWP", "JenisAnggota", "kodealias", "kodeasli")
  
  If Not dbData.eof Then
    Do While Not dbData.eof
     vaValue = Array(GetKode, "2006-01-01", GetNull(dbData!nama), "L", "O", _
                "Denpasar", "1945-12-12", "1", _
                "-", "01", "09", "01", "Denpasar", "-", _
                "-", "-", "-", "-", _
                "", "", "2010-01-01", "-", "1", GetNull(dbData!Kode), GetKode)
      objData.Add GetDSN, "registernasabah", vaField, vaValue
      dbData.MoveNext
    Loop
  End If
End Sub

Private Function GetKode() As String
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "registernasabah", "count(kode) as nomor")
  If Not db.eof Then
    GetKode = Str(Val(GetNull(db!nomor, 0)) + 1)
  End If
End Function

Private Sub GetUpdateKode()
Dim db As New ADODB.Recordset
Dim cNomor As String
  
  cNomor = ""
  Set db = objData.Browse(GetDSN, "registernasabah", "kodeasli as cNomor")
  If Not db.eof Then
    Do While Not db.eof
      cNomor = Trim(Padl(Trim(GetNull(db!cNomor)), 6, "0"))
      objData.Edit GetDSN, "registernasabah", "kodeasli = '" & GetNull(db!cNomor) & "'", Array("kode"), Array("01." & cNomor)
      db.MoveNext
    Loop
  MsgBox "selesai"
  End If
End Sub
