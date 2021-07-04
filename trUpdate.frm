VERSION 5.00
Begin VB.Form trUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   2715
   Begin VB.CommandButton cmdUpdatePajak 
      Caption         =   "Update Pajak"
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   2520
   End
End
Attribute VB_Name = "trUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdatePajak_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

cSQL = "select d.nominaldeposito,g.kode as kodegolongan,g.pajakbunga,d.sukubunga,m.faktur,m.rekening,m.jumlah,m.pajak from mutasideposito m"
cSQL = cSQL & " left join deposito d on d.rekening = m.rekening"
cSQL = cSQL & " left join golongandeposito g on g.kode = d.golongandeposito"
cSQL = cSQL & " where d.status <> 1 and m.kodemutasi = 3 and d.nominaldeposito > 7500000"

Set db = objData.SQL(GetDSN, cSQL)
If Not db.eof Then
  FrmPB.InitPB db.RecordCount
  Do While Not db.eof
    FrmPB.RunPB
      objData.Update GetDSN, "mutasideposito", "faktur = '" & GetNull(db!Faktur) & "' and rekening = '" & GetNull(db!Rekening) & "'", Array("pajak"), Array(GetPajak(GetSukuBungaDeposito(GetNull(db!nominaldeposito), GetNull(db!SukuBunga)), GetNull(db!pajakbunga)))
    db.MoveNext
  Loop
  FrmPB.EndPB
End If
End Sub

Private Function GetPajak(nBunga As Double, nPajakBunga As Double) As Double
GetPajak = 0

  GetPajak = Round(nBunga * nPajakBunga / 100)
End Function

Function GetSukuBungaDeposito(ByVal nPlafond As Double, ByVal nSukuBunga As Double) As Double
  GetSukuBungaDeposito = Round((nSukuBunga / 12) / 100 * nPlafond)
End Function

Private Sub Form_Load()
  CenterForm Me
End Sub
