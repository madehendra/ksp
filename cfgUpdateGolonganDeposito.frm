VERSION 5.00
Begin VB.Form cfgUpdateGolonganDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Golongan Deposito"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   2955
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   465
      TabIndex        =   0
      Top             =   315
      Width           =   1590
   End
End
Attribute VB_Name = "cfgUpdateGolonganDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub Command1_Click()
  Set dbData = objData.Browse(GetDSN, "deposito", "GolonganDeposito,Lama")
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.eof
    FrmPB.RunPB
      objData.Edit GetDSN, "deposito", "GolonganDeposito='" & (dbData!GolonganDeposito) & "'", Array("Lama"), Array(GetLama(objData, GetNull(dbData!GolonganDeposito)))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Function GetLama(ByVal obj As CodeSuiteLibrary.data, ByVal GolonganDeposito) As Double
Dim db As New ADODB.Recordset
  GetLama = 0
  Set db = obj.Browse(GetDSN, "golonganDeposito", "Lama,Kode", "Kode", sisAssign, GolonganDeposito)
  If Not dbData.eof Then
    GetLama = GetNull(db!Lama)
  End If
End Function
