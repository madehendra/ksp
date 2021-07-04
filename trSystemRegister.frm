VERSION 5.00
Begin VB.Form frmSystemRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update System Register"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   2625
   Begin VB.CommandButton cmdOK 
      Caption         =   "&PROSES"
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   2415
   End
End
Attribute VB_Name = "frmSystemRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset

Private Sub cmdOK_Click()
Dim nMax As Integer

Set dbData = objData.Browse(GetDSN, "registernasabah", "max(Kode) as Kode")
  If Not dbData.eof Then
    nMax = Val(Right(GetNull(dbData!Kode), Len(GetNull(dbData!Kode)) - 3))
    objData.Delete GetDSN, "NomorRegister", "Kode", sisAssign, "01"
    objData.Add GetDSN, "nomorregister", Array("Kode", "ID"), Array("01", nMax)
    MsgBox "Nomor Register Terupdate", vbInformation
  End If
End Sub

Private Sub Form_Load()
  CenterForm Me, True
End Sub
