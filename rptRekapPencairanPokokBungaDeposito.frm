VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form rptRekapPencairanPokokBungaDeposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REKAPITULASI PENCAIRAN POKOK BUNGA DEPOSITO"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4095
   Begin BiSAButtonProject.BiSAButton cmdHapus 
      Height          =   750
      Left            =   1860
      TabIndex        =   1
      Top             =   105
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1323
      Caption         =   "Hapus BB"
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
   Begin BiSAButtonProject.BiSAButton cmdPreview 
      Height          =   750
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1323
      Caption         =   "Preview"
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
Attribute VB_Name = "rptRekapPencairanPokokBungaDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim vaArray As New XArrayDB

Private Sub cmdHapus_Click()
Dim db As New ADODB.Recordset

  Set dbData = objData.Browse(GetDSN, "MutasiDeposito m", "m.Faktur,m.Rekening", "d.Status", sisDifference, "1", , , Array("Left Join Deposito d on d.Rekening = m.Rekening"))
  If Not dbData.eof Then
    'hapus data buku besar
    Do While Not dbData.eof
      objData.Delete GetDSN, "BukuBesar", "Faktur", sisAssign, GetNull(dbData!Faktur)
      dbData.MoveNext
    Loop
  End If
  
  Set db = objData.Browse(GetDSN, "MutasiDeposito m", "m.Faktur,m.Rekening", "d.Status", sisDifference, "1", , , Array("Left Join Deposito d on d.Rekening = m.Rekening"))
  If Not db.eof Then
    Do While Not db.eof
      objData.Delete GetDSN, "Mutasideposito", "Faktur", sisAssign, GetNull(db!Faktur)
      db.MoveNext
    Loop
  End If
  objData.Edit GetDSN, "Deposito", "Status='0'", Array("NominalDeposito"), Array(0)
  MsgBox "Selesai"
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub GetData()
Dim n As Single

  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.Browse(GetDSN, "mutasideposito m", "m.Rekening,r.Nama,m.Jumlah,m.Tgl,m.KodeMutasi,d.Status,d.SistemARO", , , , , "m.Tgl", Array("Left Join Deposito d on d.Rekening = m.Rekening", "Left Join RegisterNasabah r on r.Kode = d.Kode"))
  If Not dbData.eof Then
    Do While Not dbData.eof
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Rekening, "")
      vaArray(n, 1) = GetNull(dbData!nama, "")
      vaArray(n, 2) = GetNull(dbData!Jumlah, "")
      vaArray(n, 3) = GetNull(dbData!Tgl)
      vaArray(n, 4) = GetNamaMutasi(GetNull(dbData!KodeMutasi))
      vaArray(n, 5) = (GetNamaStatus(GetNull(dbData!status, "")) & "/" & GetNamaARO(GetNull(dbData!SistemARO)))
      dbData.MoveNext
    Loop
    GetRpt
  Else
    MsgBox "Maaf, data tidak ada..", vbInformation, Me.Caption
  End If
End Sub

Private Function GetNamaMutasi(ByVal KodeMutasi As String) As String
'  trPembukaan = 1
'  trPencairanPokok = 2
'  trPencairanBunga = 3
'  trPinalti = 4
'  trMaterai = 5
  Select Case KodeMutasi
    Case 1
      GetNamaMutasi = "Pembukaan"
    Case 2
      GetNamaMutasi = "Pencairan Pokok"
    Case 3
      GetNamaMutasi = "Pencairan Bunga"
    Case 4
      GetNamaMutasi = "Pinalti"
    Case 5
      GetNamaMutasi = "Materai"
  End Select
End Function

Private Function GetNamaStatus(ByVal KodeStatus As String) As String
  Select Case KodeStatus
    Case 0
      GetNamaStatus = "OPEN"
    Case 1
      GetNamaStatus = "CLOSE"
  End Select
End Function

Private Function GetNamaARO(ByVal KodeARO As String) As String
  Select Case KodeARO
    Case "Y"
      GetNamaARO = "ARO"
    Case "T"
      GetNamaARO = "NON ARO"
  End Select
End Function

Private Sub GetRpt()
    With FrmRPT
    .AddPageHeader "Mutasi Pokok/Bunga", tdbHalignCenter, , , , dbArial, 12, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "REKENING", , , , 15
    .AddTableHeader "NAMA"
    .AddTableHeader "JUMLAH", , , , 15
    .AddTableHeader "TGL", , , , 12
    .AddTableHeader "KODE MUTASI", , , , 15
    .AddTableHeader "STATUS", , , , 15
    
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody , tdbHalignCenter
    .AddTableBody
    .AddTableBody
    
    .Preview vaArray, True
  End With
End Sub
