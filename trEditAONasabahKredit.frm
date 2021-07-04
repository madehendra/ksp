VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form trEditAONasabahKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit ARO Nasabah Kredit"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5925
   Begin BiSATextBoxProject.BiSABrowse cAO 
      Height          =   330
      Left            =   60
      TabIndex        =   0
      Top             =   810
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      Button          =   -1  'True
      Caption         =   "AO"
      CaptionWidth    =   1500
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BiSATextBoxProject.BiSATextBox cNamaAO 
      Height          =   330
      Left            =   2850
      TabIndex        =   1
      Top             =   795
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   582
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      BackColor       =   12632256
      Enabled         =   0   'False
      Appearance      =   0
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BiSATextBoxProject.BiSABrowse cKode 
      Height          =   330
      Left            =   60
      TabIndex        =   2
      Top             =   105
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   582
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      Button          =   -1  'True
      Caption         =   "Kode"
      CaptionWidth    =   1500
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BiSATextBoxProject.BiSABrowse cNama 
      Height          =   330
      Left            =   60
      TabIndex        =   3
      Top             =   450
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   582
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      Button          =   -1  'True
      Caption         =   "Nama"
      CaptionWidth    =   1500
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BiSAButtonProject.BiSAButton cmdSimpan 
      Height          =   435
      Left            =   3630
      TabIndex        =   4
      Top             =   1320
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      Caption         =   "    &Save"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      BackColor       =   -2147483633
      Picture         =   "trEditAONasabahKredit.frx":0000
   End
   Begin BiSAButtonProject.BiSAButton cmdKeluar 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4710
      TabIndex        =   5
      Top             =   1320
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   767
      Caption         =   "     &Exit"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      BackColor       =   -2147483633
      Picture         =   "trEditAONasabahKredit.frx":0416
   End
End
Attribute VB_Name = "trEditAONasabahKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB

Private Sub cAO_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "ao", "kode,nama,alamat")
  If Not dbData.eof Then
    cAO.Text = cKode.Browse(dbData)
    cNamaAO.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cKode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "registernasabah", "kode,nama,alamat")
  If Not dbData.eof Then
    cKode.Text = cKode.Browse(dbData)
    cNama.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  Set dbData = objData.Browse(GetDSN, "debitur", , "kode", sisAssign, cKode.Text)
  If Not dbData.eof Then
    MsgBox "rekening ini memiliki " & dbData.RecordCount & " rekening kredit"
    Do While Not dbData.eof
      objData.Edit GetDSN, "debitur", "rekening = '" & GetNull(dbData!Rekening) & "'", Array("ao"), Array(cAO.Text)
      dbData.MoveNext
    Loop
    MsgBox "OK, data sudah disimpan"
  Else
    MsgBox "maaf nasabah ini tidak memiliki satu rekening kredit, data tidak bisa disimpan"
  End If
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "registernasabah", "nama,kode,alamat")
  If Not dbData.eof Then
    cNama.Text = cKode.Browse(dbData)
    cKode.Text = GetNull(dbData!Kode)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me, True
  cKode.Default
  cNama.Default
  cAO.Default
  cNamaAO.Default
  
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex cAO, n
  
End Sub
