VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgCetakanNeracaPembuatPemeriksa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi Cetakan Neraca"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5805
   Begin VB.CheckBox Check1 
      Caption         =   "Tampilkan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   330
      TabIndex        =   7
      Top             =   165
      Width           =   1380
   End
   Begin BiSAButtonProject.BiSAButton cmdSimpan 
      Height          =   450
      Left            =   4455
      TabIndex        =   6
      Top             =   3105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      Caption         =   "Simpan"
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
   Begin BiSATextBoxProject.BiSATextBox cNamaPemeriksa 
      Height          =   330
      Left            =   630
      TabIndex        =   1
      Top             =   960
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
      MaxLength       =   30
      Appearance      =   0
      Caption         =   "Nama"
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
   Begin BiSATextBoxProject.BiSATextBox cJabatanPemeriksa 
      Height          =   330
      Left            =   630
      TabIndex        =   2
      Top             =   1320
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
      MaxLength       =   40
      Appearance      =   0
      Caption         =   "Jabatan"
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
   Begin BiSATextBoxProject.BiSATextBox cNamaPembuat 
      Height          =   330
      Left            =   630
      TabIndex        =   4
      Top             =   2160
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
      MaxLength       =   30
      Appearance      =   0
      Caption         =   "Nama"
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
   Begin BiSATextBoxProject.BiSATextBox cJabatanPembuat 
      Height          =   330
      Left            =   630
      TabIndex        =   5
      Top             =   2535
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
      MaxLength       =   40
      Appearance      =   0
      Caption         =   "Jabatan"
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
   Begin VB.Label Label2 
      Caption         =   "Pembuat"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   345
      TabIndex        =   3
      Top             =   1815
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Pemeriksa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   330
      TabIndex        =   0
      Top             =   645
      Width           =   1140
   End
End
Attribute VB_Name = "cfgCetakanNeracaPembuatPemeriksa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub cmdSimpan_Click()
  

  UpdCfg msNamaPemeriksaNeraca, cNamaPemeriksa.Text, objData
  UpdCfg msJabatanPemeriksaNeraca, cJabatanPemeriksa.Text, objData
  UpdCfg msNamaPembuatNeraca, cNamaPembuat.Text, objData
  UpdCfg msJabatanPembuatNeraca, cJabatanPembuat.Text, objData
  UpdCfg msOptTampilkanFootNoteNeraca, Check1.Value, objData
  
  
  MsgBox "Data telah tersimpan", vbInformation
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  
  TabIndex cNamaPemeriksa, n
  TabIndex cJabatanPemeriksa, n
  TabIndex cNamaPembuat, n
  TabIndex cJabatanPembuat, n
  TabIndex cmdSimpan, n
  
  cNamaPemeriksa.Text = aCfg(msNamaPemeriksaNeraca)
  cJabatanPemeriksa.Text = aCfg(msJabatanPemeriksaNeraca)
  cNamaPembuat.Text = aCfg(msNamaPembuatNeraca)
  cJabatanPembuat.Text = aCfg(msJabatanPembuatNeraca)
End Sub
