VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgAUtoJurnal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KONFIGURASI REKENING AUTO JURNAL"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   8925
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1545
      Left            =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2725
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSATextBoxProject.BiSABrowse cRekeningLaba 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   105
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "REKENING LABA"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningLaba 
         Height          =   330
         Left            =   4650
         TabIndex        =   3
         Top             =   105
         Width           =   3915
         _ExtentX        =   6906
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
      Begin BiSATextBoxProject.BiSABrowse cRekKas 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   450
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "REKENING KAS INDUK"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekKas 
         Height          =   330
         Left            =   4650
         TabIndex        =   5
         Top             =   450
         Width           =   3915
         _ExtentX        =   6906
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
      Begin BiSATextBoxProject.BiSABrowse cRekPB 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   795
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "REKENING PB"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekPB 
         Height          =   330
         Left            =   4650
         TabIndex        =   7
         Top             =   795
         Width           =   3915
         _ExtentX        =   6906
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
      Begin BiSATextBoxProject.BiSABrowse cKodeTransaksiPB 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   1140
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "KODE TRANSAKSI PB"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaKodeTransaksiPB 
         Height          =   330
         Left            =   4635
         TabIndex        =   10
         Top             =   1140
         Width           =   3915
         _ExtentX        =   6906
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   1530
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1138
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   6600
         TabIndex        =   0
         Top             =   120
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
         Picture         =   "cfgAUtoJurnal.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   7680
         TabIndex        =   1
         Top             =   120
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
         Picture         =   "cfgAUtoJurnal.frx":0416
      End
      Begin VB.Label Label1 
         Caption         =   "F2 = SIMPAN"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   1815
      End
   End
End
Attribute VB_Name = "cfgAUtoJurnal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub Pick(cKD, cNM, cTableName)
  Set dbData = objData.Browse(GetDSN, cTableName, "Kode,Keterangan", "Kode", sisContent, cKD.Text)
  cKD.Text = cKD.Browse(dbData)
  If Not dbData.eof Then
    cNM.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cKodeTransaksiPB_ButtonClick()
  Pick cKodeTransaksiPB, cNamaKodeTransaksiPB, "KodeTransaksi"
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  

  UpdCfg msKodelaba, cRekeningLaba.Text, objData
  UpdCfg msKodeKasInduk, cRekKas.Text, objData
  UpdCfg msKodePemindahBukuan, cRekPB.Text, objData
  UpdCfg msKodeTransaksiPB, cKodeTransaksiPB.Text, objData
  
  MsgBox "Data telah tersimpan", vbInformation
End Sub

Private Sub cRekeningLaba_ButtonClick()
  Pick cRekeningLaba, cNamaRekeningLaba, "Rekening"
End Sub

Private Sub cRekeningLaba_Validate(Cancel As Boolean)
  cRekeningLaba_ButtonClick
End Sub

Private Sub cRekKas_ButtonClick()
  Pick cRekKas, cNamaRekKas, "Rekening"
End Sub

Private Sub cRekKas_Validate(Cancel As Boolean)
  cRekKas_ButtonClick
End Sub

Private Sub cRekPB_ButtonClick()
  Pick cRekPB, cNamaRekPB, "Rekening"
End Sub

Private Sub cRekPB_Validate(Cancel As Boolean)
  cRekPB_ButtonClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    cmdSimpan_Click
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  
  TabIndex cRekeningLaba, n
  TabIndex cRekKas, n
  TabIndex cRekPB, n
  TabIndex cKodeTransaksiPB, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n

  cRekeningLaba.Text = aCfg(msKodelaba, "")
  cRekKas.Text = aCfg(msKodeKasInduk, "")
  cRekPB.Text = aCfg(msKodePemindahBukuan, "")
  cKodeTransaksiPB.Text = aCfg(msKodeTransaksiPB, "")
  
  cNamaRekeningLaba.Text = GetNamaRekening(cRekeningLaba.Text)
  cNamaRekKas.Text = GetNamaRekening(cRekKas.Text)
  cNamaRekPB.Text = GetNamaRekening(cRekPB.Text)
  cNamaKodeTransaksiPB.Text = GetNamaKodeTransaksi(cKodeTransaksiPB.Text)
End Sub

Private Function GetNamaKodeTransaksi(ByVal KodeTransaksi As String)
  GetNamaKodeTransaksi = ""
  Set dbData = objData.Browse(GetDSN, "KodeTransaksi", "Kode,Keterangan", "Kode", sisAssign, KodeTransaksi)
  If Not dbData.eof Then
    GetNamaKodeTransaksi = GetNull(dbData!Keterangan)
  End If
End Function

Private Function GetNamaRekening(cRekening As String) As String
  GetNamaRekening = ""
  Set dbData = objData.Browse(GetDSN, "Rekening", "Keterangan", "Kode", sisAssign, cRekening)
  If Not dbData.eof Then
    GetNamaRekening = GetNull(dbData!Keterangan, "")
  End If
End Function
