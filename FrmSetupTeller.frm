VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form FrmSetupTeller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETUP TELLER"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1140
      Left            =   0
      Top             =   0
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   2011
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
      Begin VB.OptionButton Opt 
         Caption         =   "Tidak"
         Height          =   270
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   630
         Width           =   1380
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Ya"
         Height          =   270
         Index           =   0
         Left            =   1485
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   675
      End
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   2490
         TabIndex        =   3
         Top             =   165
         Width           =   4125
         _ExtentX        =   7276
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Left            =   375
         TabIndex        =   2
         Top             =   165
         Width           =   2085
         _ExtentX        =   3678
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
         BackColor       =   15456971
         ForeColor       =   -2147483635
         GetPicture      =   1
         Button          =   -1  'True
         Caption         =   "Default"
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
      Begin VB.Label Label1 
         Caption         =   "Lock ?"
         Height          =   270
         Left            =   420
         TabIndex        =   6
         Top             =   615
         Width           =   990
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   1125
      Width           =   6765
      _ExtentX        =   11933
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
         Left            =   4470
         TabIndex        =   0
         Top             =   105
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
         Picture         =   "FrmSetupTeller.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5550
         TabIndex        =   1
         Top             =   105
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
         Picture         =   "FrmSetupTeller.frx":0416
      End
   End
End
Attribute VB_Name = "FrmSetupTeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Produk", "Kode,Keterangan", "Kode", sisContent, cGolongan.Text, " Group By Kode")
  cGolongan.Text = cGolongan.Browse(dbData)
  If Not dbData.eof Then
    cNama.Text = GetNull(dbData!Keterangan)
  End If
End Sub

Private Sub cGolongan_Validate(Cancel As Boolean)
  If cGolongan.LastKey = 13 Or cGolongan.LastKey = 40 Then
    cGolongan_ButtonClick
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  UpdCfg msDefaultTeller, cGolongan.Text, objData
  UpdCfg msLockTeller, IIf(Opt(0).Value = True, "1", "2"), objData
  MsgBox "Data telah tersimpan..", vbInformation
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  Opt(0).Value = True
  
  TabIndex cGolongan, n
  TabIndex Opt(0), n
  TabIndex Opt(1), n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  cGolongan.Text = aCfg(msDefaultTeller)
  If aCfg(msLockTeller) = "1" Then
    Opt(0).Value = True
  Else
    Opt(1).Value = True
  End If
  Set dbData = objData.Browse(GetDSN, "Produk", "Keterangan", "Kode", sisAssign, cGolongan.Text)
  If Not dbData.eof Then
    cNama.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub Opt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub
