VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgSetupMenuTeller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETUP MENU TELLER"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   8925
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2143
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin VB.OptionButton optLock 
         Caption         =   "&Tidak"
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
         Index           =   1
         Left            =   3375
         TabIndex        =   6
         Top             =   630
         Width           =   885
      End
      Begin VB.OptionButton optLock 
         Caption         =   "&Ya"
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
         Index           =   0
         Left            =   2745
         TabIndex        =   5
         Top             =   630
         Value           =   -1  'True
         Width           =   645
      End
      Begin BiSATextBoxProject.BiSABrowse cProductDefault 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   240
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
         Caption         =   "PRODUCT DEFAULT"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaProduct 
         Height          =   330
         Left            =   4650
         TabIndex        =   1
         Top             =   240
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
      Begin VB.Label Label1 
         Caption         =   "LOCK?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   630
         Width           =   1305
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   1200
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   6600
         TabIndex        =   2
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
         Picture         =   "cfgSetupMenuTeller.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   7680
         TabIndex        =   3
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
         Picture         =   "cfgSetupMenuTeller.frx":0416
      End
   End
End
Attribute VB_Name = "cfgSetupMenuTeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New BiSAMyDLL.data

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  SaveRegistry reg_DefaultProduct, cProductDefault.Text
  If optLock(0).Value = True Then
    SaveRegistry reg_lock, "Y"
  ElseIf optLock(1).Value = True Then
    SaveRegistry reg_lock, "T"
  End If
  MsgBox "Data sudah tersimpan", vbInformation
  Unload Me
End Sub

Private Sub cProductDefault_ButtonClick()
Dim cSQL As String
  
  cSQL = "select kode,keterangan from golongantabungan"
  cSQL = cSQL & " union select kode,keterangan from golongankredit"
  cSQL = cSQL & " union select kode,keterangan from golongandeposito"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.eof Then
    cProductDefault.Text = cProductDefault.Browse(dbData)
    cProductDefault.Text = GetNull(dbData!kode, "")
    cNamaProduct.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cProductDefault_Validate(Cancel As Boolean)
  cProductDefault_ButtonClick
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  If GetRegistry(reg_lock) = "Y" Then
    optLock(0).Value = True
  Else
    optLock(1).Value = True
  End If
  TabIndex cProductDefault, n
  TabIndex cNamaProduct, n
  TabIndex optLock(0), n
  TabIndex optLock(1), n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  cProductDefault.Text = GetRegistry(reg_DefaultProduct)
End Sub

Private Sub optLock_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
