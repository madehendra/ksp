VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{9A5A311D-C750-11D4-8714-444553540000}#46.0#0"; "SisTrueTextBox.ocx"
Object = "{8164BC59-C899-11D4-8714-444553540000}#5.0#0"; "SisTLabel.ocx"
Object = "{32A289A9-C7B2-11D4-8714-444553540000}#4.1#0"; "SisTFrame.ocx"
Object = "{0235E0D6-D0F0-11D4-8750-00E04CAB774A}#21.1#0"; "SisTDate.ocx"
Begin VB.Form FrmPengesahan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Pengesahaan"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin SisTFrame.sisFrame sisFrame2 
      Height          =   570
      Left            =   -30
      Top             =   4965
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   1005
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   435
         Left            =   5460
         TabIndex        =   13
         Top             =   90
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   767
         Caption         =   ""
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
         Picture         =   "FrmPengesahan.frx":0000
      End
   End
   Begin SisTFrame.sisFrame sisFrame1 
      Height          =   4980
      Left            =   -30
      Top             =   -15
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   8784
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
      Begin SisTDate.SisDate dTgl 
         Height          =   330
         Left            =   720
         TabIndex        =   0
         Top             =   4485
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Caption         =   "Tgl Cetak"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTrueTextBox.sisTextBox cParaf 
         Height          =   330
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   465
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Paraf"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTLabel.SisLabel SisLabel1 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   75
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Pengesahaan I"
      End
      Begin SisTrueTextBox.sisTextBox cNama 
         Height          =   330
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   825
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Nama"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTrueTextBox.sisTextBox cJabatan 
         Height          =   330
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   1185
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Jabatan"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTLabel.SisLabel SisLabel1 
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   1575
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Pengesahaan II"
      End
      Begin SisTrueTextBox.sisTextBox cParaf 
         Height          =   330
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   1965
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Paraf"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTrueTextBox.sisTextBox cNama 
         Height          =   330
         Index           =   1
         Left            =   720
         TabIndex        =   7
         Top             =   2325
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Nama"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTrueTextBox.sisTextBox cJabatan 
         Height          =   330
         Index           =   1
         Left            =   720
         TabIndex        =   8
         Top             =   2685
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Jabatan"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTLabel.SisLabel SisLabel1 
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   3015
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Pengesahaan III"
      End
      Begin SisTrueTextBox.sisTextBox cParaf 
         Height          =   330
         Index           =   2
         Left            =   720
         TabIndex        =   10
         Top             =   3405
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Paraf"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTrueTextBox.sisTextBox cNama 
         Height          =   330
         Index           =   2
         Left            =   720
         TabIndex        =   11
         Top             =   3765
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Nama"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTrueTextBox.sisTextBox cJabatan 
         Height          =   330
         Index           =   2
         Left            =   720
         TabIndex        =   12
         Top             =   4125
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Jabatan"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmPengesahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objdata As New BISAMyDLL.data
Dim cFN As String

Function GetPengesahaan(cFormName As String)
  Unload Me
  Load Me
  cFN = cFormName
  dTgl.Value = Date
  Set dbData = objdata.Browse(GetDSN, "Pengesahan", , "Kode", sisAssign, cFormName)
  If dbData.RecordCount > 0 Then
    cParaf(0).Text = dbData!Paraf_1
    cNama(0).Text = dbData!Nama_1
    cJabatan(0).Text = dbData!Jabatan_1
    cParaf(1).Text = dbData!Paraf_2
    cNama(1).Text = dbData!Nama_2
    cJabatan(1).Text = dbData!Jabatan_2
    cParaf(2).Text = dbData!Paraf_3
    cNama(2).Text = dbData!Nama_3
    cJabatan(2).Text = dbData!Jabatan_3
  End If
   Me.Show vbModal
End Function

Private Sub cmdOK_Click()
Dim vaField, vaValue
  vaField = Array("Kode", "Paraf_1", "Nama_1", "Jabatan_1", "Paraf_2", "Nama_2", "Jabatan_2", "Paraf_3", "Nama_3", "Jabatan_3")
  vaValue = Array(cFN, cParaf(0).Text, cNama(0).Text, cJabatan(0).Text, cParaf(1).Text, cNama(1).Text, cJabatan(1).Text, cParaf(2).Text, cNama(2).Text, cJabatan(2).Text)
  
  objdata.Update GetDSN, "Pengesahan", "Kode = '" & cFN & "'", vaField, vaValue
  Me.Hide
End Sub

Private Sub Form_Load()
Dim n As Single

  TabIndex cParaf(0), n
  TabIndex cNama(0), n
  TabIndex cJabatan(0), n
  TabIndex cParaf(1), n
  TabIndex cNama(1), n
  TabIndex cJabatan(1), n
  TabIndex cParaf(2), n
  TabIndex cNama(2), n
  TabIndex cJabatan(2), n
  TabIndex dTgl, n
  TabIndex cmdOK, n
End Sub
