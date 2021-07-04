VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form trReksadanaCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reksadana Calculator"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton cmdReset 
      Height          =   360
      Left            =   4380
      TabIndex        =   7
      Top             =   3495
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   635
      Caption         =   "Reset"
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
   Begin BiSANumberBoxProject.BiSANumberBox nSukuBunga 
      Height          =   330
      Left            =   375
      TabIndex        =   0
      Top             =   210
      Width           =   4110
      _ExtentX        =   7250
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
      Caption         =   "% pertumbuhan NAB/Bulan"
      CaptionWidth    =   3000
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
   Begin BiSANumberBoxProject.BiSANumberBox nBulan 
      Height          =   330
      Left            =   375
      TabIndex        =   1
      Top             =   555
      Width           =   4110
      _ExtentX        =   7250
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
      Caption         =   "Lamanya Investasi (Bulan)"
      CaptionWidth    =   3000
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
   Begin BiSANumberBoxProject.BiSANumberBox nNominalInvestasi 
      Height          =   330
      Left            =   360
      TabIndex        =   2
      Top             =   900
      Width           =   5415
      _ExtentX        =   9551
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
      Caption         =   "Nominal Investasi/ Bulan"
      CaptionWidth    =   3000
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
   Begin BiSANumberBoxProject.BiSANumberBox nModal 
      Height          =   330
      Left            =   360
      TabIndex        =   3
      Top             =   1665
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   582
      Enabled         =   0   'False
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
      Caption         =   "Modal Investasi"
      CaptionWidth    =   3000
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
   Begin BiSANumberBoxProject.BiSANumberBox nHasilInvestasi 
      Height          =   330
      Left            =   360
      TabIndex        =   4
      Top             =   2010
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   582
      Enabled         =   0   'False
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
      Caption         =   "Hasil Investasi"
      CaptionWidth    =   3000
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
   Begin BiSANumberBoxProject.BiSANumberBox nProfit 
      Height          =   330
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   582
      Enabled         =   0   'False
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
      Caption         =   "Margin (Rupiah)"
      CaptionWidth    =   3000
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
   Begin BiSANumberBoxProject.BiSANumberBox nPersenMargin 
      Height          =   330
      Left            =   360
      TabIndex        =   6
      Top             =   2865
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      Enabled         =   0   'False
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
      Caption         =   "Margin (%)"
      CaptionWidth    =   3000
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
   Begin BiSAButtonProject.BiSAButton cmdOK 
      Height          =   360
      Left            =   5550
      TabIndex        =   8
      Top             =   3495
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   635
      Caption         =   "OK"
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
Attribute VB_Name = "trReksadanaCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Dim n As Integer
Dim nSaldo As Double
Dim nModalAwal As Double

  nSaldo = 0
  nModalAwal = 0
  For n = 1 To nBulan.Value
    nSaldo = nSaldo + ((nSukuBunga.Value * n * nNominalInvestasi.Value / 100) + nNominalInvestasi.Value)
    nModalAwal = nModalAwal + nNominalInvestasi.Value
  Next n
  nModal.Value = nModalAwal
  nHasilInvestasi.Value = nSaldo
  nProfit.Value = nHasilInvestasi.Value - nModal.Value
  nPersenMargin.Value = Devide((nHasilInvestasi.Value - nModal.Value), nHasilInvestasi.Value) * 100
End Sub

Private Sub cmdReset_Click()
  initvalue
End Sub

Private Sub Form_Load()
Dim n As Single

  initvalue
  CenterForm Me
  TabIndex nSukuBunga, n
  TabIndex nBulan, n
  TabIndex nNominalInvestasi, n
  TabIndex nModal, n
  TabIndex nHasilInvestasi, n
  TabIndex nProfit, n
  TabIndex nPersenMargin, n
  TabIndex cmdOK, n
  TabIndex cmdReset, n
End Sub

Private Sub nReset_Click()
  initvalue
End Sub

Private Sub initvalue()
  nSukuBunga.Default
  nBulan.Default
  nNominalInvestasi.Default
  nModal.Default
  nHasilInvestasi.Default
  nProfit.Default
  nPersenMargin.Default
End Sub

Private Sub nBulan_Validate(Cancel As Boolean)
  cmdOK_Click
End Sub

Private Sub nNominalInvestasi_Validate(Cancel As Boolean)
  cmdOK_Click
End Sub

Private Sub nSukuBunga_Validate(Cancel As Boolean)
  cmdOK_Click
End Sub
