VERSION 5.00
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin BiSANumberBoxProject.BiSANumberBox BiSANumberBox1 
      Height          =   435
      Left            =   450
      TabIndex        =   3
      Top             =   1635
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   450
      TabIndex        =   2
      Top             =   915
      Width           =   1515
   End
   Begin BiSADateProject.BiSADate BiSADate2 
      Height          =   390
      Left            =   1935
      TabIndex        =   1
      Top             =   435
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   688
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
   Begin BiSADateProject.BiSADate BiSADate1 
      Height          =   390
      Left            =   450
      TabIndex        =   0
      Top             =   435
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   688
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
   Begin BiSANumberBoxProject.BiSANumberBox BiSANumberBox2 
      Height          =   435
      Left            =   1470
      TabIndex        =   4
      Top             =   1635
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
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
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim nM As Integer
Dim nD As Integer

  GetDiff BiSADate1.Value, BiSADate2.Value, nM, nD
  BiSANumberBox1.Value = nM
  BiSANumberBox2.Value = nD
End Sub
