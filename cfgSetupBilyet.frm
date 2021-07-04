VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form cfgSetupBilyet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6795
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5475
      Left            =   0
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   9657
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
      Begin BiSANumberBoxProject.BiSANumberBox nTinggi 
         Height          =   330
         Left            =   615
         TabIndex        =   2
         Top             =   135
         Width           =   2130
         _ExtentX        =   3757
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
         Caption         =   "HEIGHT"
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
      Begin BiSANumberBoxProject.BiSANumberBox nLebar 
         Height          =   330
         Left            =   3255
         TabIndex        =   3
         Top             =   135
         Width           =   2130
         _ExtentX        =   3757
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
         Caption         =   "WIDTH"
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
      Begin BiSANumberBoxProject.BiSANumberBox xRekening 
         Height          =   330
         Left            =   1440
         TabIndex        =   6
         Top             =   675
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yRekening 
         Height          =   330
         Left            =   3165
         TabIndex        =   7
         Top             =   675
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wRekening 
         Height          =   330
         Left            =   4830
         TabIndex        =   8
         Top             =   675
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xNama 
         Height          =   330
         Left            =   1440
         TabIndex        =   9
         Top             =   1035
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yNama 
         Height          =   330
         Left            =   3165
         TabIndex        =   10
         Top             =   1035
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wNama 
         Height          =   330
         Left            =   4830
         TabIndex        =   11
         Top             =   1035
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xAlamat 
         Height          =   330
         Left            =   1440
         TabIndex        =   12
         Top             =   1395
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yAlamat 
         Height          =   330
         Left            =   3165
         TabIndex        =   13
         Top             =   1395
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wAlamat 
         Height          =   330
         Left            =   4830
         TabIndex        =   14
         Top             =   1395
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xNominal 
         Height          =   330
         Left            =   1440
         TabIndex        =   15
         Top             =   1755
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yNominal 
         Height          =   330
         Left            =   3165
         TabIndex        =   16
         Top             =   1755
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wNominal 
         Height          =   330
         Left            =   4830
         TabIndex        =   17
         Top             =   1755
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xterbilang 
         Height          =   330
         Left            =   1440
         TabIndex        =   18
         Top             =   2100
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yterbilang 
         Height          =   330
         Left            =   3165
         TabIndex        =   19
         Top             =   2100
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wterbilang 
         Height          =   330
         Left            =   4830
         TabIndex        =   20
         Top             =   2100
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xLama 
         Height          =   330
         Left            =   1440
         TabIndex        =   21
         Top             =   2460
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yLama 
         Height          =   330
         Left            =   3165
         TabIndex        =   22
         Top             =   2460
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wlama 
         Height          =   330
         Left            =   4830
         TabIndex        =   23
         Top             =   2460
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xvaluta 
         Height          =   330
         Left            =   1440
         TabIndex        =   24
         Top             =   2820
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yValuta 
         Height          =   330
         Left            =   3165
         TabIndex        =   25
         Top             =   2820
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wvaluta 
         Height          =   330
         Left            =   4830
         TabIndex        =   26
         Top             =   2820
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xTempo 
         Height          =   330
         Left            =   1440
         TabIndex        =   27
         Top             =   3180
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yTempo 
         Height          =   330
         Left            =   3165
         TabIndex        =   28
         Top             =   3180
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wTempo 
         Height          =   330
         Left            =   4830
         TabIndex        =   29
         Top             =   3180
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xBunga 
         Height          =   330
         Left            =   1440
         TabIndex        =   30
         Top             =   3525
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yBunga 
         Height          =   330
         Left            =   3165
         TabIndex        =   31
         Top             =   3525
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wBunga 
         Height          =   330
         Left            =   4830
         TabIndex        =   32
         Top             =   3525
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xPimpinan 
         Height          =   330
         Left            =   1440
         TabIndex        =   33
         Top             =   3885
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yPimpinan 
         Height          =   330
         Left            =   3165
         TabIndex        =   34
         Top             =   3885
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wPimpinan 
         Height          =   330
         Left            =   4830
         TabIndex        =   35
         Top             =   3885
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xkasir 
         Height          =   330
         Left            =   1440
         TabIndex        =   36
         Top             =   4245
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox ykasir 
         Height          =   330
         Left            =   3165
         TabIndex        =   37
         Top             =   4245
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wKasir 
         Height          =   330
         Left            =   4830
         TabIndex        =   38
         Top             =   4245
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xTerbilangSB 
         Height          =   330
         Left            =   1440
         TabIndex        =   50
         Top             =   4605
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yTerbilangSB 
         Height          =   330
         Left            =   3165
         TabIndex        =   51
         Top             =   4605
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wTerbilangSB 
         Height          =   330
         Left            =   4830
         TabIndex        =   52
         Top             =   4605
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin BiSANumberBoxProject.BiSANumberBox xTglCetak 
         Height          =   330
         Left            =   1440
         TabIndex        =   53
         Top             =   4965
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "X"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox yTglCetak 
         Height          =   330
         Left            =   3165
         TabIndex        =   54
         Top             =   4965
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Y"
         CaptionWidth    =   200
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
      Begin BiSANumberBoxProject.BiSANumberBox wTglCetak 
         Height          =   330
         Left            =   4830
         TabIndex        =   55
         Top             =   4965
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "Width"
         CaptionWidth    =   600
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
      Begin VB.Label Label5 
         Caption         =   "Terbilang S. B"
         Height          =   240
         Index           =   12
         Left            =   135
         TabIndex        =   57
         Top             =   4665
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Tgl Cetak"
         Height          =   240
         Index           =   11
         Left            =   135
         TabIndex        =   56
         Top             =   4995
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Kasir"
         Height          =   240
         Index           =   10
         Left            =   135
         TabIndex        =   49
         Top             =   4275
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Pimpinan"
         Height          =   240
         Index           =   9
         Left            =   135
         TabIndex        =   48
         Top             =   3945
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Suku Bunga"
         Height          =   240
         Index           =   8
         Left            =   135
         TabIndex        =   47
         Top             =   3570
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Jatuh Tempo"
         Height          =   240
         Index           =   7
         Left            =   135
         TabIndex        =   46
         Top             =   3210
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Tgl Valuta"
         Height          =   240
         Index           =   6
         Left            =   135
         TabIndex        =   45
         Top             =   2865
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Jangka Waktu"
         Height          =   240
         Index           =   5
         Left            =   135
         TabIndex        =   44
         Top             =   2490
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Terbilang"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   43
         Top             =   2145
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Nominal"
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   42
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Alamat"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   41
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Nama"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   40
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "No. Rekening"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   39
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label Label4 
         Caption         =   "mm"
         Height          =   270
         Left            =   2790
         TabIndex        =   5
         Top             =   195
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "mm"
         Height          =   270
         Left            =   5430
         TabIndex        =   4
         Top             =   210
         Width           =   360
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   5460
      Width           =   6750
      _ExtentX        =   11906
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
         Left            =   4410
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
         Picture         =   "cfgSetupBilyet.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5490
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
         Picture         =   "cfgSetupBilyet.frx":0416
      End
   End
End
Attribute VB_Name = "cfgSetupBilyet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.data
Dim dbData As New ADODB.Recordset

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue

  
      vaField = Array("XRekening", "YRekening", "WRekening", "XNama", "YNama", "WNama", "XALamat", "YAlamat", "WAlamat", _
                    "XJumlah", "YJumlah", "WJumlah", "XTerbilang", "YTerbilang", "WTerbilang", "XLama", "YLama", "WLama", _
                    "XValuta", "YValuta", "WVAluta", "XTempo", "YTempo", "WTempo", "XBunga", "YBunga", "WBunga", _
                    "XDirut", "YDirut", "WDirut", "Xkasir", "Ykasir", "WKasir", _
                    "Tinggi", "Lebar", "xTerbilangSB", "yTerbilangSB", "wTerbilangSB", "xTglCetak", "yTglcetak", "wTglCetak")
      vaValue = Array(xRekening.Value, yRekening.Value, wRekening.Value, xNama.Value, yNama.Value, wNama.Value, xAlamat.Value, yAlamat.Value, wAlamat.Value, _
                    xNominal.Value, yNominal.Value, wNominal.Value, xterbilang.Value, yterbilang.Value, wterbilang.Value, xLama.Value, yLama.Value, wlama.Value, _
                    xvaluta.Value, yValuta.Value, wvaluta.Value, xTempo.Value, yTempo.Value, wTempo.Value, xBunga.Value, yBunga.Value, wBunga.Value, _
                    xPimpinan.Value, yPimpinan.Value, wPimpinan.Value, xkasir.Value, ykasir.Value, wKasir.Value, _
                    nTinggi.Value, nLebar.Value, xTerbilangSB.Value, yTerbilangSB.Value, wTerbilangSB.Value, xTglCetak.Value, yTglCetak.Value, wTglCetak.Value)
      objData.Delete GetDSN, "SetupBilyet", "XRekening", sisDifference, 10000
      objData.Add GetDSN, "SetupBilyet", vaField, vaValue
      
  
  
  MsgBox "Data telah tersimpan......", vbInformation
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  Me.Top = 0
  TabIndex nTinggi, n
  TabIndex nLebar, n
  
  TabIndex xRekening, n
  TabIndex yRekening, n
  TabIndex wRekening, n
  
  TabIndex xNama, n
  TabIndex yNama, n
  TabIndex wNama, n
  
  TabIndex xAlamat, n
  TabIndex yAlamat, n
  TabIndex wAlamat, n
  
  TabIndex xNominal, n
  TabIndex yNominal, n
  TabIndex wNominal, n
  
  TabIndex xterbilang, n
  TabIndex yterbilang, n
  TabIndex wterbilang, n
  
  TabIndex xLama, n
  TabIndex yLama, n
  TabIndex wlama, n
  
  TabIndex xvaluta, n
  TabIndex yValuta, n
  TabIndex wvaluta, n
  
  TabIndex xTempo, n
  TabIndex yTempo, n
  TabIndex wTempo, n
  
  TabIndex xBunga, n
  TabIndex yBunga, n
  TabIndex wBunga, n
  
  TabIndex xPimpinan, n
  TabIndex yPimpinan, n
  TabIndex wPimpinan, n
  
  TabIndex xkasir, n
  TabIndex ykasir, n
  TabIndex wKasir, n
  
  TabIndex xTerbilangSB, n
  TabIndex yTerbilangSB, n
  TabIndex wTerbilangSB, n
  
  TabIndex xTglCetak, n
  TabIndex yTglCetak, n
  TabIndex wTglCetak, n
  
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  GetData
End Sub

Private Sub GetData()
  Set dbData = objData.SQL(GetDSN, "Select * From SetupBilyet")
  With dbData
    xRekening.Value = !xRekening
    yRekening.Value = !yRekening
    wRekening.Value = !wRekening
    xNama.Value = !xNama
    yNama.Value = !yNama
    wNama.Value = !wNama
    xAlamat.Value = !xAlamat
    yAlamat.Value = !yAlamat
    wAlamat.Value = !wAlamat
    xNominal.Value = !XJumlah
    yNominal.Value = !YJumlah
    wNominal.Value = !WJumlah
    xterbilang.Value = !xterbilang
    yterbilang.Value = !yterbilang
    wterbilang.Value = !wterbilang
    xLama.Value = !xLama
    yLama.Value = !yLama
    wlama.Value = !wlama
    xvaluta.Value = !xvaluta
    yValuta.Value = !yValuta
    wvaluta.Value = !wvaluta
    xTempo.Value = !xTempo
    yTempo.Value = !yTempo
    wTempo.Value = !wTempo
    xBunga.Value = !xBunga
    yBunga.Value = !yBunga
    wBunga.Value = !wBunga
    xPimpinan.Value = !XDirut
    yPimpinan.Value = !YDirut
    wPimpinan.Value = !WDirut
    xkasir.Value = !xkasir
    ykasir.Value = !ykasir
    wKasir.Value = !wKasir
    nTinggi.Value = !Tinggi
    nLebar.Value = !Lebar
    xTerbilangSB.Value = !xTerbilangSB
    yTerbilangSB.Value = !yTerbilangSB
    wTerbilangSB.Value = !wTerbilangSB
    xTglCetak.Value = !xTglCetak
    yTglCetak.Value = !yTglCetak
    wTglCetak.Value = !wTglCetak
    
  End With
End Sub
